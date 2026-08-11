# -*- coding: utf-8 -*-
"""
Automatizacion de carga de fotos - Refacturaciones Aguas Andinas
==================================================================

Que hace:
    1. Lee las fotos de una carpeta de mes (ej. "Junio 2026").
    2. Agrupa las fotos por ID de cliente (el numero antes de sufijos
       tipo (1), (2), -1, _1, " 1", etc).
    3. Para cada ID: busca el cliente en el sistema, revisa si YA tiene
       imagenes adjuntas (si las tiene, se asume que se subieron a mano
       y NO se toca ese ID), y si no tiene ninguna, adjunta las fotos
       nuevas (maximo 3, maximo 5MB cada una).
    4. Verifica contra la tabla real del sistema que la cantidad de
       filas despues coincide con lo que se intento subir.
    5. Deja un registro (procesados.csv) y un reporte legible al final.

IMPORTANTE - COSAS QUE HAY QUE AJUSTAR SI ALGO FALLA:
    Todos los "nombres exactos de ventana/boton" que la app usa estan
    centralizados en la seccion CONFIG mas abajo. Si algun paso falla,
    lo mas probable es que haya que ajustar un texto ahi (ver README.md).

Requiere (en la maquina donde se hace la prueba, ANTES de empaquetar
con PyInstaller):
    pip install pywinauto
"""

import csv
import os
import re
import sys
import time
import traceback
from dataclasses import dataclass, field
from datetime import datetime

# ---------------------------------------------------------------------------
# CONFIG - AJUSTAR AQUI SI ALGO NO CALZA CON LO QUE VE EN PANTALLA
# ---------------------------------------------------------------------------

# Carpeta raiz donde estan las subcarpetas de mes (Mayo 2026, Junio 2026, etc)
CARPETA_FOTOS_BASE = r"C:\Users\eramirez\Documents\AGUAS ANDINAS\REFACTURACIONES\FOTOS"

# Limites de negocio
MAX_ARCHIVOS_POR_ID = 3
MAX_TAMANO_MB = 5

# Nombre del archivo de registro (se crea junto al .exe/.py)
ARCHIVO_REGISTRO = "procesados.csv"
ARCHIVO_REPORTE = "reporte_ultima_corrida.txt"

# --- Textos / titulos de la app (para pywinauto) -------------------------
# Ajustar estos regex si la app conecta con un titulo distinto al esperado.
TITULO_VENTANA_PRINCIPAL_REGEX = r".*Consultas.*"
NOMBRE_CAMPO_NRO = "Nro:"
NOMBRE_BOTON_ADJUNTAR = "Adjuntar imagen"
NOMBRE_BOTON_VER_IMAGENES = "Ver imágen(es)"
TITULO_VENTANA_ADJUNTAR = "Adjuntar imagen"
NOMBRE_BOTON_EXPLORAR = "..."
NOMBRE_BOTON_AGREGAR = "Agregar"
NOMBRE_BOTON_CERRAR_POPUP = "Cerrar"

# Tiempo maximo de espera (segundos) para que aparezca cada ventana/control
TIMEOUT_ESPERA = 15
# Pausa corta entre acciones para que la app legacy alcance a repintar
PAUSA_CORTA = 0.4

# ---------------------------------------------------------------------------
# ESTRUCTURAS DE DATOS
# ---------------------------------------------------------------------------


@dataclass
class GrupoID:
    id_cliente: str
    archivos_validos: list = field(default_factory=list)   # rutas completas OK
    archivos_excedente: list = field(default_factory=list)  # pasaron de 3
    archivos_muy_grandes: list = field(default_factory=list)  # > 5MB


@dataclass
class ResultadoID:
    id_cliente: str
    estado: str          # OK / OMITIDO_YA_TENIA / NO_ENCONTRADO / VERIF_FALLIDA / ERROR / EXCEDENTE / MUY_GRANDE
    detalle: str = ""
    archivos: list = field(default_factory=list)


# ---------------------------------------------------------------------------
# PASO 1: ESCANEO Y AGRUPAMIENTO DE ARCHIVOS
# ---------------------------------------------------------------------------

# Reconoce sufijos tipo: (1), (2), -1, _1, " 1", "-1 " (con espacio antes del guion)
PATRON_SUFIJO = re.compile(
    r"""^
    (?P<id>\d+)                      # el ID: uno o mas digitos al inicio
    (?:                               # sufijo opcional:
        \s*\(\d+\)                    #   (1)  (2) ...
        |\s*-\s*\d+                   #   -1   - 1 ...
        |\s*_\s*\d+                   #   _1
    )?
    $""",
    re.VERBOSE,
)

EXTENSIONES_IMAGEN = {".jpg", ".jpeg", ".png", ".bmp", ".jpeg"}


def extraer_id_limpio(nombre_archivo_sin_ext: str) -> str | None:
    """Dado '522595(1)' o '312410 -1' devuelve '522595' / '312410'.
    Devuelve None si el nombre no calza con el patron esperado
    (ej. archivos con nombres no numericos, se dejan fuera y se loguean)."""
    m = PATRON_SUFIJO.match(nombre_archivo_sin_ext.strip())
    if not m:
        return None
    return m.group("id")


def escanear_carpeta(carpeta_mes: str):
    """Recorre la carpeta de un mes y devuelve:
        - dict {id_limpio: GrupoID}
        - lista de archivos que no se pudieron interpretar (nombre raro)
    """
    grupos: dict[str, GrupoID] = {}
    no_reconocidos: list[str] = []

    if not os.path.isdir(carpeta_mes):
        raise FileNotFoundError(f"No existe la carpeta: {carpeta_mes}")

    for nombre in sorted(os.listdir(carpeta_mes)):
        ruta_completa = os.path.join(carpeta_mes, nombre)
        if not os.path.isfile(ruta_completa):
            continue

        base, ext = os.path.splitext(nombre)
        if ext.lower() not in EXTENSIONES_IMAGEN:
            continue  # ignorar archivos que no son imagen

        id_limpio = extraer_id_limpio(base)
        if id_limpio is None:
            no_reconocidos.append(nombre)
            continue

        grupo = grupos.setdefault(id_limpio, GrupoID(id_cliente=id_limpio))

        tamano_mb = os.path.getsize(ruta_completa) / (1024 * 1024)
        if tamano_mb > MAX_TAMANO_MB:
            grupo.archivos_muy_grandes.append(ruta_completa)
            continue

        if len(grupo.archivos_validos) >= MAX_ARCHIVOS_POR_ID:
            grupo.archivos_excedente.append(ruta_completa)
            continue

        grupo.archivos_validos.append(ruta_completa)

    return grupos, no_reconocidos


def listar_meses_disponibles():
    """Lista las subcarpetas dentro de CARPETA_FOTOS_BASE (ej. 'Junio 2026')."""
    if not os.path.isdir(CARPETA_FOTOS_BASE):
        return []
    carpetas = [
        nombre for nombre in sorted(os.listdir(CARPETA_FOTOS_BASE))
        if os.path.isdir(os.path.join(CARPETA_FOTOS_BASE, nombre))
    ]
    return carpetas


# ---------------------------------------------------------------------------
# PASO 2: REGISTRO PERSISTENTE (procesados.csv)
# ---------------------------------------------------------------------------


def cargar_registro() -> set[str]:
    """Devuelve el set de 'id_cliente|archivo' ya marcados como OK
    en corridas anteriores, para no reintentarlos."""
    procesados = set()
    if os.path.exists(ARCHIVO_REGISTRO):
        with open(ARCHIVO_REGISTRO, newline="", encoding="utf-8-sig") as f:
            lector = csv.DictReader(f)
            for fila in lector:
                if fila.get("estado") == "OK":
                    clave = f"{fila['id_cliente']}|{os.path.basename(fila['archivo'])}"
                    procesados.add(clave)
    return procesados


def registrar_resultado(id_cliente: str, archivo: str, estado: str, detalle: str = ""):
    existe = os.path.exists(ARCHIVO_REGISTRO)
    with open(ARCHIVO_REGISTRO, "a", newline="", encoding="utf-8-sig") as f:
        escritor = csv.writer(f)
        if not existe:
            escritor.writerow(["fecha", "id_cliente", "archivo", "estado", "detalle"])
        escritor.writerow([
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            id_cliente,
            archivo,
            estado,
            detalle,
        ])


# ---------------------------------------------------------------------------
# PASO 3: AUTOMATIZACION DE LA VENTANA (pywinauto)
# ---------------------------------------------------------------------------
# Importado de forma perezosa para que el modulo de escaneo (pasos 1 y 2)
# se pueda probar en Linux/sin Windows sin que falle el import.

class AutomatizadorApp:
    def __init__(self):
        self.app = None
        self.ventana_principal = None

    def conectar(self):
        from pywinauto import Application

        # NOTA: esta app WinForms expone la ventana "Consultas [Opc:101]"
        # duplicada (aparecen 2 elementos con el mismo titulo, mismo PID
        # y misma clase - probablemente un frame MDI + la ventana hija
        # compartiendo el mismo texto). Por eso no podemos exigir que el
        # titulo sea unico; usamos found_index=0 para tomar la primera
        # coincidencia en vez de fallar con ElementAmbiguousError.
        self.app = Application(backend="win32").connect(
            title_re=TITULO_VENTANA_PRINCIPAL_REGEX,
            found_index=0,
            timeout=TIMEOUT_ESPERA,
        )
        self.ventana_principal = self.app.window(
            title_re=TITULO_VENTANA_PRINCIPAL_REGEX, found_index=0
        )
        self.ventana_principal.wait("exists ready", timeout=TIMEOUT_ESPERA)

    def buscar_cliente(self, id_cliente: str) -> bool:
        """Escribe el ID en el campo Nro y presiona Enter.
        Devuelve True si parece haber cargado un cliente (heuristica simple)."""
        campo_nro = self.ventana_principal.child_window(
            title=NOMBRE_CAMPO_NRO, control_type="Edit"
        )
        # A veces el label y el campo de edicion son controles distintos;
        # si esto falla, en pywinauto suele existir un Edit "hermano" al
        # lado del texto "Nro:". Ver README si esta linea da error.
        campo_nro.wait("ready", timeout=TIMEOUT_ESPERA)
        campo_nro.set_focus()
        campo_nro.set_edit_text("")
        time.sleep(PAUSA_CORTA)
        campo_nro.type_keys(id_cliente, with_spaces=True)
        campo_nro.type_keys("{ENTER}")
        time.sleep(PAUSA_CORTA * 2)

        # Heuristica: si el propio campo Nro sigue mostrando el ID que
        # escribimos (no se borro ni quedo en blanco por error de busqueda)
        # asumimos que cargo. Si tu sistema muestra algun mensaje de error
        # cuando el cliente no existe, avisame el texto exacto para
        # afinar esto.
        texto_actual = campo_nro.window_text().strip()
        return texto_actual != ""

    def contar_imagenes_adjuntas(self) -> int:
        """Abre 'Ver imagen(es)', cuenta filas de la tabla DataGridView,
        cierra la ventana y devuelve la cantidad."""
        boton_ver = self.ventana_principal.child_window(
            title=NOMBRE_BOTON_VER_IMAGENES, control_type="Button"
        )
        boton_ver.wait("ready", timeout=TIMEOUT_ESPERA)
        boton_ver.click()
        time.sleep(PAUSA_CORTA * 2)

        # La ventana que se abre no tiene titulo confirmado todavia.
        # Tomamos la ventana "top" mas reciente de la aplicacion.
        ventana_ver = self.app.top_window()
        ventana_ver.wait("exists ready", timeout=TIMEOUT_ESPERA)

        tabla = ventana_ver.child_window(class_name="DataGridView")
        tabla.wait("exists", timeout=TIMEOUT_ESPERA)

        try:
            filas = tabla.row_count()
        except Exception:
            # Fallback: contar via elementos hijos de tipo fila si
            # row_count() no esta disponible en esta version de control.
            filas = len(tabla.children(control_type="Row"))

        # Cerrar la ventana de verificacion (boton "Cerrar" generico o Alt+F4)
        try:
            ventana_ver.child_window(title=NOMBRE_BOTON_CERRAR_POPUP, control_type="Button").click()
        except Exception:
            ventana_ver.close()

        time.sleep(PAUSA_CORTA)
        return filas

    def adjuntar_archivo(self, ruta_archivo: str):
        """Ejecuta el flujo completo de adjuntar UN archivo:
        Adjuntar imagen -> ... -> escribir ruta -> Agregar -> Cerrar."""
        boton_adjuntar = self.ventana_principal.child_window(
            title=NOMBRE_BOTON_ADJUNTAR, control_type="Button"
        )
        boton_adjuntar.wait("ready", timeout=TIMEOUT_ESPERA)
        boton_adjuntar.click()
        time.sleep(PAUSA_CORTA * 2)

        ventana_adjuntar = self.app.window(title=TITULO_VENTANA_ADJUNTAR)
        ventana_adjuntar.wait("exists ready", timeout=TIMEOUT_ESPERA)

        boton_explorar = ventana_adjuntar.child_window(
            title=NOMBRE_BOTON_EXPLORAR, control_type="Button"
        )
        boton_explorar.wait("ready", timeout=TIMEOUT_ESPERA)
        boton_explorar.click()
        time.sleep(PAUSA_CORTA * 2)

        # Dialogo estandar de Windows para elegir archivo (clase #32770)
        dialogo = self.app.window(class_name="#32770")
        dialogo.wait("exists ready", timeout=TIMEOUT_ESPERA)
        # Escribimos la ruta completa directo en la barra de nombre de archivo
        dialogo.type_keys(ruta_archivo, with_spaces=True, set_foreground=True)
        time.sleep(PAUSA_CORTA)
        dialogo.type_keys("{ENTER}")
        time.sleep(PAUSA_CORTA * 2)

        boton_agregar = ventana_adjuntar.child_window(
            title=NOMBRE_BOTON_AGREGAR, control_type="Button"
        )
        boton_agregar.wait("ready", timeout=TIMEOUT_ESPERA)
        boton_agregar.click()
        time.sleep(PAUSA_CORTA * 2)

        boton_cerrar = ventana_adjuntar.child_window(
            title=NOMBRE_BOTON_CERRAR_POPUP, control_type="Button"
        )
        boton_cerrar.wait("ready", timeout=TIMEOUT_ESPERA)
        boton_cerrar.click()
        time.sleep(PAUSA_CORTA)


# ---------------------------------------------------------------------------
# PASO 4: FLUJO PRINCIPAL (subir / revisar)
# ---------------------------------------------------------------------------


def procesar_mes(nombre_mes: str, modo_solo_revisar: bool):
    """Recorre todos los IDs de la carpeta del mes.
    Si modo_solo_revisar=True, no adjunta nada, solo informa estado."""
    carpeta_mes = os.path.join(CARPETA_FOTOS_BASE, nombre_mes)
    grupos, no_reconocidos = escanear_carpeta(carpeta_mes)
    ya_procesados = cargar_registro()

    resultados: list[ResultadoID] = []

    if not modo_solo_revisar:
        automatizador = AutomatizadorApp()
        print("Conectando con la aplicacion...")
        automatizador.conectar()
    else:
        automatizador = None

    total = len(grupos)
    for i, (id_cliente, grupo) in enumerate(grupos.items(), start=1):
        print(f"[{i}/{total}] Procesando ID {id_cliente} ...")

        if grupo.archivos_excedente:
            for archivo in grupo.archivos_excedente:
                registrar_resultado(id_cliente, archivo, "EXCEDENTE",
                                     f"Mas de {MAX_ARCHIVOS_POR_ID} archivos para este ID")
            resultados.append(ResultadoID(id_cliente, "EXCEDENTE",
                               f"{len(grupo.archivos_excedente)} archivo(s) excedente(s)"))

        if grupo.archivos_muy_grandes:
            for archivo in grupo.archivos_muy_grandes:
                registrar_resultado(id_cliente, archivo, "MUY_GRANDE",
                                     f"Supera los {MAX_TAMANO_MB}MB")
            resultados.append(ResultadoID(id_cliente, "MUY_GRANDE",
                               f"{len(grupo.archivos_muy_grandes)} archivo(s) > {MAX_TAMANO_MB}MB"))

        # Filtrar archivos ya subidos en corridas anteriores por este script
        archivos_pendientes = [
            a for a in grupo.archivos_validos
            if f"{id_cliente}|{os.path.basename(a)}" not in ya_procesados
        ]

        if not archivos_pendientes:
            continue  # nada nuevo que hacer para este ID

        if modo_solo_revisar:
            resultados.append(ResultadoID(id_cliente, "PENDIENTE",
                               f"{len(archivos_pendientes)} archivo(s) por subir",
                               archivos_pendientes))
            continue

        try:
            encontrado = automatizador.buscar_cliente(id_cliente)
            if not encontrado:
                for a in archivos_pendientes:
                    registrar_resultado(id_cliente, a, "NO_ENCONTRADO", "")
                resultados.append(ResultadoID(id_cliente, "NO_ENCONTRADO"))
                continue

            filas_antes = automatizador.contar_imagenes_adjuntas()
            if filas_antes > 0:
                for a in archivos_pendientes:
                    registrar_resultado(id_cliente, a, "OMITIDO_YA_TENIA",
                                         f"Ya tenia {filas_antes} imagen(es) adjunta(s)")
                resultados.append(ResultadoID(id_cliente, "OMITIDO_YA_TENIA",
                                   f"Ya tenia {filas_antes} imagen(es)"))
                continue

            for archivo in archivos_pendientes:
                automatizador.adjuntar_archivo(archivo)

            filas_despues = automatizador.contar_imagenes_adjuntas()
            esperado = filas_antes + len(archivos_pendientes)

            if filas_despues == esperado:
                for a in archivos_pendientes:
                    registrar_resultado(id_cliente, a, "OK", "")
                resultados.append(ResultadoID(id_cliente, "OK",
                                   f"{len(archivos_pendientes)} archivo(s) subido(s)"))
            else:
                for a in archivos_pendientes:
                    registrar_resultado(id_cliente, a, "VERIF_FALLIDA",
                                         f"Esperado {esperado}, encontrado {filas_despues}")
                resultados.append(ResultadoID(id_cliente, "VERIF_FALLIDA",
                                   f"Esperado {esperado}, encontrado {filas_despues}"))

        except Exception as e:
            traceback.print_exc()
            for a in archivos_pendientes:
                registrar_resultado(id_cliente, a, "ERROR", str(e))
            resultados.append(ResultadoID(id_cliente, "ERROR", str(e)))

    escribir_reporte(nombre_mes, resultados, no_reconocidos, modo_solo_revisar)
    return resultados, no_reconocidos


def escribir_reporte(nombre_mes, resultados, no_reconocidos, modo_solo_revisar):
    conteos = {}
    for r in resultados:
        conteos[r.estado] = conteos.get(r.estado, 0) + 1

    with open(ARCHIVO_REPORTE, "w", encoding="utf-8") as f:
        f.write(f"Reporte - {nombre_mes} - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"Modo: {'Solo revision' if modo_solo_revisar else 'Subida real'}\n")
        f.write("=" * 60 + "\n\n")
        f.write("RESUMEN:\n")
        for estado, cantidad in conteos.items():
            f.write(f"  {estado}: {cantidad}\n")
        f.write(f"\nTotal IDs con novedades: {len(resultados)}\n")
        if no_reconocidos:
            f.write(f"\nArchivos con nombre no reconocido (revisar a mano): {len(no_reconocidos)}\n")
            for nombre in no_reconocidos:
                f.write(f"  - {nombre}\n")
        f.write("\nDETALLE POR ID:\n")
        for r in resultados:
            f.write(f"  ID {r.id_cliente}: {r.estado} - {r.detalle}\n")


# ---------------------------------------------------------------------------
# PASO 5: MENU DE CONSOLA
# ---------------------------------------------------------------------------


def elegir_mes() -> str | None:
    meses = listar_meses_disponibles()
    if not meses:
        print(f"No se encontraron carpetas de mes dentro de:\n  {CARPETA_FOTOS_BASE}")
        return None

    print("\nMeses disponibles:")
    for i, mes in enumerate(meses, start=1):
        print(f"  {i}. {mes}")

    seleccion = input("\nElige el numero del mes: ").strip()
    if not seleccion.isdigit() or not (1 <= int(seleccion) <= len(meses)):
        print("Opcion invalida.")
        return None
    return meses[int(seleccion) - 1]


def menu_principal():
    while True:
        print("\n" + "=" * 50)
        print("  CARGA DE FOTOS - REFACTURACIONES")
        print("=" * 50)
        print("  1. Subir imagenes de un mes")
        print("  2. Revisar estado (sin subir nada)")
        print("  3. Salir")
        opcion = input("\nElige una opcion: ").strip()

        if opcion == "1":
            mes = elegir_mes()
            if mes:
                print(f"\nProcesando '{mes}'... esto puede tomar varios minutos.\n")
                resultados, no_reconocidos = procesar_mes(mes, modo_solo_revisar=False)
                mostrar_resumen(resultados, no_reconocidos)

        elif opcion == "2":
            mes = elegir_mes()
            if mes:
                print(f"\nRevisando '{mes}'...\n")
                resultados, no_reconocidos = procesar_mes(mes, modo_solo_revisar=True)
                mostrar_resumen(resultados, no_reconocidos)

        elif opcion == "3":
            print("Saliendo...")
            break
        else:
            print("Opcion invalida, intenta de nuevo.")

        input("\nPresiona ENTER para volver al menu...")


def mostrar_resumen(resultados, no_reconocidos):
    conteos = {}
    for r in resultados:
        conteos[r.estado] = conteos.get(r.estado, 0) + 1
    print("\n--- RESUMEN ---")
    if not resultados:
        print("  No habia nada nuevo que procesar.")
    for estado, cantidad in conteos.items():
        print(f"  {estado}: {cantidad}")
    if no_reconocidos:
        print(f"  Archivos con nombre raro (revisar a mano): {len(no_reconocidos)}")
    print(f"\nReporte completo guardado en: {ARCHIVO_REPORTE}")


if __name__ == "__main__":
    try:
        menu_principal()
    except Exception:
        print("\n\nOcurrio un error inesperado:")
        traceback.print_exc()
        input("\nPresiona ENTER para cerrar...")
        sys.exit(1)
