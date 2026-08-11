# Automatización carga de fotos - Refacturaciones

## 1. Requisitos para TU máquina (donde vas a probar)

- Python 3.10 o superior instalado (marca "Add to PATH" al instalar).
- Abrir una terminal (cmd) en esta carpeta y correr:

```
pip install pywinauto
```

## 2. Antes de correr la primera prueba

1. Abre `main.py` y revisa la sección `CONFIG` al inicio del archivo.
   Ahí está la ruta de la carpeta de fotos (`CARPETA_FOTOS_BASE`) y los
   límites de negocio (máx. 3 archivos, máx. 5MB). Ajusta si es necesario.
2. Abre el programa del sistema y déjalo en la ventana "Consultas [Opc:101]"
   ANTES de correr el script (el script se conecta a la ventana que ya
   está abierta, no la abre él mismo).

## 3. Cómo probar

Doble clic en `ejecutar_pruebas.bat`. Se abrirá un menú de texto:

```
1. Subir imagenes de un mes
2. Revisar estado (sin subir nada)
3. Salir
```

**Recomendación:** la primera vez usa la opción **2 (Revisar estado)**
para un mes con pocas fotos. Esto NO toca la app, solo te dice cuántos
archivos hay pendientes por ID — así confirmas que el agrupamiento por
ID está leyendo bien la carpeta antes de tocar el sistema real.

Cuando eso se vea bien, prueba la opción **1 (Subir)** con un solo mes
y, si puedes, con pocos archivos primero (por ejemplo copia 3-4 fotos
de prueba a una carpeta de mes aparte antes de soltarlo con las 90).

## 4. Qué revisar si algo falla

El script imprime en pantalla y guarda en `reporte_ultima_corrida.txt`
el detalle de cada paso. Los errores más probables y dónde ajustarlos
(todo está en la sección `CONFIG` arriba de `main.py`):

| Síntoma | Qué revisar |
|---|---|
| "No se encontro la ventana" / timeout al conectar | `TITULO_VENTANA_PRINCIPAL_REGEX` — puede que el título real tenga otro texto/versión. Anota el título exacto de la barra de la ventana y ajústalo. |
| Falla al escribir el ID en el campo Nro | El campo puede no llamarse exactamente `"Nro:"` como control `Edit`, o el label y el campo de edición sean controles separados. Revisa con Accessibility Insights de nuevo apuntando justo al recuadro (no al texto "Nro:"). |
| Falla al abrir el explorador de Windows / al escribir la ruta | Algunos exploradores de Windows necesitan que el foco esté en la barra de nombre de archivo. Si falla, dime qué pasa exactamente (¿no escribe nada? ¿escribe en otro lado?) para ajustar esa parte. |
| La verificación de cantidad de filas siempre da 0 o error | El nombre de clase `DataGridView` puede variar según cómo esté anidada la ventana de "Ver imagen(es)". Puede que haga falta inspeccionar esa ventana también con Accessibility Insights (la ventana en sí, no solo la tabla) y pasarme el título exacto. |
| Todo funciona pero muy lento / se cae por timing | Sube el valor de `PAUSA_CORTA` (ej. de 0.4 a 0.8) en la sección CONFIG. |

**Cualquier error que salga en pantalla, cópiamelo completo (el
traceback en rojo/texto) y lo ajustamos.**

## 5. Empaquetar como .exe para tu colega (cuando ya funcione bien)

Una vez que probaste y funciona correcto:

```
pip install pyinstaller
pyinstaller --onefile --console --name CargaFotos main.py
```

Esto genera `dist/CargaFotos.exe`. Copia ese `.exe` (y crea un
`.bat` simple que lo llame, o directamente dale el `.exe` con doble
clic) a la carpeta que le entregarás a tu colega. El archivo
`procesados.csv` y `reporte_ultima_corrida.txt` se van a crear solos
en la misma carpeta donde esté el `.exe`.

**Importante:** el `.exe` debe generarse en una máquina Windows (no se
puede generar un .exe de Windows desde Linux/Mac), así que este paso
lo corres tú en tu PC normal.

## 6. Archivos que se generan al usar el programa

- `procesados.csv`: registro histórico de qué se subió, para no repetir.
- `reporte_ultima_corrida.txt`: resumen legible de la última vez que se usó.

Ambos quedan en la misma carpeta donde esté el `.exe`/`main.py`. No se
deben borrar entre ejecuciones (si se borra `procesados.csv`, el
sistema real sigue siendo la protección principal contra duplicados,
porque el script revisa si el ID ya tiene algo adjunto antes de subir).
