"""
Diagnostico paso 2: lista TODAS las ventanas de nivel superior que
Python puede ver en este momento (visibles, con titulo no vacio),
sin filtrar por "Consultas". Sirve para confirmar:

  a) si el script ve la ventana de la app del sistema en absoluto
     (si no aparece ninguna ventana reconocible de esa app, es un
     problema de permisos: la app corre como administrador y este
     script no).
  b) cual es el titulo EXACTO de la ventana (por si tiene mayusculas,
     espacios, o texto distinto al que asumimos).

Uso:
1. Deja la app del sistema abierta en la ventana "Consultas [Opc:101]".
2. Corre este script IGUAL que corriste ejecutar_pruebas.bat la vez
   que SI funciono (osea, si esa vez usaste "Ejecutar como
   administrador", corre este script tambien como administrador).
3. Copia toda la salida.
"""

from pywinauto import Desktop

print("Listando TODAS las ventanas visibles con titulo...\n")

ventanas = Desktop(backend="win32").windows(visible_only=True)

if not ventanas:
    print("No se detecto NINGUNA ventana visible. Esto casi seguro es")
    print("un tema de permisos: este script no tiene el mismo nivel de")
    print("privilegios que la(s) app(s) abiertas (ej. corren como")
    print("administrador y este script no, o viceversa).")
else:
    for i, w in enumerate(ventanas):
        try:
            titulo = w.window_text()
            if not titulo.strip():
                continue
            print(f"--- Ventana #{i} ---")
            print(f"  Titulo exacto : {titulo!r}")
            print(f"  Process ID    : {w.process_id()}")
            print(f"  Clase         : {w.class_name()}")
            print()
        except Exception as e:
            print(f"  (error leyendo detalles de esta ventana: {e})\n")

print("Listo. Copia esta salida completa (o avisa si salio vacia).")
