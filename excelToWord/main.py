from pathlib import Path
from tkinter import filedialog
from openpyxl import load_workbook
from unicodedata import normalize

import tkinter as tk

# Bloque de codigo para seleccionar donde se crear el Word
def seleccionar_carpeta_destino():
    ventana = tk.Tk()
    ventana.withdraw()
    ventana.attributes("-topmost", True)

    ruta = filedialog.askdirectory(
        title="Selecciona donde crear el Documento"
    )

    ventana.destroy()

    if not ruta:
        print("No se selecciono una carpeta")
        return None

    return Path(ruta)

carpeta_destino = seleccionar_carpeta_destino()

# Bloque de codigo para seleccionar el archivo a leer
def seleccionar_xslx():
    ventana = tk.Tk()
    ventana.withdraw()
    ventana.attributes("-topmost", True)

    archivo = filedialog.askopenfilename(
        title="Selecciona el archivo Excel a leer",
        filetypes=[
            ("Archivos Excel", "*.xlsx *.xls"),
            ("Todos los archivos", "*.*")
        ]
    )

    ventana.destroy()
    return Path(archivo) if archivo else None

# Modificacion del texto para evitar errores
def normalizar_texto(valor):
    # Converitr a mayusculas sin tildes
    texto = str(valor or "").strip().upper()
    return "".join(
        caracter
        for caracter in normalize("NFD", texto)
        if not ("\u0300" <= caracter <= "\u036F")
    )

# Obtener datos de los alumnos
def leer_alumnos(archivo_xls):
    libro = load_workbook(archivo_xls, data_only=True)
    hoja = libro.active

    print(f"Hoja que se leerá: {hoja.title}")

    encabezados = {
        normalizar_texto(celda.value): celda.column
        for celda in hoja[1]
        if celda.value
    }

    columnas__necesarias = {
        "1 COGNOM": "apellido_1",
        "2 COGNOM": "apellido_2",
        "NOM": "nombre",
        "DNI/NIE": "nif",
    }

    for encabezado in columnas__necesarias:
        if encabezado not in encabezados:
            raise ValueError(
                f"No se encontró la columna «{encabezado}» en la fila 1."
            )

    alumnos = []

    for fila in hoja.iter_rows(min_row=2, values_only=True):
        apellido_1 = fila[encabezados["1 COGNOM"] - 1]
        apellido_2 = fila[encabezados["2 COGNOM"] - 1]
        nombre = fila[encabezados["NOM"] - 1]
        nif = fila[encabezados["DNI/NIE"] - 1]

        # ignorar filas sin alumnos
        if not any([apellido_1,apellido_2,nombre,nif]):
            continue

        alumnos.append({
            "apellido_1": str(apellido_1 or "").strip(),
            "apellido_2": str(nombre or "").strip(),
            "nif": str(nif or "").strip(),
        })

    libro.close()
    return alumnos


# Validar si las selecciones fueron correctas
if carpeta_destino is None:
    print("Operacion cancelada")
else:
    archivo_xls = seleccionar_xslx()

    if archivo_xls is None:
        print ("Operacion cancelada")

    else:
        print(f"El Word se guardara en: {carpeta_destino}")
        print(f"Archivo seleccionado: {archivo_xls.name}")

alumnos = leer_alumnos(archivo_xls)

print(f"\nSe han leído {len(alumnos)} alumnos:\n")

for alumno in alumnos:
    print(alumno)