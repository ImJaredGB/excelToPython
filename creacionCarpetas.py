from pathlib import Path
import tkinter as tk
from tkinter import filedialog

CARPETAS = [
    "0 BIOCAP",
    "ALUMNES",
    "ASSISTÈNCIA TRANSPORTS",
    "CRONOGRAMA",
    "FOTOS",
]

def seleccionar_carpeta():
    ventana = tk.Tk()
    ventana.withdraw()  # Oculta la ventana principal vacía
    ventana.attributes("-topmost", True)

    ruta = filedialog.askdirectory(
        title="Selecciona dónde crear la carpeta principal"
    )

    ventana.destroy()
    return ruta

def main():
    ruta_base = seleccionar_carpeta()

    if not ruta_base:
        print("Operación cancelada: no se seleccionó ninguna carpeta.")
        return

    nombre_principal = input("Nombre de la carpeta principal: ").strip()

    if not nombre_principal:
        print("Error: el nombre es obligatorio.")
        return

    carpeta_principal = Path(ruta_base) / nombre_principal

    if carpeta_principal.exists():
        print(f"Error: la carpeta «{nombre_principal}» ya existe.")
        return

    try:
        carpeta_principal.mkdir()

        for nombre in CARPETAS:
            (carpeta_principal / nombre).mkdir()

        print(f"\nEstructura creada correctamente en:\n{carpeta_principal}")

    except OSError as error:
        print(f"Error al crear las carpetas: {error}")

if __name__ == "__main__":
    main()