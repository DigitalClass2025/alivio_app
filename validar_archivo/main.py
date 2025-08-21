import tkinter as tk
from tkinter import filedialog
from validators import validar_archivo

def seleccionar_archivo():
    root = tk.Tk()
    root.withdraw()  # Ocultar ventana principal

    archivo = filedialog.askopenfilename(
        title="Seleccionar archivo normalizado",
        filetypes=[("Archivos Excel", "*.xlsx"), ("Archivos CSV", "*.csv")]
    )

    if archivo:
        print(f"\n📂 Archivo seleccionado: {archivo}")
        resultado = validar_archivo(archivo)
        if resultado:
            print("\n🎉 Validación finalizada con éxito.")
        else:
            print("\n⚠ Validación detenida por errores en el archivo.")
    else:
        print("❌ No se seleccionó ningún archivo.")

if __name__ == "__main__":
    seleccionar_archivo()

