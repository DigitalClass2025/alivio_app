import tkinter as tk
from tkinter import Toplevel, scrolledtext, filedialog, messagebox
import pandas as pd
from validators import validar_columnas, validar_tipos

# --------------------------------------------
# Función que abre la ventana de validación
# --------------------------------------------
def abrir_validacion():
    # Crear nueva ventana
    ventana_val = Toplevel(root)
    ventana_val.title("Validador de Catálogos")
    ventana_val.geometry("700x500")

    # Función para seleccionar y validar archivo
    def validar_archivo():
        file_path = filedialog.askopenfilename(filetypes=[("CSV and Excel files", "*.csv *.xlsx *.xls")])
        if not file_path:
            return

        try:
            if file_path.endswith(".csv"):
                df = pd.read_csv(file_path)
            else:
                df = pd.read_excel(file_path)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer el archivo: {e}")
            return

        errores = []
        errores += validar_columnas(df)
        errores += validar_tipos(df)

        text_area.delete("1.0", tk.END)
        if errores:
            text_area.insert(tk.END, "Errores encontrados:\n\n")
            for err in errores:
                text_area.insert(tk.END, f"{err}\n")
        else:
            text_area.insert(tk.END, "Catálogo validado correctamente!")

    # Botón para validar archivo
    btn_validar = tk.Button(ventana_val, text="Seleccionar archivo y validar", command=validar_archivo, font=("Arial", 14))
    btn_validar.pack(pady=20)

    # Área de texto para mostrar errores
    text_area = scrolledtext.ScrolledText(ventana_val, width=80, height=25, font=("Arial", 12))
    text_area.pack(padx=10, pady=10)

# --------------------------------------------
# Ventana principal
# --------------------------------------------
root = tk.Tk()
root.title("Aplicación Principal - AlivioMeds")
root.geometry("400x200")

lbl = tk.Label(root, text="Bienvenido a la aplicación de AlivioMeds", font=("Arial", 14))
lbl.pack(pady=30)

btn_abrir = tk.Button(root, text="Validar", command=abrir_validacion, font=("Arial", 16), bg="green", fg="white")
btn_abrir.pack(pady=20)

root.mainloop()

