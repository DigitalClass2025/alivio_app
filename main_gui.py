import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
from validators import validar_columnas, validar_tipos

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

# Crear ventana principal
root = tk.Tk()
root.title("Validador de Catálogos AlivioMeds")
root.geometry("700x500")

# Botón para validar
btn_validar = tk.Button(root, text="Seleccionar archivo y validar", command=validar_archivo, font=("Arial", 14))
btn_validar.pack(pady=20)

# Área de texto para mostrar errores
text_area = scrolledtext.ScrolledText(root, width=80, height=25, font=("Arial", 12))
text_area.pack(padx=10, pady=10)

root.mainloop()