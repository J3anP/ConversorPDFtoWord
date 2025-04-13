import tkinter as tk
from tkinter import filedialog, messagebox
from pdf2docx import Converter
import os

def convertir_pdf_a_word():
    archivos = filedialog.askopenfilenames(
        title="Selecciona archivos PDF",
        filetypes=[("Archivos PDF", "*.pdf")]
    )

    if not archivos:
        return

    errores = []

    for archivo in archivos:
        if archivo.lower().endswith(".pdf"):
            try:
                salida = archivo.replace(".pdf", ".docx")
                cv = Converter(archivo)
                cv.convert(salida, start=0, end=None)
                cv.close()
                print(f"✅ Convertido: {salida}")
            except Exception as e:
                errores.append((archivo, str(e)))

    if errores:
        msg = "Algunos archivos no se pudieron convertir:\n"
        msg += "\n".join(f"{a} - {e}" for a, e in errores)
        messagebox.showerror("Errores durante la conversión", msg)
    else:
        messagebox.showinfo("Éxito", "Todos los archivos fueron convertidos correctamente.")

# Crear ventana oculta
root = tk.Tk()
root.withdraw()

# Ejecutar conversión
convertir_pdf_a_word()
