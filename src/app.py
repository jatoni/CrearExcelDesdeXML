import tkinter as tk
from tkinter import filedialog, messagebox
from processor import (
    readXMLAndBuildData,
    agregar_hoja_nueva_excel,
    obtenerDatosAlumnos
)

def generar_folios():
    try:
        excel = filedialog.askopenfilename(
            title="Selecciona el Excel base",
            filetypes=[("Excel", "*.xlsx")]
        )
        if not excel:
            return

        xml = filedialog.askdirectory(title="Selecciona carpeta de XML")
        if not xml:
            return

        data = readXMLAndBuildData(xml, True)
        agregar_hoja_nueva_excel(excel, data)

        messagebox.showinfo("Éxito", "Folios generados correctamente")

    except Exception as e:
        messagebox.showerror("Error", str(e))


def obtener_datos():
    try:
        xml = filedialog.askdirectory(title="Selecciona carpeta de XML")
        if not xml:
            return

        data = readXMLAndBuildData(xml, False)
        obtenerDatosAlumnos(data)

        messagebox.showinfo("Éxito", "Datos generados correctamente")

    except Exception as e:
        messagebox.showerror("Error", str(e))


root = tk.Tk()
root.title("Sistema de Títulos Electrónicos")
root.geometry("450x300")
root.resizable(False, False)

tk.Label(
    root,
    text="Sistema de Títulos Electrónicos",
    font=("Arial", 14, "bold")
).pack(pady=20)

tk.Button(
    root,
    text="Generar Folios de Títulos",
    width=30,
    height=2,
    command=generar_folios
).pack(pady=10)

tk.Button(
    root,
    text="Obtener Datos de Alumnos",
    width=30,
    height=2,
    command=obtener_datos
).pack(pady=10)

tk.Button(
    root,
    text="Salir",
    width=30,
    command=root.destroy
).pack(pady=10)

root.mainloop()