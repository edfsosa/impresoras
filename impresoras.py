import requests
from bs4 import BeautifulSoup
import re
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Progressbar
import threading

# Diccionario de modelos y su mapeo de índices
MODELOS_CONFIG = {
    "Lexmark MX611dhe": [0, 2, 3],
    "Lexmark X466de": [0, 2, 3],
    "Lexmark X464de": [0, 2, 3],
    "Lexmark MX710": [0, 2, 4],
    "Lexmark MS811": [0, 2, 4],
    "Lexmark MS812": [0, 2, 4],
    "Lexmark T654": [0, None, 2]  # None significa que no tiene unidad de imagen
}

# Obtener el modelo desde el topbar
def obtener_topbar(ip):
    try:
        url = f"http://{ip}/cgi-bin/dynamic/topbar.html"
        response = requests.get(url, timeout=5)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
        modelo = soup.find('span', class_='top_prodname')
        return modelo.text.strip() if modelo else None  # Retorna None si no encuentra el modelo
    except:
        return None  # No modificar la celda si hay error

# Obtener los porcentajes desde el estado de la impresora
def obtener_status(ip, modelo):
    try:
        url = f"http://{ip}/cgi-bin/dynamic/printer/PrinterStatus.html"
        response = requests.get(url, timeout=5)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
        porcentajes = re.findall(r'\b\d+%|\b\d+\.\d+%', soup.get_text())

        if modelo not in MODELOS_CONFIG:
            return None, None, None  # No modificar si el modelo no está definido

        indices = MODELOS_CONFIG[modelo]
        toner = float(porcentajes[indices[0]].replace('%', '')) / 100 if indices[0] is not None and len(porcentajes) > indices[0] else None
        kit_mantenimiento = float(porcentajes[indices[1]].replace('%', '')) / 100 if indices[1] is not None and len(porcentajes) > indices[1] else None
        unidad_imagen = float(porcentajes[indices[2]].replace('%', '')) / 100 if indices[2] is not None and len(porcentajes) > indices[2] else None

        return toner, kit_mantenimiento, unidad_imagen
    except:
        return None, None, None  # No modificar la celda en caso de error

# Función principal para actualizar el Excel
def actualizar_excel(archivo_excel, resultado_label, barra_progreso):
    try:
        # Cargar el archivo Excel
        wb = load_workbook(archivo_excel)
        ws = wb.active

        # Leer las IPs desde la columna D y los modelos desde la columna E
        filas = [(fila, ws[f"D{fila}"].value, ws[f"E{fila}"].value) for fila in range(2, ws.max_row + 1) if ws[f"D{fila}"].value]
        
        total_impresoras = len(filas)
        fecha_actual = datetime.today().strftime("%d/%m/%Y")

        # Recorrer todas las IPs y actualizar valores en el Excel
        for i, (fila, ip, modelo) in enumerate(filas, start=1):
            progreso = int((i / total_impresoras) * 100)
            barra_progreso["value"] = progreso
            resultado_label.config(text=f"Monitoreando: {ip} ({progreso}%)")
            resultado_label.update_idletasks()

            toner, kit_mantenimiento, unidad_imagen = obtener_status(ip, modelo)

            # Si se obtuvieron valores, actualizarlos en el Excel
            if toner is not None:
                ws[f"I{fila}"].value = toner
                ws[f"I{fila}"].number_format = "0.00%"

            if unidad_imagen is not None:
                ws[f"J{fila}"].value = unidad_imagen
                ws[f"J{fila}"].number_format = "0.00%"

            if kit_mantenimiento is not None:
                ws[f"K{fila}"].value = kit_mantenimiento
                ws[f"K{fila}"].number_format = "0.00%"

            # Guardar la fecha de monitoreo
            ws[f"L{fila}"].value = fecha_actual

        # Guardar los cambios en el Excel
        wb.save(archivo_excel)

        resultado_label.config(text=f"Monitoreo completado. Archivo actualizado: {archivo_excel}")
        messagebox.showinfo("Éxito", f"Monitoreo completado. Archivo actualizado: {archivo_excel}")
    except Exception as e:
        resultado_label.config(text=f"Error: {e}")
        messagebox.showerror("Error", f"Ocurrió un error: {e}")

# Función para iniciar el monitoreo en un hilo separado
def iniciar_monitoreo(archivo_excel, resultado_label, barra_progreso):
    hilo = threading.Thread(target=actualizar_excel, args=(archivo_excel, resultado_label, barra_progreso))
    hilo.start()

# Función para seleccionar un archivo
def seleccionar_archivo(entry_widget):
    archivo = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
    entry_widget.delete(0, tk.END)
    entry_widget.insert(0, archivo)

# Crear la interfaz gráfica
def crear_interfaz():
    ventana = tk.Tk()
    ventana.title("Monitoreo de Impresoras")
    ventana.geometry("600x300")

    # Archivo de Excel
    tk.Label(ventana, text="Archivo de Excel:").grid(row=0, column=0, padx=10, pady=10)
    entrada_excel = tk.Entry(ventana, width=50)
    entrada_excel.grid(row=0, column=1, padx=10, pady=10)
    tk.Button(ventana, text="Seleccionar", command=lambda: seleccionar_archivo(entrada_excel)).grid(row=0, column=2, padx=10, pady=10)

    # Barra de progreso
    barra_progreso = Progressbar(ventana, orient="horizontal", length=400, mode="determinate")
    barra_progreso.grid(row=1, column=0, columnspan=3, pady=20)

    # Botón para iniciar monitoreo
    resultado_label = tk.Label(ventana, text="", fg="green")
    resultado_label.grid(row=3, column=0, columnspan=3, pady=10)

    tk.Button(
        ventana, text="Iniciar Monitoreo",
        command=lambda: iniciar_monitoreo(entrada_excel.get(), resultado_label, barra_progreso)
    ).grid(row=2, column=0, columnspan=3, pady=10)

    ventana.mainloop()

# Ejecutar la interfaz
crear_interfaz()
