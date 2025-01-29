import threading
import requests
from bs4 import BeautifulSoup
import re
import json
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Progressbar

# Función para cargar las IPs desde un archivo JSON
def cargar_ips(filename):
    with open(filename, "r") as file:
        data = json.load(file)
    return data["impresoras"]

# Función para cargar la configuración de modelos desde un archivo JSON
def cargar_modelos(filename="modelos.json"):
    with open(filename, "r") as file:
        return json.load(file)

# Obtener el modelo desde el topbar
def obtener_topbar(ip):
    try:
        url = f"http://{ip}/cgi-bin/dynamic/topbar.html"
        response = requests.get(url, timeout=5)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
        modelo = soup.find('span', class_='top_prodname')
        return modelo.text.strip() if modelo else "No encontrado"
    except:
        return "Error desconocido"

# Obtener los porcentajes desde el estado de la impresora
def obtener_status(ip, modelo, config_modelos):
    try:
        url = f"http://{ip}/cgi-bin/dynamic/printer/PrinterStatus.html"
        response = requests.get(url, timeout=5)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
        porcentajes = re.findall(r'\b\d+%|\b\d+\.\d+%', soup.get_text())
        
        if modelo not in config_modelos:
            return "No encontrado", "No encontrado", "No encontrado"
        
        indices = config_modelos[modelo]
        toner = porcentajes[indices[0]].replace('%', '') if indices[0] is not None and len(porcentajes) > indices[0] else "No encontrado"
        kit_mantenimiento = porcentajes[indices[1]].replace('%', '') if indices[1] is not None and len(porcentajes) > indices[1] else "No encontrado"
        unidad_imagen = porcentajes[indices[2]].replace('%', '') if indices[2] is not None and len(porcentajes) > indices[2] else "No encontrado"
        
        return toner, kit_mantenimiento, unidad_imagen
    except:
        return "Error desconocido", "Error desconocido", "Error desconocido"

# Aplicar formato condicional
def aplicar_formato_condicional(archivo_excel, umbral_critico):
    wb = load_workbook(archivo_excel)
    ws = wb.active
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    columnas_interes = ["C", "D", "E"]
    for columna in columnas_interes:
        for celda in ws[columna][1:]:
            if celda.value != "No encontrado" and celda.value != "Error desconocido":
                try:
                    if int(celda.value) < umbral_critico:
                        celda.fill = red_fill
                except ValueError:
                    pass
    wb.save(archivo_excel)

# Función principal para monitorear impresoras
def monitorear_impresoras(archivo_ips, archivo_modelos, umbral_critico, archivo_salida, resultado_label, barra_progreso):
    try:
        # Cargar las IPs y modelos
        ips = cargar_ips(archivo_ips)
        config_modelos = cargar_modelos(archivo_modelos)
        total_impresoras = len(ips)

        # Obtener datos de las impresoras
        datos = []
        for i, ip in enumerate(ips, start=1):
            # Actualizar barra de progreso y mensaje
            progreso = int((i / total_impresoras) * 100)
            barra_progreso["value"] = progreso
            resultado_label.config(text=f"Monitoreando: {ip} ({progreso}%)")
            resultado_label.update_idletasks()
            
            modelo = obtener_topbar(ip)
            toner, kit_mantenimiento, unidad_imagen = obtener_status(ip, modelo, config_modelos)
            datos.append({
                "IP": ip,
                "Modelo": modelo,
                "Tóner Negro (%)": toner,
                "Kit Mantenimiento (%)": kit_mantenimiento,
                "Unidad Imagen (%)": unidad_imagen
            })
        
        # Guardar los datos en un archivo Excel
        df = pd.DataFrame(datos)
        df.to_excel(archivo_salida, index=False)

        # Aplicar formato condicional
        aplicar_formato_condicional(archivo_salida, umbral_critico)

        # Actualizar la interfaz al finalizar
        resultado_label.config(text=f"Monitoreo completado. Archivo guardado como {archivo_salida}.")
        messagebox.showinfo("Éxito", f"Monitoreo completado. Archivo guardado como {archivo_salida}.")
    except Exception as e:
        resultado_label.config(text=f"Error: {e}")
        messagebox.showerror("Error", f"Ocurrió un error: {e}")

# Función para iniciar el monitoreo en un hilo separado
def iniciar_monitoreo(archivo_ips, archivo_modelos, umbral_critico, archivo_salida, resultado_label, barra_progreso):
    hilo = threading.Thread(target=monitorear_impresoras, args=(archivo_ips, archivo_modelos, umbral_critico, archivo_salida, resultado_label, barra_progreso))
    hilo.start()

# Función para seleccionar un archivo JSON
def seleccionar_archivo(entry_widget):
    archivo = filedialog.askopenfilename(filetypes=[("Archivos JSON", "*.json")])
    entry_widget.delete(0, tk.END)
    entry_widget.insert(0, archivo)

# Crear la interfaz gráfica
def crear_interfaz():
    ventana = tk.Tk()
    ventana.title("Monitoreo de Impresoras")
    ventana.geometry("600x500")

    # Archivo de IPs
    tk.Label(ventana, text="Archivo de IPs:").grid(row=0, column=0, padx=10, pady=10)
    entrada_ips = tk.Entry(ventana, width=50)
    entrada_ips.grid(row=0, column=1, padx=10, pady=10)
    tk.Button(ventana, text="Seleccionar", command=lambda: seleccionar_archivo(entrada_ips)).grid(row=0, column=2, padx=10, pady=10)

    # Archivo de modelos
    tk.Label(ventana, text="Archivo de Modelos:").grid(row=1, column=0, padx=10, pady=10)
    entrada_modelos = tk.Entry(ventana, width=50)
    entrada_modelos.grid(row=1, column=1, padx=10, pady=10)
    tk.Button(ventana, text="Seleccionar", command=lambda: seleccionar_archivo(entrada_modelos)).grid(row=1, column=2, padx=10, pady=10)

    # Umbral crítico
    tk.Label(ventana, text="Umbral crítico (%):").grid(row=2, column=0, padx=10, pady=10)
    entrada_umbral = tk.Entry(ventana, width=10)
    entrada_umbral.insert(0, "20")  # Valor predeterminado
    entrada_umbral.grid(row=2, column=1, sticky="w", padx=10, pady=10)

    # Archivo de salida
    tk.Label(ventana, text="Archivo de salida:").grid(row=3, column=0, padx=10, pady=10)
    entrada_salida = tk.Entry(ventana, width=50)
    entrada_salida.insert(0, "informacion_impresoras.xlsx")  # Valor predeterminado
    entrada_salida.grid(row=3, column=1, padx=10, pady=10)

    # Barra de progreso
    barra_progreso = Progressbar(ventana, orient="horizontal", length=400, mode="determinate")
    barra_progreso.grid(row=4, column=0, columnspan=3, pady=20)

    # Botón para iniciar monitoreo
    resultado_label = tk.Label(ventana, text="", fg="green")
    resultado_label.grid(row=6, column=0, columnspan=3, pady=10)

    tk.Button(
        ventana, text="Iniciar Monitoreo",
        command=lambda: iniciar_monitoreo(
            entrada_ips.get(), entrada_modelos.get(), int(entrada_umbral.get()), entrada_salida.get(), resultado_label, barra_progreso
        )
    ).grid(row=5, column=0, columnspan=3, pady=10)

    ventana.mainloop()

# Ejecutar la interfaz
crear_interfaz()
