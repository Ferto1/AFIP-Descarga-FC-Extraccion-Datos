import tkinter as tk
from tkinter import filedialog
import glob
import pdfplumber
import pandas as pd
from pathlib import Path
import re

# Crear una ventana de diálogo para seleccionar la carpeta
root = tk.Tk()
root.withdraw()
folder_path = filedialog.askdirectory(title="Seleccionar carpeta con archivos PDF")

# Verificar si se seleccionó una carpeta
if not folder_path:
    print("No se seleccionó ninguna carpeta.")
    exit()

# Patrón de búsqueda para los archivos PDF
pdf_pattern = "*.pdf"

# Obtener la lista de archivos PDF en la carpeta
pdf_files = glob.glob(f"{folder_path}/{pdf_pattern}")

# Lista para almacenar los DataFrames de cada PDF
df_list = []

for file_path in pdf_files:
    with pdfplumber.open(file_path) as pdf:
        texts = []
        for page in pdf.pages:
            texts.append(page.extract_text())

        # Definir patrones para extraer información
        punto_venta_pattern = r"Punto de Venta:\s?(\d+)"
        numero_factura_pattern = r"Comp\. Nro:\s?(\d+)"
        fecha_emision_pattern = r"Fecha de Emisión: (\d{2}/\d{2}/\d{4})"
        importe_total_pattern = r"Importe Total: \$ ([\d.,]+)"
        periodo_facturado_pattern = r"Período Facturado Desde: (\d{2}/\d{2}/\d{4}) Hasta:(\d{2}/\d{2}/\d{4})"
        cuit_pattern = r"CUIT:\s?(\d+)"
        rs_pattern = r"Apellido y Nombre / Razón Social:(.+)"
        
        # Inicializar variables para almacenar información
        punto_venta = "0"
        numero_factura = "0"
        fecha_emision = None
        fecha_desde = None
        fecha_hasta = None
        importe_total = "0"
        cuit = "0"
        rs = None
        
        # Iterar a través de los textos extraídos y extraer información usando expresiones regulares
        for text in texts:
            # Extraer información usando patrones regex
            punto_venta_match = re.search(punto_venta_pattern, text)
            if punto_venta_match:
                punto_venta = punto_venta_match.group(1)
                
            numero_factura_match = re.search(numero_factura_pattern, text)
            if numero_factura_match:
                numero_factura = numero_factura_match.group(1)
                
            fecha_emision_match = re.search(fecha_emision_pattern, text)
            if fecha_emision_match:
                fecha_emision = fecha_emision_match.group(1)
                
            importe_total_match = re.search(importe_total_pattern, text)
            if importe_total_match:
                importe_total = importe_total_match.group(1)
                
            periodo_facturado_match = re.search(periodo_facturado_pattern, text)
            if periodo_facturado_match:
                fecha_desde = periodo_facturado_match.group(1)
                fecha_hasta = periodo_facturado_match.group(2)
                
            rs_match = re.search(rs_pattern, text)
            if rs_match:
                rs = rs_match.group(1)
                
            cuit_count = 0  # Contador de coincidencias de CUIT
            for match in re.finditer(cuit_pattern, text):
                cuit_count += 1
                if cuit_count == 2:
                    cuit = match.group(1)
                    break
        
        # Crear diccionario con la informacion extraida
        extracted_info = {
            "CUIT": cuit,
            "Razon Social": rs,
            "Punto de venta": punto_venta,
            "N° de comprobante": numero_factura,
            "Fecha de Emision": fecha_emision,
            "Desde": fecha_desde,
            "Hasta": fecha_hasta,
            "Importe Total": importe_total,
        }

        # Agregar el DataFrame a la lista
        df_list.append(pd.DataFrame([extracted_info]))
        
        # Combinar los DataFrames en uno solo
df_combined = pd.concat(df_list)

# Especificar la ruta del archivo de salida
output_file = Path(folder_path) / "informacion_facturas.xlsx"

# Guardar el DataFrame en un archivo de Excel
df_combined.to_excel(output_file, index=False)

print(f"La información ha sido guardada en el archivo: {output_file}")