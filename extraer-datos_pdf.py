import pdfplumber
import pandas as pd
import os
import re

# Ruta definida por Richard
ruta_carpeta = r'C:\RICHARD\FDL\Usme\2026\Pruebas_pagos'
archivo_salida = os.path.join(ruta_carpeta, "Consolidado_Pagos_Usme.xlsx")

def limpiar_monto(texto):
    if not texto: return 0
    # Extrae solo números, puntos y comas
    limpio = re.sub(r'[^\d,]', '', str(texto)).replace(',', '.')
    try:
        return float(limpio)
    except:
        return 0

def procesar_causacion(pdf_path):
    datos = {}
    with pdfplumber.open(pdf_path) as pdf:
        pagina = pdf.pages[0]
        texto = pagina.extract_text()
        tablas = pagina.extract_tables()
        
        # --- Datos Básicos ---
        datos['Contrato'] = re.search(r'CPS\s*(\d+-\d+)', texto).group(0) if re.search(r'CPS\s*(\d+-\d+)', texto) else "N/A"
        datos['Contratista'] = re.search(r'MARIA\s+[A-Z\s]+', texto).group(0) if re.search(r'MARIA\s+[A-Z\s]+', texto) else "N/A"
        
        # Captura de "DOCUMENTO No." (Ej: PAGO 1 DE 5)
        doc_match = re.search(r'PAGO\s+\d+\s+DE\s+\d+', texto)
        datos['Documento No.'] = doc_match.group(0) if doc_match else "N/A"
        
        # --- Extracción de Valores desde Tablas ---
        for tabla in tablas:
            for fila in tabla:
                # Limpiar celdas de saltos de línea
                fila_limpia = [str(c).replace('\n', ' ').strip() for c in fila if c is not None]
                linea = " ".join(fila_limpia)
                
                if "VALOR BRUTO" in linea:
                    datos['Valor Bruto'] = limpiar_monto(fila_limpia[-1])
                
                if "Reteica" in linea:
                    datos['Reteica'] = limpiar_monto(fila_limpia[-1])
                
                # Captura robusta del NETO A PAGAR
                if "NETO A PAGAR" in linea:
                    # Buscamos el valor que contenga números en esa fila
                    valor_neto = fila_limpia[-1] if fila_limpia[-1] != "" else fila_limpia[-2]
                    datos['Neto a Pagar'] = limpiar_monto(valor_neto)

        # --- Seguridad Social ---
        ss_match = re.search(r'Base para pago.*?([\d\.]+)', texto)
        datos['Base SS'] = limpiar_monto(ss_match.group(1)) if ss_match else 0
        
    return datos

# --- Ejecución ---
resultados = []
if os.path.exists(ruta_carpeta):
    for archivo in os.listdir(ruta_carpeta):
        if archivo.lower().endswith(".pdf"):
            try:
                info = procesar_causacion(os.path.join(ruta_carpeta, archivo))
                info['Archivo'] = archivo
                resultados.append(info)
                print(f"Procesado: {archivo}")
            except Exception as e:
                print(f"Error en {archivo}: {e}")

    if resultados:
        df = pd.DataFrame(resultados)
        # Reordenar columnas para que sea fácil de leer
        columnas = ['Archivo', 'Contrato', 'Documento No.', 'Contratista', 'Valor Bruto', 'Reteica', 'Neto a Pagar', 'Base SS']
        df = df[columnas]
        df.to_excel(archivo_salida, index=False)
        print(f"\nListo! Excel creado con {len(resultados)} registros.")
else:
    print("La ruta no es válida.")