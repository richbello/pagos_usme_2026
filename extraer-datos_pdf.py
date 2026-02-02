import pdfplumber
import pandas as pd
import os
import re

# Ruta de Richard
ruta_carpeta = r'C:\RICHARD\FDL\Usme\2026\Pruebas_pagos'
archivo_salida = os.path.join(ruta_carpeta, "Consolidado_Pagos_Final_Corregido.xlsx")

def limpiar_monto(texto):
    if not texto: return 0
    limpio = re.sub(r'[^\d,]', '', str(texto)).replace(',', '.')
    try:
        return float(limpio)
    except:
        return 0

def extraer_cedula_robusta(texto):
    """
    Busca cualquier número con formato de identificación (ej. 1.022.962.992)
    en todo el texto del documento.
    """
    # Busca patrones de números con puntos (cédulas/NIT)
    match = re.search(r'\d{1,3}(?:\.\d{3}){2,3}', texto)
    return match.group(0) if match else "N/A"

def procesar_causacion(pdf_path):
    datos = {
        'Archivo': os.path.basename(pdf_path),
        'Contrato': "N/A",
        'Documento No.': "N/A",
        'Contratista': "N/A",
        'NIT_CC': "N/A",
        'Periodo': "N/A",
        'Valor Bruto': 0,
        'Reteica': 0,
        'Neto a Pagar': 0
    }
    
    with pdfplumber.open(pdf_path) as pdf:
        pagina = pdf.pages[0]
        texto_completo = pagina.extract_text()
        
        # 1. Extraer Cédula/NIT de cualquier parte del texto (Solución para Sindy)
        datos['NIT_CC'] = extraer_cedula_robusta(texto_completo)
        
        # 2. Extraer Contrato y Documento (Pago)
        cnt = re.search(r'CPS[-\s]*\d+[-\s]*\d+', texto_completo)
        if cnt: datos['Contrato'] = cnt.group(0)
        
        doc = re.search(r'PAGO\s+\d+\s+DE\s+\d+', texto_completo, re.IGNORECASE)
        if doc: datos['Documento No.'] = doc.group(0).upper()

        # 3. Extraer Nombre y Valores de las Tablas
        tablas = pagina.extract_tables()
        for tabla in tablas:
            for i, fila in enumerate(tabla):
                fila_l = [str(c).replace('\n', ' ').strip() if c else "" for c in fila]
                linea_txt = " ".join(fila_l).upper()
                
                # Nombre del Contratista
                if "CONTRATISTA" in linea_txt:
                    nombre = fila_l[1] if len(fila_l) > 1 and fila_l[1] != "" else "N/A"
                    # Si no está al lado, unir la fila
                    if nombre == "N/A":
                        nombre = linea_txt.replace("CONTRATISTA:", "").split("NIT")[0].strip()
                    datos['Contratista'] = nombre

                # Valores Financieros
                if "VALOR BRUTO" in linea_txt: datos['Valor Bruto'] = limpiar_monto(fila_l[-1])
                if "RETEICA" in linea_txt: datos['Reteica'] = limpiar_monto(fila_l[-1])
                if "NETO A PAGAR" in linea_txt: datos['Neto a Pagar'] = limpiar_monto(fila_l[-1])

        # 4. Periodo
        fechas = re.findall(r'\d{2}/\d{2}/\d{4}', texto_completo)
        if len(fechas) >= 2:
            datos['Periodo'] = f"{fechas[-2]} al {fechas[-1]}"
            
    return datos

# --- Ejecución ---
resultados = []
if os.path.exists(ruta_carpeta):
    archivos = [f for f in os.listdir(ruta_carpeta) if f.lower().endswith(".pdf")]
    for archivo in archivos:
        try:
            info = procesar_causacion(os.path.join(ruta_carpeta, archivo))
            resultados.append(info)
            print(f"✓ Procesado: {archivo} -> CC: {info['NIT_CC']}")
        except:
            print(f"✘ Error en: {archivo}")

    if resultados:
        df = pd.DataFrame(resultados)
        columnas = ['Archivo', 'Contrato', 'Documento No.', 'Contratista', 'NIT_CC', 'Periodo', 'Valor Bruto', 'Reteica', 'Neto a Pagar']
        df[columnas].to_excel(archivo_salida, index=False)
        print(f"\n¡ÉXITO! Datos consolidados en: {archivo_salida}")
