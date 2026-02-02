import pdfplumber
import pandas as pd
import os
import re

# Ruta definida por Richard
ruta_carpeta = r'C:\RICHARD\FDL\Usme\2026\Pruebas_pagos'
archivo_salida = os.path.join(ruta_carpeta, "Consolidado_Pagos_Usme_Final.xlsx")

def limpiar_monto(texto):
    if not texto: return 0
    # Limpia símbolos y ajusta decimales
    limpio = re.sub(r'[^\d,]', '', str(texto)).replace(',', '.')
    try:
        return float(limpio)
    except:
        return 0

def procesar_causacion(pdf_path):
    datos = {
        'Archivo': os.path.basename(pdf_path),
        'Contrato': "N/A",
        'Documento No.': "N/A",
        'Contratista': "N/A",
        'NIT_CC': "N/A",
        'Valor Bruto': 0,
        'BASE RETEICA': "N/A", # Contiene el porcentaje %
        'Valor Reteica': 0,
        'TOTAL DESCUENTOS': 0, # Nueva columna solicitada
        'Neto a Pagar': 0
    }
    
    with pdfplumber.open(pdf_path) as pdf:
        pagina = pdf.pages[0]
        texto_completo = pagina.extract_text()
        tablas = pagina.extract_tables()
        
        # 1. Identificación y Encabezados
        id_match = re.search(r'\d{1,3}(?:\.\d{3}){2,3}', texto_completo)
        if id_match: datos['NIT_CC'] = id_match.group(0)
        
        cnt = re.search(r'CPS[-\s]*\d+[-\s]*\d+', texto_completo)
        if cnt: datos['Contrato'] = cnt.group(0)
        
        doc = re.search(r'PAGO\s+\d+\s+DE\s+\d+', texto_completo, re.IGNORECASE)
        if doc: datos['Documento No.'] = doc.group(0).upper()

        # 2. Análisis de Tablas
        for tabla in tablas:
            for fila in tabla:
                fila_l = [str(c).replace('\n', ' ').strip() if c else "" for c in fila]
                linea_txt = " ".join(fila_l).upper()
                
                # Nombre del Contratista
                if "CONTRATISTA" in linea_txt and len(fila_l) > 1:
                    if re.search(r'[A-Z]{3,}', fila_l[1]):
                        datos['Contratista'] = fila_l[1]
                
                # Reteica (Porcentaje y Valor)
                if "RETEICA" in linea_txt and len(fila_l) >= 6:
                    datos['BASE RETEICA'] = fila_l[4] 
                    datos['Valor Reteica'] = limpiar_monto(fila_l[5])

                # Valores Totales
                if "VALOR BRUTO" in linea_txt:
                    datos['Valor Bruto'] = limpiar_monto(fila_l[-1])
                
                if "TOTAL DESCUENTOS" in linea_txt:
                    datos['TOTAL DESCUENTOS'] = limpiar_monto(fila_l[-1])
                
                if "NETO A PAGAR" in linea_txt:
                    datos['Neto a Pagar'] = limpiar_monto(fila_l[-1])
            
    return datos

def ejecutar_extraccion():
    resultados = []
    if os.path.exists(ruta_carpeta):
        archivos = [f for f in os.listdir(ruta_carpeta) if f.lower().endswith(".pdf") and not f.startswith("~$")]
        
        print(f"Iniciando extracción de {len(archivos)} archivos...")
        
        for archivo in archivos:
            try:
                info = procesar_causacion(os.path.join(ruta_carpeta, archivo))
                resultados.append(info)
                print(f"✓ {archivo} procesado.")
            except Exception as e:
                print(f"✘ Error en {archivo}: {e}")

        if resultados:
            df = pd.DataFrame(resultados)
            # Orden final de columnas
            columnas_orden = ['Archivo', 'Contrato', 'Documento No.', 'Contratista', 'NIT_CC', 
                              'Valor Bruto', 'BASE RETEICA', 'Valor Reteica', 'TOTAL DESCUENTOS', 'Neto a Pagar']
            df[columnas_orden].to_excel(archivo_salida, index=False)
            print(f"\n¡PROCESO COMPLETADO! Archivo disponible en:\n{archivo_salida}")
    else:
        print("La ruta no es válida.")

if __name__ == "__main__":
    ejecutar_extraccion()
