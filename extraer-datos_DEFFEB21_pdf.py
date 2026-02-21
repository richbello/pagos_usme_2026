import pdfplumber
import pandas as pd
import os
import re

ruta_carpeta = r"C:\RICHARD\FDL\Usme\2026\Pagos\Febrero\ENTREGA_3"
archivo_salida = os.path.join(ruta_carpeta, "Extracción-grupo3_feb.xlsx")

DEBUG = False  # True para ver texto crudo de archivos con campos vacíos

# ── Patrones compilados una vez ─────────────────────────────────────────────
RE_NIT      = re.compile(r'\d{1,3}(?:\.\d{3}){2,3}')
RE_CONTRATO = re.compile(r'CPS[-\s]*\d+[-\s]*\d+')
_F          = r'(\d{1,2})[/\-\.](\d{1,2})[/\-\.](\d{2,4})'
RE_DEL      = re.compile(r'\bDEL\b\s*:?\s*' + _F, re.IGNORECASE)
RE_AL       = re.compile(r'\bAL\b\s*:?\s*'  + _F, re.IGNORECASE)
RE_RANGO    = re.compile(r'\bDEL\b\s*:?\s*' + _F + r'\s+\bAL\b\s*:?\s*' + _F, re.IGNORECASE)

# Patrones para PAGO No. y PERIODO
RE_PAGO_DE     = re.compile(r'PAGO\s+(?:N[°º.]?\s*)?(\d+)\s+DE\s+(\d+)', re.IGNORECASE)      # "PAGO 3 DE 10"
RE_PAGO_ALT    = re.compile(r'(?:CUOTA|MES|ENTREGA)\s+(\d+)\s+DE\s+(\d+)', re.IGNORECASE)    # "CUOTA 3 DE 10"
RE_DOC_DE      = re.compile(r'(\d+)\s+DE\s+(\d+)')                                            # "10 DE 10" en DOCUMENTO No.
RE_PAGO_NUM    = re.compile(r'^(\d+)')                                                         # primer número de celda ("10 ULTIMO")


def limpiar_monto(texto):
    if not texto:
        return 0
    limpio = re.sub(r'[^\d,]', '', str(texto)).replace(',', '.')
    try:
        return float(limpio)
    except:
        return 0


def normalizar_fecha(m1, m2, m3):
    a = ("20" + m3) if len(m3) == 2 else m3
    return f"{m1.zfill(2)}/{m2.zfill(2)}/{a}"


def procesar_causacion(pdf_path):
    datos = {
        'Archivo': os.path.basename(pdf_path),
        'Contrato': "N/A", 'Documento No.': "N/A",
        'Contratista': "N/A", 'NIT_CC': "N/A",
        'Valor Bruto': 0, 'BASE RETEICA': "N/A",
        'Valor Reteica': 0, 'TOTAL DESCUENTOS': 0,
        'Neto a Pagar': 0, 'PERIODO': 0,
        'DEL': 0, 'AL': 0, 'PAGO No.': 0,
    }

    with pdfplumber.open(pdf_path) as pdf:
        tablas = pdf.pages[0].extract_tables()
        partes = [p.extract_text() for p in pdf.pages]
        texto_full = "\n".join(t for t in partes if t)

    texto_norm = re.sub(r'[ \t]*\n[ \t]*', ' ', texto_full)

    # ── Texto: campos simples ────────────────────────────────────────────────
    m = RE_NIT.search(texto_full)
    if m: datos['NIT_CC'] = m.group(0)

    m = RE_CONTRATO.search(texto_full)
    if m: datos['Contrato'] = m.group(0)

    # PAGO No. y PERIODO desde texto (formato "PAGO 3 DE 10")
    m = RE_PAGO_DE.search(texto_norm) or RE_PAGO_ALT.search(texto_norm)
    if m:
        datos['PAGO No.'] = int(m.group(1))
        datos['PERIODO']  = int(m.group(2))
        datos['Documento No.'] = m.group(0).upper()

    # DEL / AL desde texto
    m = RE_RANGO.search(texto_norm)
    if m:
        datos['DEL'] = normalizar_fecha(m.group(1), m.group(2), m.group(3))
        datos['AL']  = normalizar_fecha(m.group(4), m.group(5), m.group(6))
    else:
        md = RE_DEL.search(texto_norm)
        ma = RE_AL.search(texto_norm)
        if md: datos['DEL'] = normalizar_fecha(md.group(1), md.group(2), md.group(3))
        if ma: datos['AL']  = normalizar_fecha(ma.group(1), ma.group(2), ma.group(3))

    # ── Tablas ───────────────────────────────────────────────────────────────
    for tabla in tablas:
        for i, fila in enumerate(tabla):
            fila_l    = [str(c).replace('\n', ' ').strip() if c else "" for c in fila]
            linea_txt = " ".join(fila_l).upper()

            # ── Montos ──────────────────────────────────────────────────────
            if "CONTRATISTA" in linea_txt and len(fila_l) > 1:
                if re.search(r'[A-Z]{3,}', fila_l[1]):
                    datos['Contratista'] = fila_l[1]
            if "RETEICA" in linea_txt and len(fila_l) >= 6:
                datos['BASE RETEICA']  = fila_l[4]
                datos['Valor Reteica'] = limpiar_monto(fila_l[5])
            if "VALOR BRUTO"      in linea_txt: datos['Valor Bruto']      = limpiar_monto(fila_l[-1])
            if "TOTAL DESCUENTOS" in linea_txt: datos['TOTAL DESCUENTOS'] = limpiar_monto(fila_l[-1])
            if "NETO A PAGAR"     in linea_txt: datos['Neto a Pagar']     = limpiar_monto(fila_l[-1])

            # ── PAGO No. y PERIODO ───────────────────────────────────────────
            # Formato A: "PAGO 3 DE 10" en cualquier celda
            if datos['PAGO No.'] == 0:
                m = RE_PAGO_DE.search(linea_txt) or RE_PAGO_ALT.search(linea_txt)
                if m:
                    datos['PAGO No.'] = int(m.group(1))
                    datos['PERIODO']  = int(m.group(2))

            # Formato B (el de la imagen):
            # Fila tipo: | PAGO No. | 10 ULTIMO | DOCUMENTO No. | 10 DE 10 |
            if datos['PAGO No.'] == 0 and "PAGO No" in linea_txt.upper():
                # Buscar celda con "DOCUMENTO No" y extraer "X DE Y"
                for j, celda in enumerate(fila_l):
                    if "DOCUMENTO" in celda.upper():
                        # La celda siguiente tiene el valor "10 DE 10"
                        val_doc = fila_l[j + 1] if j + 1 < len(fila_l) else ""
                        if not val_doc:
                            # A veces el valor está en la misma celda
                            val_doc = celda
                        m_doc = RE_DOC_DE.search(val_doc)
                        if m_doc:
                            datos['PERIODO'] = int(m_doc.group(2))
                            datos['Documento No.'] = val_doc.strip()

                    # Celda con el número del pago ("10 ULTIMO" o solo "10")
                    if "PAGO No" in celda.upper() or "PAGO No" in (fila_l[j-1].upper() if j > 0 else ""):
                        # El valor está en la celda a la derecha
                        val_pago = fila_l[j + 1] if j + 1 < len(fila_l) else ""
                        m_num = RE_PAGO_NUM.match(val_pago.strip())
                        if m_num:
                            datos['PAGO No.'] = int(m_num.group(1))

            # Formato B alternativo: la celda dice exactamente "PAGO No."
            # y la siguiente tiene el número
            if datos['PAGO No.'] == 0:
                for j, celda in enumerate(fila_l):
                    if re.match(r'^PAGO\s*No', celda.strip(), re.IGNORECASE):
                        val = fila_l[j + 1] if j + 1 < len(fila_l) else ""
                        m_num = RE_PAGO_NUM.match(val.strip())
                        if m_num:
                            datos['PAGO No.'] = int(m_num.group(1))

            # ── DEL / AL desde celdas de tabla ──────────────────────────────
            if datos['DEL'] == 0 or datos['AL'] == 0:
                # Buscar fechas en celdas individuales de la fila
                # (útil cuando DEL y AL son celdas separadas: | DEL | 1/01/2026 | AL | 8/01/2026 |)
                for j, celda in enumerate(fila_l):
                    cel_up = celda.upper().strip()
                    if cel_up == "DEL" and datos['DEL'] == 0:
                        siguiente = fila_l[j + 1] if j + 1 < len(fila_l) else ""
                        md = RE_DEL.search("DEL " + siguiente) or re.match(_F, siguiente.strip())
                        if md:
                            if hasattr(md, 'lastindex') and md.lastindex >= 3:
                                datos['DEL'] = normalizar_fecha(md.group(1), md.group(2), md.group(3))
                    if cel_up == "AL" and datos['AL'] == 0:
                        siguiente = fila_l[j + 1] if j + 1 < len(fila_l) else ""
                        ma = re.match(_F, siguiente.strip())
                        if ma:
                            datos['AL'] = normalizar_fecha(ma.group(1), ma.group(2), ma.group(3))

                # También intentar en linea_txt completa
                if datos['DEL'] == 0 or datos['AL'] == 0:
                    m = RE_RANGO.search(linea_txt)
                    if m:
                        datos['DEL'] = normalizar_fecha(m.group(1), m.group(2), m.group(3))
                        datos['AL']  = normalizar_fecha(m.group(4), m.group(5), m.group(6))
                    else:
                        if datos['DEL'] == 0:
                            md = RE_DEL.search(linea_txt)
                            if md: datos['DEL'] = normalizar_fecha(md.group(1), md.group(2), md.group(3))
                        if datos['AL'] == 0:
                            ma = RE_AL.search(linea_txt)
                            if ma: datos['AL'] = normalizar_fecha(ma.group(1), ma.group(2), ma.group(3))

    datos['_texto_norm'] = texto_norm  # solo para debug
    return datos


def ejecutar_extraccion():
    if not os.path.exists(ruta_carpeta):
        print("La ruta no es válida.")
        return

    archivos = [f for f in os.listdir(ruta_carpeta)
                if f.lower().endswith(".pdf") and not f.startswith("~$")]
    print(f"Procesando {len(archivos)} archivos...\n")

    resultados = []
    for archivo in archivos:
        try:
            info = procesar_causacion(os.path.join(ruta_carpeta, archivo))
            resultados.append(info)
            pago_str = str(info['PAGO No.']) if info['PAGO No.'] else "❌"
            del_str  = info['DEL'] if info['DEL'] else "❌"
            al_str   = info['AL']  if info['AL']  else "❌"
            print(f"  ✓ {archivo:45s} | PAGO: {pago_str:>3} | DEL: {del_str} | AL: {al_str}")
        except Exception as e:
            print(f"  ✘ {archivo}: {e}")

    if not resultados:
        return

    columnas = ['Archivo','Contrato','Documento No.','Contratista','NIT_CC',
                'PERIODO','DEL','AL','PAGO No.','Valor Bruto',
                'BASE RETEICA','Valor Reteica','TOTAL DESCUENTOS','Neto a Pagar']

    df_full = pd.DataFrame(resultados)
    # Asegurar que todas las columnas existen aunque estén vacías
    for col in columnas:
        if col not in df_full.columns:
            df_full[col] = 0

    df_out = df_full[columnas].copy()
    df_out.to_excel(archivo_salida, index=False)
    print(f"\n¡LISTO! → {archivo_salida}")

    # ── Diagnóstico ──────────────────────────────────────────────────────────
    print("\n── Campos vacíos ──────────────────────────────────────────────────")
    hay_vacios = False
    for campo in ['PAGO No.', 'PERIODO', 'DEL', 'AL']:
        mascara = df_out[campo].astype(str).isin(['0', '', 'N/A'])
        vacios  = df_out[mascara]['Archivo'].tolist()
        if vacios:
            hay_vacios = True
            print(f"  ⚠ Sin {campo} ({len(vacios)}): {vacios}")
    if not hay_vacios:
        print("  ✅ Todos los campos extraídos correctamente.")

    # ── Debug ────────────────────────────────────────────────────────────────
    if DEBUG:
        sin_pago = df_full[df_full['PAGO No.'] == 0]
        for _, row in sin_pago.iterrows():
            print(f"\n{'='*60}\nDEBUG → {row['Archivo']}\n{'='*60}")
            print(row['_texto_norm'][:2000])


if __name__ == "__main__":
    ejecutar_extraccion()

