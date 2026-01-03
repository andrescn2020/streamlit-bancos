import streamlit as st
import pdfplumber
import re
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.formatting.rule import CellIsRule

# Regex
ILLEGAL_CHARACTERS_RE = re.compile(r'[\000-\010]|[\013-\014]|[\016-\037]')

def clean_for_excel(text):
    if not text: return ""
    text = str(text)
    text = ILLEGAL_CHARACTERS_RE.sub("", text)
    return text.strip()

def parse_amount(s):
    if not s: return 0.0
    s = s.strip().replace("$", "").replace(" ", "")
    # "-1.144,92" o "1.144,92-"
    sign = 1.0
    if s.endswith("-"):
        sign = -1.0
        s = s[:-1]
    elif s.startswith("-"):
        sign = -1.0
        s = s[1:]
        
    try:
        val = float(s.replace(".", "").replace(",", "."))
        return val * sign
    except:
        return 0.0

def procesar_macro_formato_3(archivo_pdf):
    st.info("Procesando archivo del Banco Macro (Formato 3 - Lógica Geométrica)...")
    try:
        archivo_pdf.seek(0)
        
        all_rows = []
        
        # Metadatos Defaults
        titular = "Sin Especificar"
        periodo = "Sin Especificar" 
        cuenta = "Sin Especificar"

        # 1. Extracción Geométrica (Words)
        with pdfplumber.open(io.BytesIO(archivo_pdf.read())) as pdf:
            # Debug texto raw
            full_text = ""
            for p in pdf.pages: full_text += p.extract_text() + "\n"
            
            with st.expander("Ver Texto Extraído (Debug)"):
                st.text_area("Contenido del PDF", full_text, height=300)

            # Metadata Regex
            match_cuenta = re.search(r"Número de cuenta\s+(\d+)", full_text)
            if match_cuenta:
                cuenta = match_cuenta.group(1)
            
            # --- Auto-detectar columnas (Sampleo de primeras paginas) ---
            # Necesitamos saber donde empieza "Importe" y "Saldo" aprox.
            # Fecha suele estar a la izquierda (x < 100)
            # Importes a la derecha (x > 300)
            
            # Valores por defecto conservadores
            X_DATE_END = 80
            X_IMPORTE_START = 350 # A partir de aqui buscamos montos
            
            # Iteramos paginas para extraer filas
            for page in pdf.pages:
                words = page.extract_words(x_tolerance=2, y_tolerance=3, keep_blank_chars=True)
                
                # Agrupar palabras en renglones visuales
                # Ordenar por Y (top)
                words.sort(key=lambda w: w['top'])
                
                lines = []
                if not words: continue
                
                current_line = [words[0]]
                for w in words[1:]:
                    # Si la palabra esta en la misma altura (con tolerancia)
                    last_w = current_line[-1]
                    if abs(w['top'] - last_w['top']) < 5:
                        current_line.append(w)
                    else:
                        lines.append(current_line)
                        current_line = [w]
                lines.append(current_line)
                
                # Clasificar contenido de cada línea
                for line_words in lines:
                    # Ordenar palabras por X
                    line_words.sort(key=lambda w: w['x0'])
                    
                    row_data = {
                        "has_date": False,
                        "date_str": "",
                        "desc_words": [],
                        "amount_words": [], # Lista de palabras candidatas a monto (derecha)
                        "y": line_words[0]['top'] 
                    }
                    
                    # Regex checkers
                    re_date = re.compile(r"^\d{2}/\d{2}/\d{4}$")
                    re_amt = re.compile(r"^-?\$?[\d\.,]+-?$") # $ 1.000,00 o 100-
                    
                    # Filtro Header: Si la linea contiene palabras clave de header, la ignoramos completamente
                    line_text = " ".join([w['text'] for w in line_words])
                    if any(x in line_text for x in ["Fecha Descripción", "Últimos movimientos", "Número de cuenta", "Transacción", "Nro."]):
                         # Check extra: que no sea un Falso Positivo (una descripcion real que menciona eso?)
                         # Headers suelen estar muy arriba (top < 150) o tener formato especifico.
                         # Por simplicidad, si matchea keywords fuertes de header, skip.
                         continue

                    for w in line_words:
                        text = w['text'].strip()
                        x_mid = (w['x0'] + w['x1']) / 2
                        
                        # Columna FECHA
                        if x_mid < X_DATE_END and re_date.match(text):
                            row_data["has_date"] = True
                            row_data["date_str"] = text
                            continue
                            
                        # Columna IMPORTES (Derecha)
                        # Chequeamos si parece monto y está a la derecha
                        if x_mid > X_IMPORTE_START:
                             # Es monto o parte de saldo?
                             # Criterio laxo: si tiene numeros y comas/puntos
                             if re.search(r"[\d]+", text):
                                 row_data["amount_words"].append(text)
                                 continue
                        
                        # Si no es Fecha ni Monto Derecha -> ES DESCRIPCIÒN
                        # (Incluso si está a la derecha pero es texto, como "Saldo")
                        if "Saldo" in text: continue # Ignorar label Saldo
                        
                        row_data["desc_words"].append(text)
                    
                    # Post-procesamiento de la línea
                    # Convertir lista de palabras de monto en valores float
                    # A veces "$ 100" son dos palabras. extract_words a veces las separa.
                    # Vamos a unir las palabras de amount y buscar patterns
                    amt_text_full = " ".join(row_data["amount_words"])
                    amounts_found = re.findall(r"-?\$?\s*[\d\.\s]+,\d{2}-?", amt_text_full)
                    
                    row_data["amounts_vals"] = [parse_amount(a) for a in amounts_found]
                    
                    # Guardar fila si tiene contenido relevante
                    if row_data["has_date"] or row_data["desc_words"] or row_data["amounts_vals"]:
                         all_rows.append(row_data)

        # 2. Reconstrucción Lógica (Anchor based)
        transactions = []
        
        # Identificar indices de anchors
        anchor_indices = [i for i, r in enumerate(all_rows) if r["has_date"]]
        used_rows = set()
        
        for i, idx in enumerate(anchor_indices):
            row = all_rows[idx]
            used_rows.add(idx)
            
            fecha = row["date_str"]
            m_vals = row["amounts_vals"]
            
            amount_val = 0.0
            saldo_val = 0.0
            
            resolved = False
            
            # --- Resolución Montos ---
            if len(m_vals) >= 2:
                # [Importe, Saldo] o [Importe_neg, Saldo]
                # Asumimos orden visual: Izq=Importe, Der=Saldo.
                # Como extrajimos palabras sorted x0, el orden se mantiene
                amount_val = m_vals[-2]
                saldo_val = m_vals[-1]
                resolved = True
            
            # Buscar Arriba (Orphan Amount)
            if not resolved:
                k = idx - 1
                if k >= 0 and k not in used_rows:
                    prev_row = all_rows[k]
                    # Si tiene montos y poca descripcion (o descripcion basura de monto)
                    if prev_row["amounts_vals"]:
                        needed = 2 - len(m_vals)
                        if len(prev_row["amounts_vals"]) >= needed:
                             # Tomamos prestado
                             if len(m_vals) == 1: # Tenemos saldo, falta importe
                                  amount_val = prev_row["amounts_vals"][-1]
                                  saldo_val = m_vals[0]
                                  resolved = True
                                  # Marcamos usada la parte de montos solamente? No, toda la fila es safer
                                  used_rows.add(k)
                             elif len(m_vals) == 0:
                                  if len(prev_row["amounts_vals"]) >= 2:
                                      amount_val = prev_row["amounts_vals"][-2]
                                      saldo_val = prev_row["amounts_vals"][-1]
                                      resolved = True
                                      used_rows.add(k)
            
            # Buscar Abajo (Orphan Amount)
            if not resolved:
                next_anchor = anchor_indices[i+1] if i+1 < len(anchor_indices) else len(all_rows)
                for k in range(idx + 1, next_anchor):
                    if k in used_rows: continue
                    next_row = all_rows[k]
                    
                    if next_row["amounts_vals"]:
                         # Check text Mismatch (Prevent stealing next transaction amount)
                         txt_down = " ".join(next_row["desc_words"]).upper()
                         start_kws = ["TEF ", "TRANSF", "COMPENSACION", "DEBIN", "SUELDO", "PAGO SERV", "CABLEVISION", "IMP. AFIP", "IMP.AFIP"]
                         
                         found_kw_down = next((kw for kw in start_kws if kw in txt_down), None)
                         current_anchor_txt = " ".join(row["desc_words"]).upper()
                         
                         # Si la línea de abajo tiene keyword fuerte y yo no la tengo -> SKIP
                         if found_kw_down:
                             if found_kw_down not in current_anchor_txt:
                                 continue

                         if len(next_row["amounts_vals"]) >= 2:
                             amount_val = next_row["amounts_vals"][-2]
                             saldo_val = next_row["amounts_vals"][-1]
                             resolved = True
                             used_rows.add(k)
                             break
                        # Si encontramos solo 1 monto abajo y teniamos 1, match?
                         elif len(next_row["amounts_vals"]) == 1 and len(m_vals) == 1:
                             amount_val = next_row["amounts_vals"][0]
                             saldo_val = m_vals[0]
                             resolved = True
                             used_rows.add(k)
                             break
            
            # --- Resolución Descripción ---
            desc_tokens = []
            
            # 1. Inline Description
            if row["desc_words"]:
                desc_tokens.append(" ".join(row["desc_words"]))
                
            # 2. Prefix (Mirar Arriba)
            # Buscamos texto que pertenezca al inicio de esta transacción pero quedó colgado arriba
            k = idx - 1
            prev_anc_real = anchor_indices[i-1] if i > 0 else -1
            prefix_buff = []
            while k > prev_anc_real:
                if k in used_rows: 
                    k -= 1
                    continue
                
                r_scan = all_rows[k]
                txt = " ".join(r_scan["desc_words"])
                if not txt: 
                    k -= 1
                    continue

                # Heurística Prefix:
                # Aceptamos si es la linea INMEDIATA superior
                # O si tiene keyword fuerte de inicio
                is_immediate = (k == idx - 1)
                has_start_keyword = any(kw in txt.upper() for kw in ["TRANSF", "COMPENSACION", "DEBIN", "SUELDO", "PAGO SERV", "CABLEVISION"])
                
                # Stop conditions (pertenece a la anterior)
                is_prev_tail = False
                if re.match(r"^\d{9,}$", txt) or "VARIOS" in txt.upper(): # CUIT solo o VARIOS suele ser fin
                     is_prev_tail = True
                
                if is_prev_tail and not has_start_keyword:
                    break
                
                if is_immediate or has_start_keyword:
                    prefix_buff.insert(0, txt)
                    used_rows.add(k)
                else:
                    # Si hay un hueco (linea vacia o no aceptada), dejamos de buscar hacia arriba
                    break
                k -= 1
            
            # 3. Suffix (Mirar Abajo)
            # Generalmente el texto fluye hacia abajo. Absorbemos todo hasta el siguiente anchor.
            k = idx + 1
            next_anc_real = anchor_indices[i+1] if i+1 < len(anchor_indices) else len(all_rows)
            suffix_buff = []
            while k < next_anc_real:
                if k in used_rows: 
                    k += 1
                    continue
                
                r_scan = all_rows[k]
                txt = " ".join(r_scan["desc_words"])
                if not txt:
                    k += 1
                    continue
                    
                # Filtro Headers archivo por si acaso (aunque ya filtramos al inicio, el cleaner es por linea raw)
                if any(x in txt for x in ["Fecha Descripción", "Últimos movimientos", "Número de cuenta"]):
                    break

                # Ya NO cortamos por keywords como TEF o TRANSF, porque suelen ser parte de este bloque (linea 2 o 3)
                # Solo paramos si encontramos algo que CLARAMENTE es el prefix del siguiente?
                # Como el prefix scan del siguiente loop va a reclamar lo suyo (usando used_rows? no, no hemos llegado ahi),
                # hay riesgo de robarnos el prefix del siguiente.
                # Pero en PDF bancarios, el prefix del siguiente suele estar pegado al siguiente.
                # Si estamos lejos (k < next_anc_real - 2), safe.
                
                # Heuristica simple: Absorbemos todo. 
                # Salvo que parezca un monto huérfano (ya procesado, estaria en used_rows).
                
                suffix_buff.append(txt)
                used_rows.add(k)
                k += 1
            
            full_desc = " ".join(prefix_buff + desc_tokens + suffix_buff)
            cleaned_desc = clean_for_excel(full_desc)
             # Limpieza Ref numerica inicial
            match_ref = re.match(r"^(\d{5,})\s+(.*)", cleaned_desc)
            if match_ref: cleaned_desc = match_ref.group(2)

            transactions.append({
                "Fecha": fecha,
                "Descripcion": cleaned_desc,
                "Importe": amount_val,
                "Saldo": saldo_val
            })

        if not transactions:
            st.warning("No se encontraron movimientos")
            return None

        # Ordenar Cronológico
        df = pd.DataFrame(transactions)
        df["FechaDt"] = pd.to_datetime(df["Fecha"], format="%d/%m/%Y")
        # Macro F3 viene descendente
        df = df.iloc[::-1].reset_index(drop=True)
        
        # Saldos Inicial y Final
        saldo_inicial = 0.0
        saldo_final = 0.0
        if not df.empty:
            saldo_final = df.iloc[-1]["Saldo"]
            saldo_inicial = df.iloc[0]["Saldo"] - df.iloc[0]["Importe"]

        # Periodo
        if not df.empty:
            fecha_min = df["FechaDt"].min().strftime("%d/%m/%Y")
            fecha_max = df["FechaDt"].max().strftime("%d/%m/%Y")
            periodo = f"{fecha_min} al {fecha_max}"

        # 3. Generar Excel (Codigo Reutilizado)
        output = io.BytesIO()
        wb = Workbook()
        ws = wb.active
        ws.title = "Reporte Macro"
        ws.sheet_view.showGridLines = False
        
        color_bg_main = "003366" 
        color_txt_main = "FFFFFF"
        
        thin_border = Border(left=Side(style='thin', color="A6A6A6"), 
                            right=Side(style='thin', color="A6A6A6"), 
                            top=Side(style='thin', color="A6A6A6"), 
                            bottom=Side(style='thin', color="A6A6A6"))
                            
        fill_head_deb = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
        fill_col_deb = PatternFill(start_color="F2DCDB", end_color="F2DCDB", fill_type="solid")
        fill_row_deb = PatternFill(start_color="FDE9D9", end_color="FDE9D9", fill_type="solid")

        fill_head_cred = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")
        fill_col_cred = PatternFill(start_color="EBF1DE", end_color="EBF1DE", fill_type="solid")
        fill_row_cred = PatternFill(start_color="F2F9F1", end_color="F2F9F1", fill_type="solid")

        creditos = df[df["Importe"] > 0].copy()
        debitos = df[df["Importe"] < 0].copy()
        debitos["Importe"] = debitos["Importe"].abs()

        # Header
        ws.merge_cells("A1:G1")
        tit = ws["A1"]
        tit.value = f"REPORTE MACRO - CTA {clean_for_excel(cuenta)}"
        tit.font = Font(size=14, bold=True, color=color_txt_main)
        tit.fill = PatternFill(start_color=color_bg_main, end_color=color_bg_main, fill_type="solid")
        tit.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 25

        # Metadata
        ws["A3"] = "SALDO INICIAL"
        ws["A3"].font = Font(bold=True, size=10, color="666666")
        ws["B3"] = saldo_inicial
        ws["B3"].number_format = '"$ "#,##0.00'
        ws["B3"].font = Font(bold=True, size=11)
        ws["B3"].border = Border(bottom=Side(style='thin', color="DDDDDD"))

        ws["A4"] = "SALDO FINAL"
        ws["A4"].font = Font(bold=True, size=10, color="666666")
        ws["B4"] = saldo_final
        ws["B4"].number_format = '"$ "#,##0.00'
        ws["B4"].font = Font(bold=True, size=11)
        ws["B4"].border = Border(bottom=Side(style='thin', color="DDDDDD"))

        ws["D3"] = "TITULAR"
        ws["D3"].alignment = Alignment(horizontal='right')
        ws["D3"].font = Font(bold=True, color="666666", size=10)
        ws["E3"] = clean_for_excel(titular)
        ws["E3"].font = Font(bold=True, size=11)
        ws["E3"].alignment = Alignment(horizontal='center')
        ws.merge_cells("E3:G3")
        for c in ["E","F","G"]: ws[f"{c}3"].border = Border(bottom=Side(style='thin', color="DDDDDD"))

        ws["D4"] = "PERÍODO"
        ws["D4"].alignment = Alignment(horizontal='right')
        ws["D4"].font = Font(bold=True, color="666666", size=10)
        ws["E4"] = clean_for_excel(periodo)
        ws["E4"].font = Font(bold=True, size=11)
        ws["E4"].alignment = Alignment(horizontal='center')
        ws.merge_cells("E4:G4")
        for c in ["E","F","G"]: ws[f"{c}4"].border = Border(bottom=Side(style='thin', color="DDDDDD"))
        
        # Tablas
        fila_inicio = 10
        f_header = fila_inicio
        
        ws.merge_cells(f"A{f_header}:C{f_header}")
        ws[f"A{f_header}"] = "CRÉDITOS" 
        ws[f"A{f_header}"].fill = fill_head_cred
        ws[f"A{f_header}"].font = Font(bold=True, color="FFFFFF")
        ws[f"A{f_header}"].alignment = Alignment(horizontal='center')
        ws[f"A{f_header}"].border = thin_border
        
        ws.merge_cells(f"E{f_header}:G{f_header}")
        ws[f"E{f_header}"] = "DÉBITOS" 
        ws[f"E{f_header}"].fill = fill_head_deb
        ws[f"E{f_header}"].font = Font(bold=True, color="FFFFFF")
        ws[f"E{f_header}"].alignment = Alignment(horizontal='center')
        ws[f"E{f_header}"].border = thin_border

        headers = ["Fecha", "Descripción", "Importe"]
        cols_cred = ["A", "B", "C"]
        cols_deb = ["E", "F", "G"]
        f_sub = f_header + 1
        
        for i, h in enumerate(headers):
            c = ws[f"{cols_cred[i]}{f_sub}"]
            c.value = h
            c.fill = fill_col_cred
            c.font = Font(bold=True)
            c.alignment = Alignment(horizontal='center')
            c.border = thin_border

            d = ws[f"{cols_deb[i]}{f_sub}"]
            d.value = h
            d.fill = fill_col_deb
            d.font = Font(bold=True)
            d.alignment = Alignment(horizontal='center')
            d.border = thin_border

        fila_a_llenar = f_sub + 1
        
        # Creditos
        f_c = fila_a_llenar
        if creditos.empty:
            ws.merge_cells(f"A{f_c}:C{f_c}")
            ws[f"A{f_c}"] = "SIN MOVIMIENTOS"
            ws[f"A{f_c}"].border = thin_border
            f_c += 1
        else:
            start_c = f_c
            for _, r in creditos.iterrows():
                ws[f"A{f_c}"] = r["Fecha"]
                ws[f"A{f_c}"].fill = fill_row_cred
                ws[f"A{f_c}"].border = thin_border
                ws[f"A{f_c}"].alignment = Alignment(horizontal='center')
                ws[f"B{f_c}"] = r["Descripcion"]
                ws[f"B{f_c}"].fill = fill_row_cred
                ws[f"B{f_c}"].border = thin_border
                ws[f"C{f_c}"] = r["Importe"]
                ws[f"C{f_c}"].number_format = '"$ "#,##0.00'
                ws[f"C{f_c}"].fill = fill_row_cred
                ws[f"C{f_c}"].border = thin_border
                f_c += 1
            ws.merge_cells(f"A{f_c}:B{f_c}")
            ws[f"A{f_c}"] = "TOTAL CRÉDITOS"
            ws[f"A{f_c}"].font = Font(bold=True)
            ws[f"A{f_c}"].alignment = Alignment(horizontal='right')
            ws[f"C{f_c}"] = f"=SUM(C{start_c}:C{f_c-1})"
            ws[f"C{f_c}"].font = Font(bold=True)
            ws[f"C{f_c}"].number_format = '"$ "#,##0.00'
            f_c += 1

        # Debitos
        f_d = fila_a_llenar
        if debitos.empty:
            ws.merge_cells(f"E{f_d}:G{f_d}")
            ws[f"E{f_d}"] = "SIN MOVIMIENTOS"
            ws[f"E{f_d}"].border = thin_border
            f_d += 1
        else:
            start_d = f_d
            for _, r in debitos.iterrows():
                ws[f"E{f_d}"] = r["Fecha"]
                ws[f"E{f_d}"].fill = fill_row_deb
                ws[f"E{f_d}"].border = thin_border
                ws[f"E{f_d}"].alignment = Alignment(horizontal='center')
                ws[f"F{f_d}"] = r["Descripcion"]
                ws[f"F{f_d}"].fill = fill_row_deb
                ws[f"F{f_d}"].border = thin_border
                ws[f"G{f_d}"] = r["Importe"]
                ws[f"G{f_d}"].number_format = '"$ "#,##0.00'
                ws[f"G{f_d}"].fill = fill_row_deb
                ws[f"G{f_d}"].border = thin_border
                f_d += 1
            ws.merge_cells(f"E{f_d}:F{f_d}")
            ws[f"E{f_d}"] = "TOTAL DÉBITOS"
            ws[f"E{f_d}"].font = Font(bold=True)
            ws[f"E{f_d}"].alignment = Alignment(horizontal='right')
            ws[f"G{f_d}"] = f"=SUM(G{start_d}:G{f_d-1})"
            ws[f"G{f_d}"].font = Font(bold=True)
            ws[f"G{f_d}"].number_format = '"$ "#,##0.00'
            f_d += 1

        # Control
        ws["D6"] = "CONTROL DE SALDOS"
        ws["D6"].font = Font(bold=True, size=10, color="666666")
        ws["D6"].alignment = Alignment(horizontal='center')
        
        ref_tot_c = f"C{f_c-1}" if not creditos.empty else "0"
        ref_tot_d = f"G{f_d-1}" if not debitos.empty else "0"
        ws["D7"] = f"=ROUND(B3+{ref_tot_c}-{ref_tot_d}-B4, 2)"
        ws["D7"].number_format = '"$ "#,##0.00'
        ws["D7"].font = Font(bold=True)
        ws["D7"].alignment = Alignment(horizontal='center')
        ws["D7"].border = thin_border
        
        red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
        red_font = Font(color='9C0006', bold=True)
        ws.conditional_formatting.add('D7', CellIsRule(operator='notEqual', formula=['0'], stopIfTrue=True, fill=red_fill, font=red_font))

        # Anchos
        ws.column_dimensions["A"].width = 12
        ws.column_dimensions["B"].width = 40
        ws.column_dimensions["C"].width = 18
        ws.column_dimensions["D"].width = 25
        ws.column_dimensions["E"].width = 12
        ws.column_dimensions["F"].width = 40
        ws.column_dimensions["G"].width = 18

        wb.save(output)
        output.seek(0)
        return output.getvalue()
        
    except Exception as e:
        import traceback
        st.error(f"Error al procesar: {e}")
        print(traceback.format_exc())
        return None
