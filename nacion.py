import streamlit as st
import pdfplumber
import re
import pandas as pd
import io


def procesar_nacion(archivo_pdf):
    """Procesa archivos PDF del banco Nación"""
    st.info("Procesando archivo del banco Nación...")

    try:
        # Reinicializar el archivo para lectura
        archivo_pdf.seek(0)

        # Expresión regular para buscar una fecha en formato dd/mm/yyyy
        patron_fecha = r"\d{2}/\d{2}/\d{4}"

        with pdfplumber.open(io.BytesIO(archivo_pdf.read())) as pdf:
            texto = ""
            for page in pdf.pages:
                texto += page.extract_text()

            texto = texto.splitlines()

        inicio = next(
            (i for i, line in enumerate(texto) if "SALDO ANTERIOR" in line), None
        )
        fin = next((i for i, line in enumerate(texto) if "SALDO FINAL" in line), None)

        if inicio is None or fin is None:
            st.error(
                "No se encontraron las secciones 'SALDO ANTERIOR' o 'SALDO FINAL' en el PDF"
            )
            return None

        movimientos_extraidos = texto[inicio - 1 : fin + 1]

        transactions = []
        previous_balance = None
        saldo_inicial = None
        saldo_final = None

        for i, line in enumerate(movimientos_extraidos):
            parts = line.split()
            if "SALDO ANTERIOR" in line:
                try:
                    previous_balance = float(
                        parts[-1].replace(".", "").replace(",", ".")
                    )
                    saldo_inicial = float(parts[-1].replace(".", "").replace(",", "."))
                except ValueError:
                    st.warning(f"Error procesando la línea de saldo anterior: {line}")
                continue
            if "SALDO FINAL" in line:
                match = re.search(r"(\d{1,3}(?:\.\d{3})*,\d{2})", line)
                if match:
                    saldo_str = match.group(0)
                    saldo_num = float(saldo_str.replace(".", "").replace(",", "."))
                    saldo_final = saldo_num

            if "FECHA MOVIMIENTOS" in line:
                continue

            if len(parts) < 3:
                st.warning(f"Línea inválida: {line}")
                continue

            date = parts[0]
            description = " ".join(parts[1:-3])
            amount_str = parts[-2]
            balance_str = parts[-1]

            try:
                amount = float(
                    amount_str.replace(".", "").replace(",", ".").replace("-", "")
                ) * (-1 if "-" in amount_str else 1)
                balance = float(
                    balance_str.replace(".", "").replace(",", ".").replace("-", "")
                ) * (-1 if "-" in balance_str else 1)
            except ValueError:
                st.warning(f"No se pudo procesar la línea: {line}")
                continue

            if previous_balance is not None:
                if balance < previous_balance:
                    transaction_type = "Débito"
                    amount = -abs(amount)
                elif balance > previous_balance:
                    transaction_type = "Crédito"
                    amount = abs(amount)
                else:
                    transaction_type = "Neutral"
            else:
                transaction_type = "Desconocido"

            transactions.append((date, description, amount))
            previous_balance = balance

        if not transactions:
            st.warning("No se encontraron movimientos en el PDF")
            return None

        df = pd.DataFrame(transactions, columns=["Fecha", "Descripcion", "Importe"])

        df_debitos = df[df["Importe"] < 0].copy()
        df_debitos["Importe"] = df_debitos["Importe"].abs()

        df_creditos = df[df["Importe"] > 0].copy()

        st.success(f"Se procesaron {len(transactions)} movimientos")

        # Crear el archivo Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            hoja = "Movimientos"
            df_vacia = pd.DataFrame()
            df_vacia.to_excel(writer, sheet_name=hoja, index=False)

            workbook = writer.book
            worksheet = writer.sheets[hoja]

            worksheet["J2"] = "Saldo Inicial"
            worksheet["J3"] = saldo_inicial
            worksheet["J5"] = "Saldo Final"
            worksheet["J6"] = saldo_final

            worksheet.cell(row=1, column=2).value = "Débitos"
            worksheet.cell(row=1, column=7).value = "Créditos"

            for idx, col in enumerate(df_debitos.columns, start=1):
                worksheet.cell(row=2, column=idx, value=col)
            for idx, col in enumerate(df_creditos.columns, start=1):
                worksheet.cell(row=2, column=5 + idx, value=col)

            for r_idx, row in enumerate(df_debitos.itertuples(index=False), start=3):
                for c_idx, value in enumerate(row, start=1):
                    worksheet.cell(row=r_idx, column=c_idx, value=value)

            for r_idx, row in enumerate(df_creditos.itertuples(index=False), start=3):
                for c_idx, value in enumerate(row, start=1):
                    worksheet.cell(row=r_idx, column=5 + c_idx, value=value)

            total_filas_debitos = len(df_debitos) + 2
            total_filas_creditos = len(df_creditos) + 2

            worksheet.cell(row=total_filas_debitos + 1, column=3).value = (
                f"=SUM(C3:C{total_filas_debitos})"
            )
            worksheet.cell(row=total_filas_creditos + 1, column=8).value = (
                f"=SUM(H3:H{total_filas_creditos})"
            )

            worksheet["J9"] = "CONTROL"
            worksheet["J10"] = (
                f"=ROUND(SUM(H{total_filas_creditos + 1}-C{total_filas_debitos + 1}+J3-J6), 2)"
            )

        # Preparar el archivo para descarga
        output.seek(0)
        return output.getvalue()

    except Exception as e:
        st.error(f"Error al procesar el archivo: {str(e)}")
        return None
