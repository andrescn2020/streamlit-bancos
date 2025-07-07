import streamlit as st
import PyPDF2
import pandas as pd
import re
import io


def procesar_macro(archivo_pdf):
    """Procesa archivos PDF del banco Macro"""
    st.info("Procesando archivo del banco Macro...")

    try:
        # Reinicializar el archivo para lectura
        archivo_pdf.seek(0)

        saldo_inicial = None
        saldo_final = None

        def procesar_linea(linea):
            patron = r"(.*?)(-?\d{1,3}(?:\.\d{3})*(?:,\d{2}))$"
            match = re.search(patron, linea)
            if match:
                descripcion = match.group(1).strip()
                importe = match.group(2).replace(".", "").replace(",", ".")
                return (descripcion, float(importe))
            else:
                return (linea.strip(), None)

        patron_fecha = r"\d{2}/\d{2}/\d{4}"

        # Abrir el PDF usando PyPDF2
        reader = PyPDF2.PdfReader(io.BytesIO(archivo_pdf.read()))
        texto = "".join(page.extract_text() + "\n" for page in reader.pages)
        texto = texto.splitlines()

        # Cortar donde diga "Transferencias entre Cuentas"
        for i, linea in enumerate(texto):
            if "Transferencias entre Cuentas" in linea:
                texto = texto[:i]
                break

        texto = texto[20:]  # Saltarse encabezado inicial

        movimientos = []
        for linea in texto:
            if "Saldos Finales" in linea:
                match = re.search(r"\b(\d{1,3}(?:\.\d{3})*,\d{2})\b", linea)
                if match:
                    saldo_final = float(
                        match.group(1).replace(".", "").replace(",", ".")
                    )
            elif "Saldos Anteriores" in linea:
                match = re.search(r"\b(\d{1,3}(?:\.\d{3})*,\d{2})\b", linea)
                if match:
                    saldo_inicial = float(
                        match.group(1).replace(".", "").replace(",", ".")
                    )
            elif re.search(patron_fecha, linea) and not "Saldos" in linea:
                movimientos.append(linea)

        # Procesar movimientos
        resultado = []
        for linea in movimientos:
            if not linea.strip():
                continue
            try:
                fecha, resto = linea.split(" ", 1)
                descripcion, importe = procesar_linea(resto)
                resultado.append((fecha, descripcion, importe))
            except ValueError:
                st.warning(f"Línea con formato inesperado: {linea}")

        if not resultado:
            st.warning("No se encontraron movimientos en el PDF")
            return None

        df = pd.DataFrame(resultado, columns=["Fecha", "Descripcion", "Importe"])

        df_debitos = df[df["Importe"] < 0].copy()
        df_debitos["Importe"] = df_debitos["Importe"].abs()

        df_creditos = df[df["Importe"] > 0].copy()

        st.success(f"Se procesaron {len(resultado)} movimientos")

        # Crear el archivo Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            hoja = "Movimientos"
            pd.DataFrame().to_excel(writer, sheet_name=hoja, index=False)

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
