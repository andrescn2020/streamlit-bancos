import streamlit as st
import PyPDF2
import re
import pandas as pd
import io


def procesar_provincia(archivo_pdf):
    """Procesa archivos PDF del banco Provincia"""
    st.info("Procesando archivo del banco Provincia...")

    try:
        # Reinicializar el archivo para lectura
        archivo_pdf.seek(0)

        saldo_inicial = None
        saldo_final = None

        # Abrir y leer el archivo PDF
        with io.BytesIO(archivo_pdf.read()) as pdf_file:
            reader = PyPDF2.PdfReader(pdf_file)
            texto = "".join(page.extract_text() + "\n" for page in reader.pages)
            texto = texto.splitlines()

        # Delimitar las líneas de movimientos
        inicio = next(
            (i for i, line in enumerate(texto) if "SALDO ANTERIOR" in line), None
        )
        fin = next(
            (i for i, line in enumerate(texto) if "Todas las comisiones" in line), None
        )

        if inicio is None or fin is None:
            st.error(
                "No se encontraron las secciones 'SALDO ANTERIOR' o 'Todas las comisiones' en el PDF"
            )
            return None

        movimientos_extraidos = texto[inicio:fin]

        # Variables para acumular movimientos y procesarlos
        movimientos = []
        saldo_anterior = None
        linea_actual = ""

        # Patrón que busca los movimientos
        patron_movimiento = re.compile(
            r"^(\d{2}/\d{2}/\d{4})\s+(.*?)\s+(\d{2}-\d{2})\s+([-+]?\d+\.\d{2})$"
        )

        for linea in movimientos_extraidos:
            if "SALDO ANTERIOR" in linea:
                match = re.search(r"SALDO ANTERIOR\s+([-+]?\d+\.\d{2})", linea)
                if match:
                    saldo_anterior = float(match.group(1))
                    saldo_inicial = float(match.group(1))
                linea_actual = linea.strip()
                continue

            if re.match(r"^\d{2}/\d{2}/\d{4}", linea):
                if linea_actual:
                    movimiento_str = linea_actual.strip()
                    m = patron_movimiento.match(movimiento_str)
                    if m and saldo_anterior is not None:
                        fecha = m.group(1)
                        descripcion = m.group(2).strip()
                        saldo_actual = float(m.group(4))
                        importe = saldo_actual - saldo_anterior
                        movimientos.append((fecha, descripcion, importe))
                        saldo_anterior = saldo_actual
                    else:
                        st.warning(
                            f"No se pudo procesar el movimiento: {movimiento_str}"
                        )
                linea_actual = linea.strip()
            else:
                linea_actual += " " + linea.strip()

        # Procesar el último movimiento
        if linea_actual.strip():
            movimiento_str = linea_actual.strip()
            m = patron_movimiento.match(movimiento_str)
            if m and saldo_anterior is not None:
                fecha = m.group(1)
                descripcion = m.group(2).strip()
                saldo_actual = float(m.group(4))
                importe = saldo_actual - saldo_anterior
                movimientos.append((fecha, descripcion, importe))
                saldo_anterior = saldo_actual
                saldo_final = saldo_actual

        if not movimientos:
            st.warning("No se encontraron movimientos en el PDF")
            return None

        st.success(f"Se procesaron {len(movimientos)} movimientos")

        # Generar Excel
        df = pd.DataFrame(movimientos, columns=["Fecha", "Descripcion", "Importe"])
        df_debitos = df[df["Importe"] < 0].copy()
        df_debitos["Importe"] = df_debitos["Importe"].abs()
        df_creditos = df[df["Importe"] > 0].copy()

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
