import streamlit as st
import re
import pandas as pd
import PyPDF2
import io


def procesar_icbc(archivo_pdf):
    """Procesa archivos PDF del banco ICBC"""
    st.info("Procesando archivo del banco ICBC...")

    try:
        # Reinicializar el archivo para lectura
        archivo_pdf.seek(0)

        def procesar_pdf(file_bytes):
            # Abrir el PDF usando PyPDF2 a partir de los bytes subidos
            reader = PyPDF2.PdfReader(io.BytesIO(file_bytes))
            texto = "".join(page.extract_text() + "\n" for page in reader.pages)
            lineas = texto.splitlines()

            movimientos = []
            saldo_inicial = 0
            saldo_final = 0

            # Buscar saldos en las líneas
            for linea in lineas:
                if "SALDO FINAL AL" in linea:
                    # Buscar el último número al final de la línea
                    match = re.search(r"(\d{1,3}(?:\.\d{3})*,\d{2})\s*$", linea)
                    if match:
                        saldo_str = match.group(1)
                        saldo_num = float(saldo_str.replace(".", "").replace(",", "."))
                        saldo_final = saldo_num
                elif "SALDO ULTIMO EXTRACTO AL" in linea:
                    match = re.search(r"(\d{1,3}(?:\.\d{3})*,\d{2})", linea)
                    if match:
                        saldo_str = match.group(1)
                        saldo_num = float(saldo_str.replace(".", "").replace(",", "."))
                        saldo_inicial = saldo_num

            # Filtrar las líneas que comienzan con fecha (formato dd-mm-aa)
            lineas_con_fecha = [
                linea for linea in lineas if re.match(r"^\d{1,2}-\d{2}", linea)
            ]

            # Procesar cada línea para extraer fecha, descripción e importe
            for linea in lineas_con_fecha:
                movimiento = {}
                movimiento["Fecha"] = linea[:5]  # Extrae "dd-mm"
                movimiento["Descripcion"] = linea[6:50]
                movimiento["Importe"] = linea[62:]

                # Buscar el número con formato de miles/decimales (con posible signo negativo)
                match = re.search(
                    r"(\d{1,3}(?:\.\d{3})*,\d{2}-?)", movimiento["Importe"]
                )
                if match:
                    importe_str = match.group(1)
                    movimiento["Importe"] = importe_str  # Se guarda el string original

                    # Convertir a float: se eliminan puntos y se reemplaza la coma por punto
                    importe_num = float(
                        importe_str.replace(".", "").replace(",", ".").replace("-", "")
                    )
                    if "-" in importe_str:
                        importe_num *= -1
                    movimiento["Importe"] = importe_num
                else:
                    movimiento["Importe"] = None

                movimientos.append(movimiento)

            return movimientos, saldo_inicial, saldo_final

        # Procesar el PDF usando la función definida
        datos, saldo_inicial, saldo_final = procesar_pdf(archivo_pdf.read())

        if not datos:
            st.warning("No se encontraron movimientos en el PDF")
            return None

        df = pd.DataFrame(datos)

        # Crear DataFrames para Débitos y Créditos
        df_debitos = df[df["Importe"] < 0].copy()
        df_debitos["Importe"] = df_debitos["Importe"].abs()
        df_creditos = df[df["Importe"] > 0].copy()

        st.success(f"Se procesaron {len(datos)} movimientos")

        # Crear el archivo Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            hoja = "Movimientos"
            # Crear una hoja vacía
            pd.DataFrame().to_excel(writer, sheet_name=hoja, index=False)

            workbook = writer.book
            worksheet = writer.sheets[hoja]

            # Escribir saldos
            worksheet["J2"] = "Saldo Inicial"
            worksheet["J3"] = saldo_inicial
            worksheet["J5"] = "Saldo Final"
            worksheet["J6"] = saldo_final

            # Título de secciones
            worksheet.cell(row=1, column=2).value = "Débitos"
            worksheet.cell(row=1, column=7).value = "Créditos"

            # Títulos de columnas para Débitos y Créditos
            for idx, col in enumerate(df_debitos.columns, start=1):
                worksheet.cell(row=2, column=idx, value=col)
            for idx, col in enumerate(df_creditos.columns, start=1):
                worksheet.cell(row=2, column=5 + idx, value=col)

            # Escribir datos de Débitos
            for r_idx, row in enumerate(df_debitos.itertuples(index=False), start=3):
                for c_idx, value in enumerate(row, start=1):
                    worksheet.cell(row=r_idx, column=c_idx, value=value)

            # Escribir datos de Créditos
            for r_idx, row in enumerate(df_creditos.itertuples(index=False), start=3):
                for c_idx, value in enumerate(row, start=1):
                    worksheet.cell(row=r_idx, column=5 + c_idx, value=value)

            # Fórmulas de totales
            total_filas_debitos = len(df_debitos) + 2  # +2 por encabezados
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
