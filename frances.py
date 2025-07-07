import streamlit as st
import io
import pdfplumber
import re
import pandas as pd


def procesar_bbva_frances(archivo_pdf):
    """Procesa archivos PDF de BBVA Frances"""
    st.info("Procesando archivo de BBVA Frances...")

    try:
        # Leer el PDF usando pdfplumber
        with pdfplumber.open(io.BytesIO(archivo_pdf.read())) as pdf:
            texto = ""
            for page in pdf.pages:
                texto += page.extract_text()  # Extrae el texto de cada página
            texto = texto.splitlines()

        # Buscar los índices de las líneas de inicio y fin para extraer los movimientos
        inicio = next(
            (i for i, line in enumerate(texto) if "Movimientos en cuentas" in line),
            None,
        )
        fin = next(
            (i for i, line in enumerate(texto) if "Transferencias" in line), None
        )

        if inicio is None or fin is None:
            st.error(
                "No se encontraron las secciones 'Movimientos en cuentas' o 'Transferencias' en el PDF"
            )
            return None
        else:
            movimientos_extraidos = texto[inicio + 1 : fin]

            # Expresión regular para identificar las líneas de cuenta ("CA ..." o "CC ...")
            pattern_cuenta = r"^(CA|CC)\s"

            # Lista para almacenar las cuentas
            cuentas = []

            # Recorrer cada línea para identificar el inicio de un bloque de cuenta
            for index, movimiento in enumerate(movimientos_extraidos):
                if re.match(pattern_cuenta, movimiento):
                    corte = movimiento.find("(") - 1
                    cuenta = {
                        "cuenta": movimiento[:corte],
                        "inicio": index,
                        "saldo_inicial": 0,
                        "saldo_final": 0,
                        "fin": 0,
                    }
                    # Buscar el final del bloque (línea que contenga "TOTAL MOVIMIENTOS")
                    for j in range(index, len(movimientos_extraidos)):
                        if "TOTAL MOVIMIENTOS" in movimientos_extraidos[j]:
                            cuenta["fin"] = j
                            cuentas.append(cuenta)
                            break

            if not cuentas:
                st.warning("No se encontraron cuentas en el PDF")
                return None
            else:
                st.success(f"Se encontraron {len(cuentas)} cuenta(s)")

                # Crear el ExcelWriter para escribir el archivo Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    # Recorrer cada cuenta encontrada
                    for cuenta in cuentas:
                        # Expresión regular para extraer fecha, descripción e importe
                        pattern = r"(\d{2}/\d{2})\s([A-Za-z0-9\s\./,\-+]+)\s([-]?\d{1,3}(?:[\.,]\d{3})*(?:[\.,]\d{2}))\s"
                        resultados = []
                        movimientos = movimientos_extraidos[
                            cuenta["inicio"] + 1 : cuenta["fin"]
                        ]
                        for movimiento in movimientos:
                            # Extraer saldo inicial
                            if "SALDO ANTERIOR" in movimiento:
                                matches = re.findall(
                                    r"\d{1,3}(?:\.\d{3})*,\d{2}", movimiento
                                )
                                if matches:
                                    numero_str = (
                                        matches[0].replace(".", "").replace(",", ".")
                                    )
                                    numero = float(numero_str)
                                    cuenta["saldo_inicial"] = numero

                            # Extraer saldo final
                            if "SALDO AL" in movimiento:
                                matches = re.findall(
                                    r"\d{1,3}(?:\.\d{3})*,\d{2}", movimiento
                                )
                                if matches:
                                    numero_str = (
                                        matches[0].replace(".", "").replace(",", ".")
                                    )
                                    numero = float(numero_str)
                                    cuenta["saldo_final"] = numero

                            # Ajustar el movimiento en caso de contener "SIRCREB"
                            if "SIRCREB" in movimiento:
                                inicio_mov = movimiento.index("F:")
                                fin_mov = inicio_mov + 10
                                movimiento = (
                                    movimiento[:inicio_mov] + movimiento[fin_mov:]
                                )

                            match = re.match(pattern, movimiento)
                            if match:
                                fecha = match.group(1)
                                descripcion = match.group(2).strip()
                                importe = match.group(3).replace(",", ".")
                                # Convertir importe a float manejando los puntos de miles
                                importe = round(
                                    float(
                                        importe.replace(".", "", importe.count(".") - 1)
                                    ),
                                    2,
                                )
                                resultados.append((fecha, descripcion, importe))

                        # Si se han extraído movimientos, se crea la hoja correspondiente
                        if len(resultados) > 0:
                            df = pd.DataFrame(
                                resultados, columns=["Fecha", "Descripcion", "Importe"]
                            )
                            df_debitos = df[df["Importe"] < 0].copy()
                            df_debitos["Importe"] = df_debitos["Importe"].abs()
                            df_creditos = df[df["Importe"] > 0].copy()

                            # Sanitizar el nombre de la hoja (evitar caracteres inválidos)
                            nombre_hoja = cuenta["cuenta"].replace("/", "-")
                            pd.DataFrame().to_excel(
                                writer, sheet_name=nombre_hoja, index=False
                            )

                            workbook = writer.book
                            worksheet = writer.sheets[nombre_hoja]

                            # Escribir saldos y encabezados
                            worksheet["J2"] = "Saldo Inicial"
                            worksheet["J3"] = cuenta["saldo_inicial"]
                            worksheet["J5"] = "Saldo Final"
                            worksheet["J6"] = cuenta["saldo_final"]

                            worksheet.cell(row=1, column=2).value = "Débitos"
                            worksheet.cell(row=1, column=7).value = "Créditos"

                            # Escribir encabezados para débito y crédito
                            for idx, col in enumerate(df_debitos.columns, start=1):
                                worksheet.cell(row=2, column=idx, value=col)
                            for idx, col in enumerate(df_creditos.columns, start=1):
                                worksheet.cell(row=2, column=5 + idx, value=col)

                            # Escribir los datos de débito
                            for r_idx, row in enumerate(
                                df_debitos.itertuples(index=False), start=3
                            ):
                                for c_idx, value in enumerate(row, start=1):
                                    worksheet.cell(row=r_idx, column=c_idx, value=value)
                            # Escribir los datos de crédito
                            for r_idx, row in enumerate(
                                df_creditos.itertuples(index=False), start=3
                            ):
                                for c_idx, value in enumerate(row, start=1):
                                    worksheet.cell(
                                        row=r_idx, column=5 + c_idx, value=value
                                    )

                            total_filas_debitos = len(df_debitos) + 2
                            total_filas_creditos = len(df_creditos) + 2

                            worksheet.cell(
                                row=total_filas_debitos + 1, column=3
                            ).value = f"=SUM(C3:C{total_filas_debitos})"
                            worksheet.cell(
                                row=total_filas_creditos + 1, column=8
                            ).value = f"=SUM(H3:H{total_filas_creditos})"

                            worksheet["J9"] = "CONTROL"
                            worksheet["J10"] = (
                                f"=ROUND(SUM(H{total_filas_creditos + 1}-C{total_filas_debitos + 1}+J3-J6), 2)"
                            )
                        else:
                            # Si no hay movimientos, crear una hoja indicando que no hay movimientos
                            nombre_hoja = cuenta["cuenta"].replace("/", "-")
                            sheet = writer.book.create_sheet(nombre_hoja)
                            sheet["A1"] = "No tiene movimientos"

                # Preparar el archivo para descarga
                output.seek(0)
                return output.getvalue()

    except Exception as e:
        st.error(f"Error al procesar el archivo: {str(e)}")
        return None
