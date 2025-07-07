import streamlit as st
import re
import pandas as pd
import PyPDF2
import io


def procesar_supervielle(archivo_pdf):
    """Procesa archivos PDF del banco Supervielle"""
    st.info("Procesando archivo del banco Supervielle...")

    try:
        # Reinicializar el archivo para lectura
        archivo_pdf.seek(0)

        def procesar_pdf(file_bytes):
            reader = PyPDF2.PdfReader(io.BytesIO(file_bytes))
            texto = "".join(page.extract_text() + "\n" for page in reader.pages)
            lineas = texto.splitlines()

            capturar = False
            numero_de_cuenta_temporal = ""
            movimientos = []
            cuentas = []

            for linea in lineas:
                if capturar:
                    linea = linea.strip()
                    if re.match(r"^\d{2}/\d{2}/\d{2}", linea):
                        movimientos.append(linea)

                if "NUMERO DE CUENTA" in linea:
                    capturar = True
                    match = re.search(r"NUMERO DE CUENTA\s+(\d{2}-\d{8}/\d)", linea)
                    if match:
                        cuenta = {}
                        numero_cuenta = match.group(1)
                        cuenta["cuenta"] = numero_cuenta
                        cuentas.append(cuenta)
                        numero_de_cuenta_temporal = numero_cuenta

                if "Saldo del período anterior" in linea:
                    match = re.search(r"([\d\.]+,\d{2})$", linea)
                    if match:
                        importe_str = match.group(1).replace(".", "").replace(",", ".")
                        importe = float(importe_str)
                        resultado = next(
                            (
                                d
                                for d in cuentas
                                if d["cuenta"] == numero_de_cuenta_temporal
                            ),
                            None,
                        )
                        if resultado:
                            resultado["saldo_inicial"] = importe

                if "SALDO PERIODO ACTUAL" in linea:
                    resultado_movimientos = next(
                        (
                            d
                            for d in cuentas
                            if d["cuenta"] == numero_de_cuenta_temporal
                        ),
                        None,
                    )
                    if resultado_movimientos:
                        resultado_movimientos["movimientos"] = movimientos.copy()
                        movimientos = []
                    capturar = False

                    match = re.search(r"([\d\.]+,\d{2})$", linea)
                    if match:
                        importe_str = match.group(1).replace(".", "").replace(",", ".")
                        importe = float(importe_str)
                        resultado = next(
                            (
                                d
                                for d in cuentas
                                if d["cuenta"] == numero_de_cuenta_temporal
                            ),
                            None,
                        )
                        if resultado:
                            resultado["saldo_final"] = importe

            def procesar_movimientos(movimientos_cuenta, saldo_inicial):
                movimientos_limpios = []
                for movimiento in movimientos_cuenta:
                    movimiento_limpio = {}
                    fecha = movimiento[0:8]
                    descripcion = movimiento[9:40].strip()
                    valor_str = movimiento[85:].strip()

                    if "-" in valor_str:
                        saldo_actual = (
                            float(
                                valor_str.replace(".", "")
                                .replace(",", ".")
                                .replace("-", "")
                            )
                            * -1
                        )
                    else:
                        saldo_actual = float(
                            valor_str.replace(".", "").replace(",", ".")
                        )

                    importe_movimiento = saldo_actual - saldo_inicial

                    movimiento_limpio["Fecha"] = fecha
                    movimiento_limpio["Descripcion"] = descripcion
                    movimiento_limpio["Importe"] = round(importe_movimiento, 2)

                    movimientos_limpios.append(movimiento_limpio)
                    saldo_inicial = saldo_actual

                return movimientos_limpios

            return cuentas, procesar_movimientos

        # Procesar el PDF
        cuentas, procesar_movimientos = procesar_pdf(archivo_pdf.read())

        if not cuentas:
            st.warning("No se encontraron cuentas en el PDF")
            return None

        st.success(f"Se encontraron {len(cuentas)} cuenta(s)")

        # Crear el archivo Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for cuenta in cuentas:
                saldo_inicial = cuenta.get("saldo_inicial", 0)
                saldo_final = cuenta.get("saldo_final", 0)
                numero_cuenta = cuenta["cuenta"]
                movimientos_cuenta = cuenta.get("movimientos", [])

                nombre_hoja = numero_cuenta.replace("/", "-")

                if len(movimientos_cuenta) == 0:
                    pd.DataFrame(["No tiene movimientos"]).to_excel(
                        writer, sheet_name=nombre_hoja, index=False, header=False
                    )
                    continue

                if len(movimientos_cuenta) > 0:
                    movimientos_cuenta.pop(0)

                datos = procesar_movimientos(movimientos_cuenta, saldo_inicial)
                df = pd.DataFrame(datos, columns=["Fecha", "Descripcion", "Importe"])

                df_debitos = df[df["Importe"] < 0].copy()
                df_debitos["Importe"] = df_debitos["Importe"].abs()
                df_creditos = df[df["Importe"] > 0].copy()

                pd.DataFrame().to_excel(writer, sheet_name=nombre_hoja, index=False)
                worksheet = writer.book[nombre_hoja]

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

                for r_idx, row in enumerate(
                    df_debitos.itertuples(index=False), start=3
                ):
                    for c_idx, value in enumerate(row, start=1):
                        worksheet.cell(row=r_idx, column=c_idx, value=value)

                for r_idx, row in enumerate(
                    df_creditos.itertuples(index=False), start=3
                ):
                    for c_idx, value in enumerate(row, start=1):
                        worksheet.cell(row=r_idx, column=5 + c_idx, value=value)

                total_filas_debitos = len(df_debitos) + 2
                total_filas_creditos = len(df_creditos) + 2

                if len(df_debitos) > 0:
                    formula_debitos = f"=SUM(C3:C{total_filas_debitos})"
                    worksheet.cell(
                        row=total_filas_debitos + 1, column=3, value=formula_debitos
                    )
                else:
                    worksheet.cell(row=3, column=3, value=0)

                if len(df_creditos) > 0:
                    formula_creditos = f"=SUM(H3:H{total_filas_creditos})"
                    worksheet.cell(
                        row=total_filas_creditos + 1, column=8, value=formula_creditos
                    )
                else:
                    worksheet.cell(row=3, column=8, value=0)

                if len(df_debitos) > 0 or len(df_creditos) > 0:
                    cell_debitos = (
                        f"C{total_filas_debitos + 1}" if len(df_debitos) > 0 else "0"
                    )
                    cell_creditos = (
                        f"H{total_filas_creditos + 1}" if len(df_creditos) > 0 else "0"
                    )
                    formula_control = (
                        f"=ROUND({cell_creditos} - {cell_debitos} + J3 - J6, 2)"
                    )
                else:
                    formula_control = "=ROUND(J3 - J6, 2)"

                worksheet["J9"] = "CONTROL"
                worksheet["J10"] = formula_control

        # Preparar el archivo para descarga
        output.seek(0)
        return output.getvalue()

    except Exception as e:
        st.error(f"Error al procesar el archivo: {str(e)}")
        return None
