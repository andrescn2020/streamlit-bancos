import streamlit as st
import io
import PyPDF2
import re
import pandas as pd
import openpyxl
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE


def procesar_santander_rio(archivo_pdf):
    """Procesa archivos PDF de Santander Rio"""
    st.info("Procesando archivo de Santander Rio...")

    try:
        # Reinicializar el archivo para lectura
        archivo_pdf.seek(0)

        saldo_inicial = 0
        saldo_final = 0

        def limpiar_texto_excel(valor):
            if isinstance(valor, str):
                return ILLEGAL_CHARACTERS_RE.sub("", valor)
            return valor

        # Abrir el PDF usando PyPDF2
        reader = PyPDF2.PdfReader(io.BytesIO(archivo_pdf.read()))
        texto = "".join(page.extract_text() + "\n" for page in reader.pages)
        movimientos_en_pesos = texto.splitlines()

        # Buscar el índice de "Movimientos en pesos"
        indice_inicio = None
        for i, linea in enumerate(movimientos_en_pesos):
            if "Movimientos en pesos" in linea:
                indice_inicio = i
                break

        # Buscar el índice de "Detalle impositivo" y extraer saldos
        indice_final = None
        for i, linea in enumerate(movimientos_en_pesos):
            if "Saldo Inicial" in linea:
                matches = re.findall(r"(-?)\$\s?([\d\.]+,\d{2})", linea)
                if matches:
                    signo, ultimo_valor = matches[-1]
                    numero = float(ultimo_valor.replace(".", "").replace(",", "."))
                    if signo == "-":
                        numero *= -1
                    saldo_inicial = numero

            if "Saldo total" in linea:
                valores = re.findall(r"-?\$\s?([\d\.]+,\d{2})", linea)
                if valores:
                    signos = re.findall(r"(-?)\$\s?[\d\.]+,\d{2}", linea)
                    ultimo_valor = valores[-1]
                    signo = signos[-1]
                    numero = float(ultimo_valor.replace(".", "").replace(",", "."))
                    if signo == "-":
                        numero *= -1
                    saldo_final = numero

            if indice_final is None and "Así usaste tu dinero este mes" in linea:
                indice_final = i
                break
            if "Detalle impositivo" in linea:
                indice_final = i
                break

        # Extraer movimientos en dólares (si es necesario)
        try:
            inicio_dolares = movimientos_en_pesos.index("Movimientos en dólares")
        except ValueError:
            inicio_dolares = None
        try:
            fin_dolares = movimientos_en_pesos.index("Detalle impositivo")
        except ValueError:
            try:
                fin_dolares = movimientos_en_pesos.index(
                    "Así usaste tu dinero este mes"
                )
            except ValueError:
                fin_dolares = None

        if inicio_dolares is not None and fin_dolares is not None:
            movimientos_en_dolares = movimientos_en_pesos[
                inicio_dolares + 1 : fin_dolares
            ]
        else:
            movimientos_en_dolares = []

        # Reducir movimientos_en_pesos a la sección deseada
        if indice_inicio is not None and indice_final is not None:
            movimientos_en_pesos = movimientos_en_pesos[
                indice_inicio + 1 : indice_final
            ]
        elif indice_inicio is not None:
            movimientos_en_pesos = movimientos_en_pesos[indice_inicio + 1 :]
        elif indice_final is not None:
            movimientos_en_pesos = movimientos_en_pesos[:indice_final]

        def procesar_movimientos(lineas):
            """
            Procesa los movimientos bancarios dados en una lista de líneas y retorna una lista de tuplas:
            (fecha, descripción, importe)
            """
            movimientos = []
            linea_actual = ""
            for linea in lineas:
                # Detectar líneas que inician con fecha (formato dd/mm/aa)
                if re.match(r"\d{2}/\d{2}/\d{2}", linea):
                    if linea_actual:
                        movimientos.append(linea_actual.strip())
                    linea_actual = linea
                else:
                    linea_actual += " " + linea
            if linea_actual:
                movimientos.append(linea_actual.strip())

            data = []
            for movimiento in movimientos:
                # Omitir movimientos en dólares
                if "U$S" in movimiento:
                    continue

                if "sircreb" in movimiento:
                    patron = r"([+-]?\$\s*[0-9]+(?:\.[0-9]{3})*,[0-9]{2})"

                    # Buscar todos los valores que cumplan el patrón
                    valores = re.findall(patron, movimiento)
                    # Eliminar el símbolo de dólar y espacios
                    valores[1] = valores[1].replace("$", "").strip()
                    # Determinar el signo y eliminarlo para procesar el número
                    signo = -1 if valores[1].startswith("-") else 1
                    valores[1] = valores[1].lstrip("-").strip()
                    # Eliminar espacios internos
                    valores[1] = valores[1].replace(" ", "")
                    # Quitar separador de miles (punto) y reemplazar la coma decimal por punto
                    importe_str = valores[1].replace(".", "").replace(",", ".")
                    importe = signo * float(importe_str)

                    fecha = movimiento[:8]
                    descripcion = "SIRCREB"

                    data.append((fecha, descripcion, importe))
                    continue

                # Caso especial: "Retencion arba"
                if "Retencion arba" in movimiento:
                    fecha = movimiento[:8]
                    movimiento_sin_fecha = movimiento[9:]
                    indice = [
                        movimiento_sin_fecha.rfind("/"),
                        movimiento_sin_fecha.rfind("-"),
                    ]
                    pattern = r"(-?\$ [\d\.,]+)"
                    resultados = []
                    match = re.findall(pattern, movimiento)
                    if len(match) >= 2:
                        numero_del_medio = match[0]
                        numero_limpio = float(
                            numero_del_medio.replace("$", "")
                            .replace(" ", "")
                            .replace(".", "")
                            .replace(",", ".")
                        )
                        resultados.append(numero_limpio)
                        movimiento = (
                            fecha
                            + " "
                            + movimiento_sin_fecha[: indice[0]]
                            + numero_del_medio
                            + match[1]
                        )

                if "Saldo total" in movimiento:
                    fecha = movimiento[:8]
                    movimiento_sin_fecha = movimiento[8:]
                    indice = movimiento_sin_fecha.rfind("Saldo total")
                    movimiento = movimiento_sin_fecha[:indice]
                    indice = movimiento.rfind("$")
                    movimiento = fecha + " " + movimiento[:indice]
                else:
                    fecha = movimiento[:8]
                    movimiento_sin_fecha = movimiento[8:]
                    indice = movimiento_sin_fecha.rfind("$")
                    movimiento = fecha + " " + movimiento_sin_fecha[:indice]

                if "U$S" in movimiento:
                    movimiento = movimiento.replace("U$S", "$").replace("U", "")

                # Expresión regular para extraer fecha, descripción e importe
                match = re.match(
                    r"(\d{2}/\d{2}/\d{2})\s+(.*?)\s+([+-]?\$?\s?\d{1,3}(?:\.\d{3})*,\d{2})",
                    movimiento,
                )
                if match:
                    if "Saldo Inicial" in movimiento:
                        continue
                    else:
                        fecha, descripcion, importe = match.groups()
                        # Si la descripción comienza con números, se eliminan
                        if re.match(r"^\d+", descripcion):
                            descripcion = re.sub(r"^\d+|\d+$", "", descripcion).strip()
                        importe = (
                            importe.replace(" ", "")
                            .replace(".", "")
                            .replace(",", ".")
                            .replace("$", "")
                        )
                        importe = float(importe)
                        data.append((fecha, descripcion, importe))
            return data

        pesos = procesar_movimientos(movimientos_en_pesos)
        df_pesos = pd.DataFrame(pesos, columns=["Fecha", "Descripcion", "Importe"])
        df_debitos_pesos = df_pesos[df_pesos["Importe"] < 0].copy()
        df_debitos_pesos["Importe"] = df_debitos_pesos["Importe"].abs()
        df_creditos_pesos = df_pesos[df_pesos["Importe"] > 0].copy()

        st.success(f"Se procesaron {len(pesos)} movimientos")

        # Crear el archivo Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            # Crear hoja "Movimientos"
            pd.DataFrame().to_excel(writer, sheet_name="Movimientos", index=False)
            workbook = writer.book
            worksheet = writer.sheets["Movimientos"]

            # Escribir saldos en celdas específicas
            worksheet["J2"] = "Saldo Inicial"
            worksheet["J3"] = saldo_inicial
            worksheet["J5"] = "Saldo Final"
            worksheet["J6"] = saldo_final

            worksheet.cell(row=1, column=2).value = "Débitos"
            worksheet.cell(row=1, column=7).value = "Créditos"

            # Encabezados para débito y crédito
            for idx, col in enumerate(df_debitos_pesos.columns, start=1):
                worksheet.cell(row=2, column=idx, value=col)
            for idx, col in enumerate(df_creditos_pesos.columns, start=1):
                worksheet.cell(row=2, column=5 + idx, value=col)

            # Escribir datos de débito
            for r_idx, row in enumerate(
                df_debitos_pesos.itertuples(index=False), start=3
            ):
                for c_idx, value in enumerate(row, start=1):
                    worksheet.cell(
                        row=r_idx, column=c_idx, value=limpiar_texto_excel(value)
                    )

            # Escribir datos de crédito
            for r_idx, row in enumerate(
                df_creditos_pesos.itertuples(index=False), start=3
            ):
                for c_idx, value in enumerate(row, start=1):
                    worksheet.cell(
                        row=r_idx, column=5 + c_idx, value=limpiar_texto_excel(value)
                    )

            total_filas_debitos = len(df_debitos_pesos) + 2
            total_filas_creditos = len(df_creditos_pesos) + 2

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
