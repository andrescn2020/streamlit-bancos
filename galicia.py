import streamlit as st
import io
import PyPDF2
import re
import pandas as pd


def procesar_galicia(archivo_pdf):
    """Procesa archivos PDF del banco Galicia"""
    st.info("Procesando archivo del banco Galicia...")

    try:
        # Reinicializar el archivo para lectura
        archivo_pdf.seek(0)

        # Abrir el PDF usando PyPDF2 con io.BytesIO
        with io.BytesIO(archivo_pdf.read()) as pdf_file:
            reader = PyPDF2.PdfReader(pdf_file)
            texto = "".join(page.extract_text() + "\n" for page in reader.pages)
            texto = texto.splitlines()

        # Eliminar líneas vacías y espacios extra
        texto = [line.strip() for line in texto if line.strip()]

        # Inicialización de variables para saldos
        saldo_cuenta = 0
        saldo_inicial = 0
        saldo_final = 0

        # Procesar cada línea para extraer saldos a partir de la palabra "Saldos"
        for lineas in texto:
            if "Saldos" in lineas:
                # Buscar todos los valores en formato $X,XX
                patron = r"([+-]?\$\s*\d{1,3}(?:\.\d{3})*,\d{2})"
                valores = re.findall(patron, lineas)
                if len(valores) >= 2:
                    valor_medio = valores[
                        1
                    ]  # El segundo valor, por ejemplo: "$59.711,26"
                    valor_inicio = valores[0]
                    # Limpieza y conversión a float
                    # 1. Eliminar el símbolo de dólar y espacios
                    valor_inicio = valor_inicio.replace("$", "").strip()
                    valor_limpio = valor_medio.replace("$", "").strip()
                    # 2. Determinar el signo y removerlo
                    signo_medio = -1 if valor_limpio.startswith("-") else 1
                    signo_inicio = -1 if valor_inicio.startswith("-") else 1
                    valor_limpio = valor_limpio.lstrip("-").strip()
                    valor_inicio = valor_inicio.lstrip("-").strip()
                    # 3. Eliminar separadores de miles (puntos)
                    valor_limpio = valor_limpio.replace(".", "")
                    valor_inicio = valor_inicio.replace(".", "")
                    # 4. Reemplazar la coma decimal por punto
                    valor_limpio = valor_limpio.replace(",", ".")
                    valor_inicio = valor_inicio.replace(",", ".")

                    # Convertir a float y aplicar el signo
                    saldo_final = signo_inicio * float(valor_inicio)
                    valor_float = signo_medio * float(valor_limpio)

                    saldo_inicial = valor_float
                    saldo_cuenta = valor_float
            # Otra forma de extraer saldos, según el formato de la línea
            if re.search(
                r"\$\d{1,3}(?:\.\d{3})*(,\d{2})?-\$\d{1,3}(?:\.\d{3})*(,\d{2})?-Saldos",
                lineas,
            ):
                partes = lineas.split("$")
                if "-" in partes[2]:
                    saldo_limpio = partes[2].split("-")
                    saldo_inicial = float(
                        saldo_limpio[0].replace(".", "").replace(",", ".")
                    )
                    saldo_inicial = saldo_inicial * -1
                else:
                    saldo_inicial = float(partes[2].replace(".", "").replace(",", "."))
                if "-" in partes[1]:
                    saldo_final = float(
                        partes[1].replace("-", "").replace(".", "").replace(",", ".")
                    )
                    saldo_final = saldo_final * -1
                else:
                    saldo_final = float(
                        partes[1].replace("-", "").replace(".", "").replace(",", ".")
                    )

        # Buscar el índice de la línea que contiene "Movimientos"
        inicio = next(
            (i for i, line in enumerate(texto) if "Movimientos" in line), None
        )
        # Buscar el índice de la línea que contiene "Total"
        fin = next((i for i, line in enumerate(texto) if "Total" in line), None)

        if inicio is None or fin is None:
            st.error(
                "No se encontraron las secciones 'Movimientos' o 'Total' en el PDF"
            )
            return None

        movimientos_extraidos = texto[inicio + 1 : fin]

        # Unir las líneas que corresponden a un mismo movimiento
        movimientos = []
        linea_actual = ""
        for linea in movimientos_extraidos:
            # Si la línea inicia con fecha (formato dd/mm/aa) se considera nueva línea
            if re.match(r"\d{2}/\d{2}/\d{2}", linea):
                if linea_actual:
                    movimientos.append(linea_actual.strip())
                linea_actual = linea
            else:
                linea_actual += " " + linea

        if linea_actual:
            movimientos.append(linea_actual.strip())

        # Procesar movimientos para extraer fecha, descripción e importe
        movimientos_procesados = []
        for linea in movimientos[1:]:  # Se ignora la primera línea (encabezado)
            matches = re.findall(r"-?\d{1,3}(?:\.\d{3})*,\d{2}-?", linea)
            match_fecha = re.match(r"(\d{2}/\d{2}/\d{2})", linea)
            fecha = match_fecha.group(1) if match_fecha else None
            # Remover la fecha para trabajar con el resto de la línea
            linea_sin_fecha = linea[len(fecha) :].strip() if fecha else linea
            # Extraer la descripción (todo lo anterior al primer número decimal)
            descripcion = re.split(r"\d+\.\d+", linea_sin_fecha, maxsplit=1)[0]
            # Limpiar la descripción quitando posibles números restantes
            descripcion_limpia = re.sub(r"-?\d+[\.,]\d+", "", descripcion)

            if matches:
                saldo_str = matches[-1]  # Se toma el último número (saldo)
                saldo = float(
                    saldo_str.replace(".", "").replace(",", ".").replace("-", "")
                )
                if "-" in saldo_str:
                    saldo *= -1
                # Se calcula el importe como la diferencia respecto al saldo inicial
                importe = round(saldo - saldo_inicial, 2)
                saldo_inicial = saldo

                # Limpiar guiones en la descripción, si los hubiera
                if "-" in descripcion_limpia:
                    descripcion_limpia = descripcion_limpia.replace("-", "")

                movimiento = {
                    "Fecha": fecha,
                    "Descripcion": descripcion_limpia,
                    "Importe": importe,
                }
                movimientos_procesados.append(movimiento)

        # Si se procesaron movimientos, se genera el Excel con los datos
        if movimientos_procesados:
            st.success(f"Se procesaron {len(movimientos_procesados)} movimientos")

            df = pd.DataFrame(
                movimientos_procesados, columns=["Fecha", "Descripcion", "Importe"]
            )
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
                worksheet["J3"] = saldo_cuenta
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
        else:
            st.warning("No se encontraron movimientos en el PDF")
            return None

    except Exception as e:
        st.error(f"Error al procesar el archivo: {str(e)}")
        return None
