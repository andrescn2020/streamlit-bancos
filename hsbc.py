import streamlit as st
import PyPDF2
import re
import pandas as pd
import io

# Configurar la página para usar todo el ancho disponible
st.set_page_config(layout="wide")


def limpiar_nombre_hoja(nombre):
    """Limpia el nombre para que sea válido como nombre de hoja de Excel"""
    # Caracteres no permitidos en nombres de hojas de Excel
    caracteres_invalidos = ["\\", "/", "*", "[", "]", ":", "?"]
    nombre_limpio = nombre
    for char in caracteres_invalidos:
        nombre_limpio = nombre_limpio.replace(char, "_")

    # Limitar longitud a 31 caracteres (límite de Excel)
    if len(nombre_limpio) > 31:
        nombre_limpio = nombre_limpio[:31]

    return nombre_limpio


def procesar_hsbc(archivo_pdf):
    """Procesa archivos PDF del banco HSBC - Versión básica para análisis"""
    st.info("Procesando archivo del banco HSBC...")

    try:
        # Reinicializar el archivo para lectura
        archivo_pdf.seek(0)

        # Abrir y leer el archivo PDF
        reader = PyPDF2.PdfReader(io.BytesIO(archivo_pdf.read()))
        texto = "".join(page.extract_text() + "\n" for page in reader.pages)
        lineas = texto.splitlines()

        st.info(
            f"PDF procesado: {len(reader.pages)} páginas, {len(lineas)} líneas de texto"
        )

        # Mostrar el texto completo para análisis
        st.subheader("Texto extraído del PDF:")
        st.text_area("Contenido del PDF:", texto, height=600, max_chars=None)

        # Lista para almacenar las cuentas encontradas
        cuentas = []
        saldo_inicial = None
        saldo_final = None

        st.subheader("Cuentas encontradas:")
        for i, linea in enumerate(lineas):
            try:

                # Debug: mostrar las primeras líneas para ver el formato
                if i < 10:  # Mostrar solo las primeras 10 líneas para debug
                    st.write(f"Línea {i}: {linea}")

                # Verificar si la línea contiene información de cuenta
                if (
                    (
                        linea.strip().startswith("CAJA DE AHORRO")
                        or linea.strip().startswith("CUENTA CORRIENTE")
                    )
                    and not linea.strip().startswith("CAJA DE AHORRO EN")
                    and not linea.strip().startswith("CUENTA CORRIENTE EN")
                    and re.search(
                        r"\d{3}-\d-\d{5}-\d", linea
                    )  # Buscar patrón de número de cuenta
                ):
                    # Acá procesás la línea válida
                    st.info(f"Línea de cuenta encontrada: {linea}")
                    print(f"Línea de cuenta: {linea}")

                    # Extraer información de la línea
                    # Formato: CAJA DE AHORRO u$s    SLEIL  068-8-23173-6 1500018300006882317364               306.39                306.39
                    partes = linea.split()

                    # Buscar el tipo de cuenta y moneda
                    tipo_cuenta = ""
                    moneda = ""
                    numero_cuenta = ""

                    for i, parte in enumerate(partes):
                        if (
                            parte == "CAJA"
                            and i + 1 < len(partes)
                            and partes[i + 1] == "DE"
                            and i + 2 < len(partes)
                            and partes[i + 2] == "AHORRO"
                        ):
                            tipo_cuenta = "CAJA DE AHORRO"
                            if i + 3 < len(partes):
                                moneda = partes[i + 3]
                        elif (
                            parte == "CUENTA"
                            and i + 1 < len(partes)
                            and partes[i + 1] == "CORRIENTE"
                        ):
                            tipo_cuenta = "CUENTA CORRIENTE"
                            if i + 2 < len(partes):
                                moneda = partes[i + 2]

                    # Buscar número de cuenta con regex
                    match_cuenta = re.search(r"(\d{3}-\d-\d{5}-\d)", linea)
                    if match_cuenta:
                        numero_cuenta = match_cuenta.group(1)

                    # Buscar saldos (últimos dos números en la línea)
                    numeros = re.findall(r"\d{1,3}(?:,\d{3})*\.\d{2}", linea)
                    if len(numeros) >= 2:
                        saldo_actual = numeros[-2]
                        saldo_anterior = numeros[-1]
                        st.success(
                            f"Tipo: {tipo_cuenta}, Moneda: {moneda}, Cuenta: {numero_cuenta}"
                        )
                        st.success(
                            f"Saldo actual: {saldo_actual}, Saldo anterior: {saldo_anterior}"
                        )

                    # Crear objeto cuenta
                    if numero_cuenta:
                        # Determinar tipo de cuenta y moneda para el nombre de hoja
                        nombre_hoja = ""
                        if tipo_cuenta == "CUENTA CORRIENTE":
                            if moneda.upper() in ["U$S", "USD", "U$D"]:
                                nombre_hoja = f"CC U$D {numero_cuenta}"
                            else:
                                nombre_hoja = f"CC $ {numero_cuenta}"
                        elif tipo_cuenta == "CAJA DE AHORRO":
                            if moneda.upper() in ["U$S", "USD", "U$D"]:
                                nombre_hoja = f"CA U$D {numero_cuenta}"
                            else:
                                nombre_hoja = f"CA $ {numero_cuenta}"

                        cuenta_obj = {
                            "cuenta": numero_cuenta,
                            "nombre_hoja": nombre_hoja,
                            "movimientos": [],
                            "saldo_inicial": (
                                saldo_anterior if "saldo_anterior" in locals() else None
                            ),
                            "saldo_final": (
                                saldo_actual if "saldo_actual" in locals() else None
                            ),
                        }

                        cuentas.append(cuenta_obj)
                        st.success(f"Cuenta encontrada: {nombre_hoja}")

                # Lógica específica para cuentas corrientes con "SALDO ANTERIOR"
                if "SALDO ANTERIOR" in linea and "CUENTA CORRIENTE" in linea:
                    # Buscar todos los números con coma y punto
                    numeros = re.findall(r"\d{1,3}(?:,\d{3})*\.\d{2}", linea)

                    # Tomar los dos últimos solo si hay al menos 2 elementos
                    if len(numeros) >= 2:
                        ultimos_dos = numeros[-2:]
                        saldo_inicial = ultimos_dos[0]  # Primer número (saldo anterior)
                        saldo_final = ultimos_dos[1]  # Segundo número (saldo final)
                        print(
                            f"Saldo inicial: {saldo_inicial}, Saldo final: {saldo_final}"
                        )
                        st.success(
                            f"Saldo inicial: {saldo_inicial}, Saldo final: {saldo_final}"
                        )

                # Lógica para cuentas corrientes con "CUENTA CORRIENTE EN"
                if "CUENTA CORRIENTE EN" in linea:
                    st.write(linea)

                    # Extraer el número de cuenta usando regex
                    match = re.search(r"NRO\.\s+([\d-]+)", linea)
                    if match:
                        numero_cuenta = match.group(1)

                        # Determinar tipo de cuenta y moneda
                        nombre_hoja = ""
                        if "U$S" in linea:
                            nombre_hoja = f"CC U$D {numero_cuenta}"
                        else:
                            nombre_hoja = f"CC $ {numero_cuenta}"

                        # Crear objeto cuenta
                        cuenta_obj = {
                            "cuenta": numero_cuenta,
                            "nombre_hoja": nombre_hoja,
                            "movimientos": [],  # Por ahora vacío
                            "saldo_inicial": saldo_inicial,
                            "saldo_final": saldo_final,
                        }

                        cuentas.append(cuenta_obj)
                        st.success(f"Cuenta encontrada: {nombre_hoja}")

                # Debug adicional: mostrar líneas que contienen "CORRIENTE" pero no entran al if
                elif "CORRIENTE" in linea and not linea.strip().startswith(
                    "CUENTA CORRIENTE EN"
                ):
                    st.warning(f"Línea con CORRIENTE no procesada: {linea}")
                    print(f"Línea con CORRIENTE no procesada: {linea}")
            except Exception as e:
                st.warning(f"Error procesando línea {i}: {str(e)}")
                continue

        # Crear el archivo Excel con hojas para cada cuenta
        if cuentas:
            try:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    # Crear una hoja para cada cuenta
                    for cuenta in cuentas:
                        # Crear DataFrame vacío para la cuenta
                        df = pd.DataFrame(columns=["Fecha", "Descripcion", "Importe"])

                        # Escribir en Excel con el nombre descriptivo de la cuenta
                        nombre_limpio = limpiar_nombre_hoja(cuenta["nombre_hoja"])
                        df.to_excel(writer, sheet_name=nombre_limpio, index=False)

                        # Obtener la hoja de trabajo para agregar los saldos
                        worksheet = writer.sheets[nombre_limpio]

                        # Agregar saldo inicial en celda J3
                        if cuenta.get("saldo_inicial"):
                            worksheet["J3"] = "Saldo Inicial:"
                            worksheet["K3"] = cuenta["saldo_inicial"]
                            st.info(
                                f"Saldo inicial agregado en J3: {cuenta['saldo_inicial']}"
                            )

                        # Agregar saldo final en celda J7
                        if cuenta.get("saldo_final"):
                            worksheet["J7"] = "Saldo Final:"
                            worksheet["K7"] = cuenta["saldo_final"]
                            st.info(
                                f"Saldo final agregado en J7: {cuenta['saldo_final']}"
                            )

                        st.info(f"Hoja creada: {nombre_limpio}")

                # Preparar el archivo para descarga
                output.seek(0)

                st.success(
                    f"Se crearon {len(cuentas)} hojas de Excel para las cuentas encontradas"
                )
                return output.getvalue()
            except Exception as e:
                st.error(f"Error creando archivo Excel: {str(e)}")
                return None
        else:
            st.warning("No se encontraron cuentas en el PDF")
            return None

    except Exception as e:
        st.error(f"Error al procesar el archivo: {str(e)}")
        import traceback

        st.error(f"Detalles del error: {traceback.format_exc()}")
        return None
