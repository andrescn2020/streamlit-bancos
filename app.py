import streamlit as st
from frances import procesar_bbva_frances
from santander import procesar_santander_rio
from galicia import procesar_galicia
from icbc import procesar_icbc
from macro import procesar_macro
from nacion import procesar_nacion
from provincia import procesar_provincia
from supervielle import procesar_supervielle
from hsbc import procesar_hsbc
from credicoop import procesar_credicoop
from mercadopago import procesar_mercadopago

# Lista de bancos
bancos = [
    "BBVA Frances",
    "Santander Rio",
    # "Credicoop",
    # "HSBC",
    "ICBC",
    "Supervielle",
    "Galicia",
    "Macro",
    "Nacion",
    "Provincia",
    "MercadoPago",
]


def procesar_banco(banco_seleccionado, archivo_pdf):
    """Función principal que dirige el procesamiento según el banco seleccionado"""
    if banco_seleccionado == "BBVA Frances":
        return procesar_bbva_frances(archivo_pdf)
    elif banco_seleccionado == "Santander Rio":
        return procesar_santander_rio(archivo_pdf)
    elif banco_seleccionado == "Galicia":
        return procesar_galicia(archivo_pdf)
    elif banco_seleccionado == "ICBC":
        return procesar_icbc(archivo_pdf)
    elif banco_seleccionado == "Macro":
        return procesar_macro(archivo_pdf)
    elif banco_seleccionado == "Nacion":
        return procesar_nacion(archivo_pdf)
    elif banco_seleccionado == "Provincia":
        return procesar_provincia(archivo_pdf)
    elif banco_seleccionado == "Supervielle":
        return procesar_supervielle(archivo_pdf)
    elif banco_seleccionado == "HSBC":
        return procesar_hsbc(archivo_pdf)
    elif banco_seleccionado == "Credicoop":
        return procesar_credicoop(archivo_pdf)
    elif banco_seleccionado == "MercadoPago":
        return procesar_mercadopago(archivo_pdf)
    else:
        st.info(f"Lógica para {banco_seleccionado} aún no implementada")
        return None


# Interfaz principal de Streamlit
st.title("Selector de Banco y Subida de PDF")

# Selector de banco
banco_seleccionado = st.selectbox("Selecciona un banco:", bancos)

# Subida de archivo PDF
archivo_pdf = st.file_uploader("Sube un archivo PDF", type=["pdf"])

if archivo_pdf is not None:
    st.success(f"Archivo '{archivo_pdf.name}' subido correctamente.")

    # Procesar el archivo según el banco seleccionado
    resultado = procesar_banco(banco_seleccionado, archivo_pdf)

    if resultado is not None:
        # Determinar el nombre del archivo según el banco
        nombre_archivo = f"{banco_seleccionado}.xlsx"

        st.download_button(
            label="Descargar archivo Excel procesado",
            data=resultado,
            file_name=nombre_archivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
