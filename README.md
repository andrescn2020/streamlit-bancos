# 🏦 Movimientos Bancos - Streamlit

Herramienta para procesar extractos bancarios PDF y generar reportes Excel.

## Requisitos

- Python 3.10+

## Instalación y ejecución (Windows - PowerShell)

### 1. Habilitar ejecución de scripts en PowerShell (solo la primera vez)

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

### 2. Crear entorno virtual

```powershell
python -m venv venv
```

### 3. Activar entorno virtual

```powershell
.\venv\Scripts\Activate
```

### 4. Instalar dependencias

```powershell
pip install -r requirements.txt
```

### 5. Ejecutar la aplicación

```powershell
streamlit run app.py
```

---

## Uso posterior (ya teniendo el venv creado)

```powershell
.\venv\Scripts\Activate
streamlit run app.py
```
