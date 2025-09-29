# 🚀 PROCESADOR DE LOTES JYSK - VERSIÓN SIMPLE

## 📋 INSTRUCCIONES PARA EL CLIENTE:

### Paso 1: Instalar Python (5 minutos)
1. Ir a: https://python.org/downloads/
2. Descargar Python 3.11 para Windows
3. Al instalar, marcar "Add Python to PATH" ✅

### Paso 2: Instalar librerías (2 minutos)
Abrir "Command Prompt" y escribir:
```
pip install pandas openpyxl
```

### Paso 3: Usar el programa
1. Hacer doble clic en `procesar_lotes.py`
2. Seleccionar archivo Excel
3. ¡Listo! Se crea el archivo procesado

## 🔧 Si no funciona el doble clic:
Abrir Command Prompt en la carpeta del archivo y escribir:
```
python procesar_lotes.py
```

## ⚡ ALTERNATIVA: Crear tu propio .exe
Si quieres crear un ejecutable:
```
pip install pyinstaller
pyinstaller --onefile --noconsole procesar_lotes.py
```
El .exe estará en la carpeta `dist/`

---
**Total de instalación: ~7 minutos**
**Funciona en cualquier Windows**
