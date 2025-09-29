# Procesador de Lotes JYSK

Script para procesar archivos Excel de transacciones JYSK con filtrado automático, agrupación y formato profesional.

## 🚀 Descarga del Ejecutable

Ve a la sección [Releases](../../releases) o [Actions](../../actions) para descargar el ejecutable compilado para Windows.

## 📋 Características

- ✅ Elimina transacciones denegadas, reversadas y anuladas
- ✅ Mantiene las devoluciones
- ✅ Colores automáticos por caja
- ✅ Subtotales por lote en negrita
- ✅ Hojas separadas por fecha
- ✅ Ajuste automático de columnas

## 💻 Uso Manual (si prefieres Python)

1. Instalar Python 3.11+: https://python.org
2. Instalar dependencias: `pip install pandas openpyxl`
3. Ejecutar: `python procesar_lotes.py`

## 🔧 Crear tu propio ejecutable

```bash
pip install pyinstaller
pyinstaller --onefile --noconsole --name "ProcesadorLotes" procesar_lotes.py
```

El ejecutable estará en la carpeta `dist/`
