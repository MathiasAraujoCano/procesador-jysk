# Backend: procesamiento principal de lotes JYSK
import pandas as pd
import os
from openpyxl.styles import PatternFill
from datetime import datetime

def procesar_lotes(file_descarga):
    caja_nombre_map = {
        'T00003': '1-CO',
        'T00004': '2-CO',
        'T00005': '3-CO',
        'T00006': '4-CO',
        'T00007': '5-CO',
        'T00001': '6-CO Manual',
        'T00002': '7-CO Manual',
        'T00008': '1-PC',
        'T00009': '2-PC',
        'T00010': '3-PC',
        'T00011': '5-PC',
        'T00012': '4-PC',
        'T00013': 'PC-BKP',
        'T00014': 'PC-MANUAL',
        'T00015': 'PC-MANUAL'
    }
    colores_caja = {
        '1-CO': 'FFB7B2',      # Rosa pastel
        '2-CO': 'B2C7FF',      # Azul pastel
        '3-CO': 'B2FFD6',      # Verde menta
        '4-CO': 'FFF7B2',      # Amarillo suave
        '5-CO': 'D6B2FF',      # Lila pastel
        '6-CO Manual': 'FFE6B2', # Naranja claro
        '7-CO Manual': 'B2FFF7', # Turquesa pastel
        '1-PC': 'FFD6F5',      # Rosa lavanda
        '2-PC': 'B2FFF7',      # Turquesa pastel
        '3-PC': 'FFB2E6',      # Rosa fuerte
        '4-PC': 'B2FFD6',      # Verde menta
        '5-PC': 'B2C7FF',      # Azul pastel
        'PC-BKP': 'FFF7B2',    # Amarillo suave
        'PC-MANUAL': 'D6B2FF'  # Lila pastel
    }
    # Orden de las cajas dentro del Excel: primero sucursal CO, luego PC
    orden_cajas = [
        '1-CO', '2-CO', '3-CO', '4-CO', '5-CO', '6-CO Manual', '7-CO Manual',
        '1-PC', '2-PC', '3-PC', '4-PC', '5-PC', 'PC-BKP', 'PC-MANUAL'
    ]
    df = pd.read_excel(file_descarga)
    df = df[~df['Estado'].isin(['Denegada', 'Reversada', 'Error'])]
    # Detectar anulaciones sin match (autorización sin par) - PROCESANDO POR LOTE
    anulaciones_sin_match = []
    indices_a_eliminar = []
    
    if "Autorización" in df.columns and "Operación" in df.columns and "Lote" in df.columns:
        # Procesar cada lote independientemente
        for lote in df['Lote'].unique():
            df_lote = df[df['Lote'] == lote]
            anulaciones_lote = df_lote[df_lote['Operación'] == 'Anulación']['Autorización'].unique()
            
            for aut in anulaciones_lote:
                if aut:
                    # Contar cuántas veces aparece esta autorización EN ESTE LOTE
                    count_en_lote = (df_lote['Autorización'] == aut).sum()
                    
                    if count_en_lote > 1:
                        # Hay par (venta + anulación en el mismo lote) - eliminar ambas
                        indices_lote = df_lote[df_lote['Autorización'] == aut].index.tolist()
                        indices_a_eliminar.extend(indices_lote)
                    else:
                        # Anulación sin par en este lote - marcar para ajuste de monto
                        idx_anulacion = df_lote[(df_lote['Operación'] == 'Anulación') & (df_lote['Autorización'] == aut)].index.tolist()
                        anulaciones_sin_match.extend(idx_anulacion)
        
        # Eliminar las transacciones que se cancelan entre sí (dentro del mismo lote)
        if indices_a_eliminar:
            df = df.drop(indices_a_eliminar)
    df['Tarjeta'] = df['Tarjeta'].replace({'Maestro': 'Mastercard'})
    def mapear_nombre_caja(codigo_caja):
        codigo_str = str(codigo_caja).strip()
        return caja_nombre_map.get(codigo_str, f"Sin mapear: {codigo_str}")
    df['Caja_Nombre'] = df['Caja'].apply(mapear_nombre_caja)
    cajas_con_datos = df[df['Caja_Nombre'].str.contains('Sin mapear') == False]
    cajas_con_datos['Tarjeta'] = cajas_con_datos['Tarjeta'].replace({'MAESTRO': 'MASTERCARD'})
    def calcular_monto_ajustado(row):
        if row.name in anulaciones_sin_match:
            return -abs(row['Monto'])
        if row['Operación'] == 'Devolución':
            return -abs(row['Monto'])
        else:
            return row['Monto']
    color_anulacion_sin_match = 'FFB3B3'  # Rojo suave
    cajas_con_datos = cajas_con_datos.copy()
    cajas_con_datos['Monto_Ajustado'] = cajas_con_datos.apply(calcular_monto_ajustado, axis=1)
    resumen = cajas_con_datos.groupby(['Fecha', 'Lote', 'Caja_Nombre', 'Tarjeta']).agg({
        'Monto_Ajustado': ['sum', 'count'],
        'Operación': lambda x: ', '.join(x.unique()),
        'Estado': lambda x: ', '.join(x.unique()),
        'Caja': 'first'
    }).round(2)
    resumen.columns = ['Monto_Total', 'Cantidad_Transacciones', 'Operaciones', 'Estados', 'Codigo_Caja']
    resumen = resumen.reset_index()
    resumen = resumen.sort_values(['Fecha', 'Lote', 'Caja_Nombre', 'Tarjeta'])
    ahora = datetime.now()
    timestamp = ahora.strftime("%d%m%Y_%H%M")
    output_dir = os.path.dirname(file_descarga)
    output_file = os.path.join(output_dir, f"resumen_{timestamp}.xlsx")
    if os.name == 'nt':
        try:
            os.makedirs(output_dir, exist_ok=True)
        except:
            output_dir = os.path.join(os.path.expanduser("~"), "Desktop")
            output_file = os.path.join(output_dir, f"resumen_{timestamp}.xlsx")
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        fechas_unicas = resumen['Fecha'].unique()
        for fecha in sorted(fechas_unicas):
            datos_fecha = resumen[resumen['Fecha'] == fecha].copy()
            datos_fecha['Orden_Caja'] = datos_fecha['Caja_Nombre'].apply(lambda nombre: orden_cajas.index(nombre) if nombre in orden_cajas else len(orden_cajas))
            datos_fecha = datos_fecha.sort_values(['Orden_Caja', 'Lote', 'Tarjeta'])
            columnas_mantener = ['Lote', 'Caja_Nombre', 'Tarjeta', 'Monto_Total', 'Cantidad_Transacciones', 'Codigo_Caja']
            datos_finales = datos_fecha[columnas_mantener].reset_index(drop=True)
            try:
                if hasattr(fecha, 'strftime'):
                    fecha_str = fecha.strftime("%d-%m-%Y")
                else:
                    if isinstance(fecha, str):
                        for fmt in ["%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%Y/%m/%d"]:
                            try:
                                fecha_obj = datetime.strptime(fecha, fmt)
                                fecha_str = fecha_obj.strftime("%d-%m-%Y")
                                break
                            except:
                                continue
                        else:
                            fecha_str = str(fecha).replace('/', '-').replace(' 00:00:00', '')[:10]
                    else:
                        fecha_str = str(fecha).replace('/', '-').replace(' 00:00:00', '')[:10]
                sheet_name = f"Fecha {fecha_str}"
            except Exception:
                fecha_str = str(fecha).replace('/', '-').replace(' 00:00:00', '').replace(' ', '')[:10]
                sheet_name = f"Fecha {fecha_str}"
            if not datos_finales.empty:
                datos_finales.to_excel(writer, sheet_name=sheet_name, index=False)
                worksheet = writer.sheets[sheet_name]
                for row_idx, (_, row) in enumerate(datos_finales.iterrows(), start=2):
                    caja_nombre = row['Caja_Nombre']
                    # Detectar si es una anulación sin match
                    es_anulacion_sin_match = False
                    # Usamos el index original guardado en anulaciones_sin_match
                    if row_idx-2 < len(datos_finales) and datos_finales.index[row_idx-2] in anulaciones_sin_match:
                        es_anulacion_sin_match = True
                    if es_anulacion_sin_match:
                        fill = PatternFill(start_color=color_anulacion_sin_match, end_color=color_anulacion_sin_match, fill_type="solid")
                        for col_idx in range(1, len(datos_finales.columns) + 1):
                            cell = worksheet.cell(row=row_idx, column=col_idx)
                            cell.fill = fill
                    elif caja_nombre in colores_caja:
                        color_hex = colores_caja[caja_nombre]
                        fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")
                        for col_idx in range(1, len(datos_finales.columns) + 1):
                            cell = worksheet.cell(row=row_idx, column=col_idx)
                            cell.fill = fill
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 30)
                    worksheet.column_dimensions[column_letter].width = adjusted_width
    return {
        'output_file': output_file,
        'resumen': resumen
    }
