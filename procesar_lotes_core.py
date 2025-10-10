# Backend: procesamiento principal de lotes JYSK
import pandas as pd
import os
from openpyxl.styles import PatternFill, Font
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
    }
    colores_caja = {
        '1-CO': 'FFE6E6',
        '2-CO': 'E6F3FF',
        '3-CO': 'E6FFE6',
        '4-CO': 'FFF2E6',
        '5-CO': 'F0E6FF',
        '6-CO Manual': 'FFFFE6',
        '7-CO Manual': 'F0F0F0'
    }
    df = pd.read_excel(file_descarga)
    df = df[~df['Estado'].isin(['Denegada', 'Reversada'])]
    # Detectar anulaciones sin match (autorización sin par)
    anulaciones_sin_match = []
    if "Autorización" in df.columns and "Operación" in df.columns:
        anulaciones = df[df['Operación'] == 'Anulación']['Autorización'].unique()
        autorizaciones_a_eliminar = []
        for aut in anulaciones:
            if aut and (df['Autorización'] == aut).sum() > 1:
                autorizaciones_a_eliminar.append(aut)
        # Guardar anulaciones sin match
        for aut in anulaciones:
            if aut and (df['Autorización'] == aut).sum() == 1:
                idxs = df[(df['Operación'] == 'Anulación') & (df['Autorización'] == aut)].index.tolist()
                anulaciones_sin_match.extend(idxs)
        if autorizaciones_a_eliminar:
            df = df[~df['Autorización'].isin(autorizaciones_a_eliminar)]
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
            datos_fecha = datos_fecha.sort_values(['Caja_Nombre', 'Lote', 'Tarjeta'])
            columnas_mantener = ['Lote', 'Caja_Nombre', 'Tarjeta', 'Monto_Total', 'Cantidad_Transacciones', 'Codigo_Caja']
            datos_fecha_limpio = datos_fecha[columnas_mantener]
            datos_con_subtotales = []
            for (caja, lote), grupo_lote in datos_fecha_limpio.groupby(['Caja_Nombre', 'Lote']):
                for _, fila in grupo_lote.iterrows():
                    datos_con_subtotales.append(fila.to_dict())
                subtotal_monto = grupo_lote['Monto_Total'].sum()
                subtotal_transacciones = grupo_lote['Cantidad_Transacciones'].sum()
                fila_subtotal = {
                    'Lote': lote,
                    'Caja_Nombre': caja,
                    'Tarjeta': f'SUBTOTAL LOTE {lote}',
                    'Monto_Total': float(subtotal_monto),
                    'Cantidad_Transacciones': int(subtotal_transacciones),
                    'Codigo_Caja': grupo_lote['Codigo_Caja'].iloc[0]
                }
                datos_con_subtotales.append(fila_subtotal)
                fila_vacia = {
                    'Lote': '',
                    'Caja_Nombre': '',
                    'Tarjeta': '',
                    'Monto_Total': '',
                    'Cantidad_Transacciones': '',
                    'Codigo_Caja': ''
                }
                datos_con_subtotales.append(fila_vacia)
            if datos_con_subtotales:
                datos_finales = pd.DataFrame(datos_con_subtotales)
                datos_finales['Monto_Total'] = pd.to_numeric(datos_finales['Monto_Total'], errors='coerce')
                datos_finales['Cantidad_Transacciones'] = pd.to_numeric(datos_finales['Cantidad_Transacciones'], errors='coerce')
            else:
                datos_finales = pd.DataFrame(columns=columnas_mantener)
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
                font_negrita = Font(bold=True)
                for row_idx, (_, row) in enumerate(datos_finales.iterrows(), start=2):
                    caja_nombre = row['Caja_Nombre']
                    tarjeta = row['Tarjeta']
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
                    if isinstance(tarjeta, str) and tarjeta.startswith('SUBTOTAL LOTE'):
                        for col_idx in range(1, len(datos_finales.columns) + 1):
                            cell = worksheet.cell(row=row_idx, column=col_idx)
                            cell.font = font_negrita
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
