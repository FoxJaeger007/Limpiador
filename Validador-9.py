import pandas as pd
import logging
import os
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side

def log_null_and_duplicate_data(df, ws, excel_filename, column_prefixes):
    ws.append(["***** Procesando archivo: {} *****".format(excel_filename)])

    # Registro de datos nulos
    null_data = df.isnull().sum()
    ws.append([])
    ws.append(["Datos nulos por columna:"])

    # Aplicar negrilla a los títulos de las tablas inmediatamente y agregar bordes
    border = Border(left=Side(style='thin'), 
                    right=Side(style='thin'), 
                    top=Side(style='thin'), 
                    bottom=Side(style='thin'))

    for row in ws.iter_rows(min_row=ws.max_row, max_row=ws.max_row, min_col=1, max_col=2):
        for cell in row:
            cell.font = Font(bold=True)
            cell.border = border

    # Detalle de datos nulos con bordes
    for column, count in null_data.items():
        ws.append([column, count])
        for row in ws.iter_rows(min_row=ws.max_row, max_row=ws.max_row, min_col=1, max_col=2):
            for cell in row:
                cell.border = border

    # Ajustar el ancho de la columna A según el contenido de los datos nulos, excluyendo la primera fila
    max_length = 0
    for cell in ws['A'][1:]:  # Excluir la primera fila
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = max_length + 2
    ws.column_dimensions['A'].width = adjusted_width

    # Añadir un espacio antes de cada nueva sección
    ws.append([])
    ws.append([])

    # Encontrar columnas con los prefijos especificados
    columns_with_prefixes = [col for col in df.columns if any(col.startswith(prefix) for prefix in column_prefixes)]
    ws.append(["Columnas con los prefijos '{}'".format(', '.join(column_prefixes))])
    for row in ws.iter_rows(min_row=ws.max_row, max_row=ws.max_row, min_col=1, max_col=1):
        for cell in row:
            cell.font = Font(bold=True)
            cell.border = border

    ws.append(columns_with_prefixes)
    for row in ws.iter_rows(min_row=ws.max_row, max_row=ws.max_row, min_col=1, max_col=len(columns_with_prefixes)):
        for cell in row:
            cell.border = border

    for column_name in columns_with_prefixes:
        df[column_name] = df[column_name].astype(str).str.lower()
        ws.append([])

        # Registro de datos duplicados en cada columna con los prefijos
        duplicate_data = df.duplicated(subset=[column_name]).sum()
        ws.append(["Número de filas duplicadas en la columna '{}': {}".format(column_name, duplicate_data)])
        for row in ws.iter_rows(min_row=ws.max_row, max_row=ws.max_row, min_col=1, max_col=1):
            for cell in row:
                cell.font = Font(bold=True)
                cell.border = border

        if duplicate_data > 0:
            ws.append(["Filas duplicadas en la columna '{}':".format(column_name)])
            for row in ws.iter_rows(min_row=ws.max_row, max_row=ws.max_row, min_col=1, max_col=1):
                for cell in row:
                    cell.font = Font(bold=True)
                    cell.border = border

            duplicated_rows = df[df.duplicated(subset=[column_name], keep=False)]

            headers = list(duplicated_rows.columns)
            ws.append(headers)

            # Aplicar negrilla y bordes a los encabezados de las columnas duplicadas
            for row in ws.iter_rows(min_row=ws.max_row, max_row=ws.max_row, min_col=1, max_col=len(headers)):
                for cell in row:
                    cell.font = Font(bold=True)
                    cell.border = border

            for row in duplicated_rows.itertuples(index=False):
                ws.append(list(row))
                for row in ws.iter_rows(min_row=ws.max_row, max_row=ws.max_row, min_col=1, max_col=len(headers)):
                    for cell in row:
                        cell.border = border

    ws.append([])  # Añadir una línea en blanco al final de cada hoja para claridad

    # Autoajustar el ancho de las columnas, exceptuando la primera columna
    for col in ws.columns:
        if col[0].column_letter != 'A':  # Exceptuar la primera columna
            max_length = 0
            column = col[0].column_letter  # Obtener la letra de la columna
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = max_length + 2
            ws.column_dimensions[column].width = adjusted_width

def process_excel_files_in_folder(folder_path, column_prefixes, output_excel_filename):
    excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') or f.endswith('.xls')]

    if not excel_files:
        print("No se encontraron archivos Excel en la carpeta especificada.")
        return

    wb = Workbook()
    wb.remove(wb.active)  # Eliminar la hoja de trabajo por defecto

    for excel_file in excel_files:
        excel_path = os.path.join(folder_path, excel_file)
        sheet_name = os.path.splitext(excel_file)[0][:31]

        try:
            df = pd.read_excel(excel_path)
            print(f"Archivo '{excel_path}' leído exitosamente.")

            ws = wb.create_sheet(title=sheet_name)

            log_null_and_duplicate_data(df, ws, excel_file, column_prefixes)

            print(f"Registro completado para '{excel_file}'. Revisa el archivo '{output_excel_filename}' para los detalles.")
        except FileNotFoundError:
            print(f"Error: El archivo '{excel_path}' no fue encontrado.")
        except Exception as e:
            print(f"Ocurrió un error al procesar el archivo '{excel_path}': {e}")

    wb.save(output_excel_filename)

def main():
    if len(sys.argv) < 4:
        print("Uso: python programa.py carpeta_de_excel archivo_salida prefijo1 [prefijo2 ... prefijoN]")
        sys.exit(1)

    folder_path = sys.argv[1]
    output_excel_filename = sys.argv[2]
    column_prefixes = sys.argv[3:]

    if not os.path.isdir(folder_path):
        print(f"Error: La carpeta '{folder_path}' no existe o no es un directorio.")
        sys.exit(1)

    process_excel_files_in_folder(folder_path, column_prefixes, output_excel_filename)

if __name__ == "__main__":
    main()
