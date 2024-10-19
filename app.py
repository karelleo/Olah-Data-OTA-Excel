from flask import Flask, request, render_template, send_file
import pandas as pd
import os
import openpyxl
from openpyxl.styles import Font, Alignment
import logging
import matplotlib.pyplot as plt
import io
from openpyxl.drawing.image import Image
import numpy as np
import matplotlib
matplotlib.use('Agg')

def add_dataframe_to_worksheet(ws, df, start_row, start_col):
    # Add headers
    headers = ["Type", "Populasi Tams", "Download Completed start Scheduller", "Apply config", "Active transaksi"]
    for col, header in enumerate(headers, start=start_col):
        cell = ws.cell(row=start_row, column=col, value=header)
        cell.font = Font(bold=True)
    
    # Add data
    for r_idx, row in enumerate(df.itertuples(index=False), start=start_row+1):
        for c_idx, value in enumerate(row, start=start_col):
            ws.cell(row=r_idx, column=c_idx, value=value)

app = Flask(__name__)
logging.basicConfig(level=logging.INFO)

def create_pie_chart(data, labels, title):
    width_inch = 4.87
    height_inch = 2.15
    
    plt.figure(figsize=(width_inch, height_inch))
    colors = ['#1f77b4', '#ff7f0e']  # Biru dan oranye
    
    def make_autopct(values):
        def my_autopct(pct):
            total = sum(values)
            val = int(round(pct*total/100.0))
            return f'{pct:.1f}%\n({val:,d})'
        return my_autopct
    
    wedges, texts, autotexts = plt.pie(data, colors=colors, 
                                       autopct=make_autopct(data), 
                                       startangle=90, pctdistance=0.75)
    
    plt.title(title, fontsize=10, pad=2)
    
    plt.legend(wedges, labels,
               title="Types",
               loc="center left",
               bbox_to_anchor=(0.85, 0, 0.5, 1),
               fontsize=8,
               title_fontsize=9)
    
    plt.setp(autotexts, size=8, weight="bold")
    plt.axis('equal')
    
    plt.tight_layout(pad=0.5, w_pad=0.5, h_pad=0.5)
    
    img_buffer = io.BytesIO()
    plt.savefig(img_buffer, format='png', bbox_inches='tight', dpi=300)
    img_buffer.seek(0)
    plt.close()
    
    return img_buffer

def read_excel_file(file_path, skiprows):
    try:
        return pd.read_excel(file_path, skiprows=skiprows)
    except Exception as e:
        logging.error(f"Error reading file {file_path}: {e}")
        return None

def process_a920_data(df, version):
    filtered_a920pro = df[
        (df['Terminal Model'] == 'A920Pro') &
        (df['App Name'] == 'A920PRO_BRIREGULAR') &
        (df['Version'] == version) &
        (df['Status'] == 'Completed')
    ]
    return filtered_a920pro['Serial Number'].nunique()

def process_x990_data(df, version):
    filtered_x990 = df[
        (df['Terminal Model'] == 'X990') &
        (df['App Name'] == 'X990_BRIREGULAR') &
        (df['Version'] == version) &
        (df['Status'] == 'Completed')
    ]
    return filtered_x990['Serial Number'].nunique()

def process_files(file_paths, version_a920pro, version_x990):
    df_a920 = read_excel_file(file_paths['file_terminal_version_a920'][0], 8)
    df_x990 = read_excel_file(file_paths['file_terminal_version_x990'][0], 8)
    df_terminal_download = read_excel_file(file_paths['file_terminal_download'][0], 4)
    df_data_aktif = read_excel_file(file_paths['file_data_aktif'][0], 0)

    if any(df is None for df in [df_a920, df_x990, df_terminal_download, df_data_aktif]):
        return None, "Error: One or more required files are missing or invalid"

    a920_populasi = df_a920['Serial Number'].nunique()
    x990_populasi = df_x990['Serial Number'].nunique()

    # Perhitungan Download Completed start Scheduller
    a920_download_completed = df_terminal_download[
        (df_terminal_download['Terminal Model'] == 'A920Pro') &
        (df_terminal_download['App Name'] == 'A920PRO_BRIREGULAR') &
        (df_terminal_download['Version'] == version_a920pro) &
        (df_terminal_download['Status'] == 'Completed')
    ]['Serial Number'].nunique()

    x990_download_completed = df_terminal_download[
        (df_terminal_download['Terminal Model'] == 'X990') &
        (df_terminal_download['App Name'] == 'X990_BRIREGULAR') &
        (df_terminal_download['Version'] == version_x990) &
        (df_terminal_download['Status'] == 'Completed')
    ]['Serial Number'].nunique()

    # Perhitungan Apply config
    apply_config_a920pro = df_a920[
        (df_a920['APP Name'] == 'A920PRO_BRIREGULAR') &
        (
            (df_a920['Actual APP Version'] == version_a920pro) |
            (df_a920['Actual APP Version'] == '0.0.0.0') |
            (df_a920['Actual APP Version'].isnull())
        )
    ]['Serial Number'].nunique()

    apply_config_x990 = df_x990[
        (df_x990['APP Name'] == 'X990_BRIREGULAR') &
        (
            (df_x990['Actual APP Version'] == version_x990) |
            (df_x990['Actual APP Version'] == '0.0.0.0') |
            (df_x990['Actual APP Version'].isnull())
        )
    ]['Serial Number'].nunique()

    df_data_aktif['FSN'] = df_data_aktif['FSN'].fillna('').astype(str)
    a920_active_transaksi = df_data_aktif[df_data_aktif['FSN'].str.startswith('185')]['FSN'].nunique()
    x990_active_transaksi = df_data_aktif[df_data_aktif['FSN'].str.startswith('V1E')]['FSN'].nunique()

    total_populasi_tams = a920_populasi + x990_populasi
    total_download_completed = a920_download_completed + x990_download_completed
    total_apply_config = apply_config_a920pro + apply_config_x990
    total_active_transaksi = a920_active_transaksi + x990_active_transaksi

    result_df = pd.DataFrame({
        'Type': ['A920pro', 'X990', 'Total'],
        'Populasi Tams': [a920_populasi, x990_populasi, total_populasi_tams],
        'Download Completed start Scheduller': [a920_download_completed, x990_download_completed, total_download_completed],
        'Apply config': [apply_config_a920pro, apply_config_x990, total_apply_config],
        'Active transaksi': [a920_active_transaksi, x990_active_transaksi, total_active_transaksi]
    })

    return result_df

expected_files = {
    'sharing': ['file_data_aktif', 'file_terminal_download', 'file_terminal_version_a920', 'file_terminal_version_x990'],
    'fms': ['file_data_aktif', 'file_terminal_download', 'file_terminal_version_a920', 'file_terminal_version_x990']
}

logging.basicConfig(level=logging.DEBUG)

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        logging.info("Received POST request")
        logging.debug(f"Form data: {request.form}")
        logging.debug(f"Files received: {request.files}")
        
        try:
            file_paths = {'sharing': {}, 'fms': {}}
            
            for data_type, file_list in expected_files.items():
                logging.info(f"Processing {data_type} files")
                for file_key in file_list:
                    files = request.files.getlist(f"{file_key}_{data_type}")
                    logging.info(f"Files for {file_key}_{data_type}: {[f.filename for f in files]}")
                    
                    if files:
                        for file in files:
                            file_path = os.path.join('uploads', file.filename)
                            file.save(file_path)
                            logging.info(f"Saved file: {file_path}")
                            if file_key not in file_paths[data_type]:
                                file_paths[data_type][file_key] = []
                            file_paths[data_type][file_key].append(file_path)
                    else:
                        logging.error(f"Missing required file: {file_key} for {data_type}")
                        return f"Error: Missing required file: {file_key} for {data_type}", 400

            logging.info("All required files received successfully")

            versions = {
                'sharing': {
                    'a920pro': request.form.get('version_a920pro_sharing'),
                    'x990': request.form.get('version_x990_sharing')
                },
                'fms': {
                    'a920pro': request.form.get('version_a920pro_fms'),
                    'x990': request.form.get('version_x990_fms')
                }
            }
            logging.info(f"Versions: {versions}")

            logging.info("Starting file processing")

            wb = openpyxl.Workbook()
            ws = wb.active

            results = {}
            for data_type, title in [('sharing', 'TAMS SHARING 10.2.2.5:7000'), 
                                     ('fms', 'TAMS FMS 10.2.30.2:7000')]:
                results[data_type] = process_files(
                    file_paths[data_type], 
                    versions[data_type]['a920pro'], 
                    versions[data_type]['x990']
                )

                if results[data_type] is None:
                    return f"Error processing data for {data_type}", 400

                if ws.max_row > 1:
                    ws.append([])

                ws.append([title])
                title_cell = ws.cell(row=ws.max_row, column=1)
                title_cell.font = Font(size=16, bold=True)
                title_cell.alignment = Alignment(horizontal='center')
                ws.merge_cells(start_row=ws.max_row, start_column=1, end_row=ws.max_row, end_column=5)

                add_dataframe_to_worksheet(ws, results[data_type], ws.max_row + 1, 1)

                # Create and add pie charts
                chart_row = ws.max_row + 2

                active_transaksi = results[data_type]['Active transaksi'].iloc[:-1].sum()
                download_completed = results[data_type]['Download Completed start Scheduller'].iloc[:-1].sum()
                apply_config = results[data_type]['Apply config'].iloc[:-1].sum()

                # Chart 1: Active vs Download
                chart_data = [active_transaksi, download_completed]
                chart_labels = ['Active Transaksi', 'Download Completed']
                img_buffer = create_pie_chart(chart_data, chart_labels, f'{data_type.upper()} - Active vs Download')
                img = Image(img_buffer)
                img.width = 468
                img.height = 206
                ws.add_image(img, f'A{chart_row}')

                # Chart 2: Active vs Apply Config
                chart_data = [active_transaksi, apply_config]
                chart_labels = ['Active Transaksi', 'Apply Config']
                img_buffer = create_pie_chart(chart_data, chart_labels, f'{data_type.upper()} - Active vs Apply Config')
                img = Image(img_buffer)
                img.width = 468
                img.height = 206
                ws.add_image(img, f'J{chart_row}')

                chart_row += 16

            # Adjust column widths
            for column in ws.columns:
                max_length = 0
                column_letter = openpyxl.utils.get_column_letter(column[0].column)
                for cell in column:
                    try:
                        if cell.coordinate in ws.merged_cells:  # Skip merged cells
                            continue
                        cell_value = str(cell.value) if cell.value is not None else ''
                        if len(cell_value) > max_length:
                            max_length = len(cell_value)
                    except AttributeError:
                        pass
                adjusted_width = (max_length + 2)
                ws.column_dimensions[column_letter].width = adjusted_width

            # Add "TAMS FMS dan Sharing" section
            ws.append([])
            ws.append(["TAMS FMS dan SHARING "])
            title_cell = ws.cell(row=ws.max_row, column=1)
            title_cell.font = Font(size=16, bold=True)
            title_cell.alignment = Alignment(horizontal='center')
            ws.merge_cells(start_row=ws.max_row, start_column=1, end_row=ws.max_row, end_column=5)

            # Combine data from TAMS Sharing and TAMS FMS
            combined_data = results['sharing'].set_index('Type') + results['fms'].set_index('Type')
            combined_data = combined_data.reset_index()

            add_dataframe_to_worksheet(ws, combined_data, ws.max_row + 2, 1)  # +2 untuk memberikan satu baris kosong

            # Adjust column widths for the combined section
            for col in range(1, ws.max_column + 1):
                max_length = 0
                column_letter = openpyxl.utils.get_column_letter(col)
                for row in range(1, ws.max_row + 1):
                    cell = ws.cell(row=row, column=col)
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = (max_length + 2)
                ws.column_dimensions[column_letter].width = adjusted_width

            # Save Excel file
            output_filename = 'output.xlsx'
            wb.save(output_filename)

            # Delete uploaded files
            for data_type in file_paths:
                for file_key in file_paths[data_type]:
                    for file_path in file_paths[data_type][file_key]:
                        try:
                            os.remove(file_path)
                        except Exception as e:
                            print(f"Error deleting file: {e}")

            return send_file('output.xlsx', as_attachment=True)

        except Exception as e:
            logging.exception(f"Unexpected error occurred: {e}")
            return f"An unexpected error occurred: {str(e)}", 500

    return render_template('index.html')

if __name__ == '__main__':
    if not os.path.exists('uploads'):
        os.makedirs('uploads')
    app.run(host='0.0.0.0', port=5112, debug=True)
                    
        