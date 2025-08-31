!pip install pandas openpyxl

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from google.colab import files
import io

def merge_excel_sheets(file_path):
    all_data = []
    all_columns = set()
    
    xl_file = pd.ExcelFile(file_path)
    print(f"Found {len(xl_file.sheet_names)} sheets: {xl_file.sheet_names}")
    
    for sheet_name in xl_file.sheet_names:
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            
            if not df.empty:
                df.columns = df.columns.astype(str)
                all_columns.update(df.columns.tolist())
                df['source_sheet'] = sheet_name
                all_data.append(df)
                print(f"Processing sheet '{sheet_name}': {len(df)} rows")
        except Exception as e:
            print(f"Error in sheet '{sheet_name}': {e}")
    
    final_columns = ['source_sheet'] + sorted(list(all_columns))
    print(f"Total columns: {len(final_columns)}")
    
    merged_df = pd.DataFrame()
    for df in all_data:
        for col in final_columns:
            if col not in df.columns:
                df[col] = ''
        df = df[final_columns]
        merged_df = pd.concat([merged_df, df], ignore_index=True)
    
    return merged_df, final_columns

def save_formatted_excel(df, columns, output_path):
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='merged_data', index=False)
        
        workbook = writer.book
        worksheet = writer.sheets['merged_data']
        
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_alignment = Alignment(horizontal="center")
        
        for col_num, column in enumerate(columns, 1):
            cell = worksheet.cell(row=1, column=col_num)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
            worksheet.column_dimensions[cell.column_letter].width = 15

print("Upload the file:")
uploaded = files.upload()

file_name = list(uploaded.keys())[0]
print(f"\nProcessing file: {file_name}")

try:
    merged_data, columns = merge_excel_sheets(file_name)
    
    print(f"\nSummary:")
    print(f"Total rows: {len(merged_data)}")
    print(f"Total columns: {len(columns)}")
    
    source_counts = merged_data['source_sheet'].value_counts()
    print(f"\nBreakdown by sheets:")
    for source, count in source_counts.items():
        print(f"  {source}: {count} rows")
    
    output_filename = 'merged_report_2024.xlsx'
    save_formatted_excel(merged_data, columns, output_filename)
    
    print(f"\nFile created successfully!")
    
    print("Downloading the merged file:")
    files.download(output_filename)
    
    print(f"\nSample of merged data (first 5 rows):")
    print(merged_data.head())
    
except Exception as e:
    print(f"Error: {e}")

print("\nDone!")