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
            df = None
            for header_pos in [1, 0, 2]:
                try:
                    temp_df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_pos)
                    if not temp_df.empty:
                        temp_df.columns = temp_df.columns.astype(str)
                        col_sample = [str(col).strip() for col in temp_df.columns[:5]]
                        if any(col in ['Data1', 'Tree', 'City', 'Quant', 'Street', 'Gush', 'כמות', 'מין העץ'] for col in col_sample):
                            df = temp_df
                            print(f"Found headers at row {header_pos} for sheet '{sheet_name}'")
                            break
                except:
                    continue
            
            if df is None:
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=1)
                df.columns = df.columns.astype(str)
            
            if not df.empty:
                new_cols = []
                col_count = {}
                
                for i, col in enumerate(df.columns):
                    col_str = str(col).strip()
                    if col_str == 'nan' or col_str.startswith('Unnamed:') or col_str == '':
                        col_str = f'Column_{i+1}'
                    
                    if col_str in col_count:
                        col_count[col_str] += 1
                        col_str = f'{col_str}_{col_count[col_str]}'
                    else:
                        col_count[col_str] = 0
                    
                    new_cols.append(col_str)
                
                df.columns = new_cols
                all_columns.update(df.columns.tolist())
                df['source_sheet'] = sheet_name
                all_data.append(df)
                print(f"Processing sheet '{sheet_name}': {len(df)} rows")
                print(f"  Columns: {df.columns.tolist()[:8]}...")
                
        except Exception as e:
            print(f"Error in sheet '{sheet_name}': {e}")
    
    final_columns = ['source_sheet'] + sorted([col for col in all_columns if col != 'source_sheet'])
    
    col_count = {}
    unique_final_columns = []
    for col in final_columns:
        if col in col_count:
            col_count[col] += 1
            unique_final_columns.append(f'{col}_duplicate_{col_count[col]}')
        else:
            col_count[col] = 0
            unique_final_columns.append(col)
    
    print(f"Total unique columns: {len(unique_final_columns)}")
    
    merged_df = pd.DataFrame()
    for df in all_data:
        temp_df = df.copy()
        for col in unique_final_columns:
            if col not in temp_df.columns:
                temp_df[col] = ''
        
        temp_df = temp_df.reindex(columns=unique_final_columns, fill_value='')
        merged_df = pd.concat([merged_df, temp_df], ignore_index=True)
    
    return merged_df, unique_final_columns

def improve_column_names(df, columns):
    improved_names = []
    
    for col in columns:
        if col == 'source_sheet':
            improved_names.append('Source_Sheet')
            continue
        
        sample_values = df[col].dropna().head(50).astype(str).tolist()
        
        if not sample_values:
            improved_names.append(col)
            continue
        
        sample_text = ' '.join(sample_values).lower()
        
        if col.lower() in ['data1', 'data1name']:
            improved_names.append('License_Number')
        elif col.lower() in ['tree']:
            improved_names.append('Tree_Type')
        elif col.lower() in ['quant']:
            improved_names.append('Tree_Quantity')
        elif col.lower() in ['city', 'cityname']:
            improved_names.append('City')
        elif col.lower() in ['street']:
            improved_names.append('Street')
        elif col.lower() in ['homenumber']:
            improved_names.append('House_Number')
        elif col.lower() in ['gush']:
            improved_names.append('Land_Block')
        elif col.lower() in ['helka']:
            improved_names.append('Land_Parcel')
        elif col.lower() in ['fromdate']:
            improved_names.append('Start_Date')
        elif col.lower() in ['todate']:
            improved_names.append('End_Date')
        elif col.lower() in ['price']:
            improved_names.append('Price')
        elif col.lower() in ['siba']:
            improved_names.append('Reason_Code')
        elif col.lower() in ['sibatext']:
            improved_names.append('Reason_Description')
        elif col.lower() in ['peula']:
            improved_names.append('Action_Type')
        elif col.lower() in ['rname']:
            improved_names.append('Applicant_Name')
        elif col.lower() in ['ezor']:
            improved_names.append('Area')
        elif 'date' in col.lower():
            improved_names.append('Date')
        elif 'name' in col.lower():
            improved_names.append('Name')
        elif col.startswith('Column_'):
            improved_names.append(f'Data_{col.split("_")[1]}')
        else:
            improved_names.append(col)
    
    print(f"\nColumn name improvements:")
    for old, new in zip(columns, improved_names):
        if old != new:
            print(f"  {old} -> {new}")
    
    return improved_names

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
    
    print(f"\nImproving column names based on content...")
    improved_columns = improve_column_names(merged_data, columns)
    merged_data.columns = improved_columns
    
    save_formatted_excel(merged_data, improved_columns, output_filename)
    
    print(f"\nFile created successfully with improved column names!")
    
    print("Downloading the merged file:")
    files.download(output_filename)
    
    print(f"\nSample of merged data (first 5 rows):")
    print(merged_data.head())
    
except Exception as e:
    print(f"Error: {e}")

print("\nDone!")
