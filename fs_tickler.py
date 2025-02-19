import pandas as pd
from pathlib import Path
import numpy as np

# Configure paths
input_folder = Path('./assets')
output_folder = Path('./output')
output_folder.mkdir(parents=True, exist_ok=True)

# Process each Excel file
for excel_file in input_folder.glob('*.xlsx'):
    all_sheets = pd.read_excel(excel_file, sheet_name=None, header=None)
    for sheet_name, sheet_df in all_sheets.items():
        cleaned_data = []
        current_customer = None
        headers = []
        collecting_data = False
        skip_next = False

        for row in sheet_df.itertuples(index=False):
            row = list(row)
            
            # Skip empty rows between customer blocks
            if all(pd.isna(cell) for cell in row):
                collecting_data = False
                headers = []
                continue
                
            # Detect customer name row
            if str(row[0]).startswith('Customer Name:'):
                current_customer = str(row[0]).split(': ')[-1].strip()
                skip_next = True  # Next row will be headers
                continue
                
            # Capture headers after customer name
            if skip_next:
                headers = [str(cell).strip() for cell in row]
                skip_next = False
                collecting_data = True
                continue
                
            # Collect data rows until next blank line
            if collecting_data and not all(pd.isna(cell) for cell in row):
                data_row = {header: cell for header, cell in zip(headers, row)}
                data_row['Customer Name'] = current_customer
                cleaned_data.append(data_row)

        # Create final DataFrame for sheet
        if cleaned_data:
            final_df = pd.DataFrame(cleaned_data)
            final_df = final_df[['Customer Name'] + headers]  # Reorder columns
            
            # Save to output with original structure
            output_path = output_folder / excel_file.name
            with pd.ExcelWriter(output_path, engine='openpyxl', mode='a' if output_path.exists() else 'w') as writer:
                final_df.to_excel(writer, sheet_name=sheet_name, index=False)