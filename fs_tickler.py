import pandas as pd
from pathlib import Path

input_dir = Path('./assets')
output_dir = Path('./output')
output_dir.mkdir(parents=True, exist_ok=True)

for file_path in input_dir.glob('*.xlsx'):
    # Read all data as strings and replace NaN with empty strings
    df = pd.read_excel(
        file_path, 
        header=None, 
        dtype=str, 
        engine='openpyxl'
    ).fillna('')  # Replace NaN with empty strings

    all_data = []
    index = 0
    
    while index < len(df):
        row = df.iloc[index]
        customer_name = None
        
        # Check every cell in the row for "Customer Name: "
        for cell in row:
            if isinstance(cell, str) and cell.startswith('Customer Name: '):
                customer_name = cell.split(': ')[1].strip()
                break  # Exit loop once found
        
        if customer_name:
            index += 2  # Skip Customer Name row and blank line
            if index >= len(df):
                break
                
            # Get headers (ensure they're strings)
            headers = df.iloc[index].astype(str).tolist()
            index += 1
            
            data_rows = []
            while index < len(df):
                current_row = df.iloc[index]
                # Stop at first completely empty row
                if (current_row == '').all():
                    break
                # Convert all values to strings explicitly
                data_rows.append(current_row.astype(str).tolist())
                index += 1
                
            if data_rows:
                customer_df = pd.DataFrame(data_rows, columns=headers)
                customer_df['Customer Name'] = customer_name
                all_data.append(customer_df)
                
            index += 1  # Skip blank line after data
        else:
            index += 1  # Move to next row if no customer name found
            
    final_df = pd.concat(all_data, ignore_index=True) if all_data else pd.DataFrame()
    output_path = output_dir / file_path.name
    final_df.to_excel(output_path, index=False, engine='openpyxl')
