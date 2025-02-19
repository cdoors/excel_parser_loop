import pandas as pd
from pathlib import Path

input_dir = Path('./assets')
output_dir = Path('./output')
output_dir.mkdir(parents=True, exist_ok=True)

for file_path in input_dir.glob('*.xlsx'):
    df = pd.read_excel(file_path, header=None, engine='openpyxl')
    all_data = []
    index = 0
    while index < len(df):
        row = df.iloc[index]
        if isinstance(row[0], str) and row[0].startswith('Customer Name: '):
            customer_name = row[0].split(': ')[1].strip()
            index += 2  # Skip Customer Name row and the blank line
            if index >= len(df):
                break
            headers = df.iloc[index].tolist()
            index += 1
            data_rows = []
            while index < len(df):
                current_row = df.iloc[index]
                if current_row.isnull().all():
                    break
                data_rows.append(current_row.tolist())
                index += 1
            if data_rows:
                customer_df = pd.DataFrame(data_rows, columns=headers)
                customer_df['Customer Name'] = customer_name
                all_data.append(customer_df)
            index += 1  # Skip the blank line after data
        else:
            index += 1
    final_df = pd.concat(all_data, ignore_index=True) if all_data else pd.DataFrame()
    output_path = output_dir / file_path.name
    final_df.to_excel(output_path, index=False, engine='openpyxl')
