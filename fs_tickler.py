import pandas as pd
from pathlib import Path

def process_excel_files():
    # Setup paths
    input_dir = Path('./assets')
    output_dir = Path('./outputs')
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Process each Excel file
    for excel_path in input_dir.glob('*.xlsx'):
        # Read raw data with all values as strings
        df = pd.read_excel(
            excel_path, 
            header=None, 
            dtype=str, 
            engine='openpyxl'
        ).fillna('')  # Replace empty cells with empty strings
        
        all_customers = []
        current_customer = None
        header = None
        collecting_data = False
        
        # Iterate through each row
        for _, row in df.iterrows():
            # Check for customer name row (always in first column)
            if row[0].startswith('Customer Name: '):
                # Save previous customer data if exists
                if current_customer and current_customer['data']:
                    all_customers.append(current_customer)
                
                # Start new customer
                current_customer = {
                    'name': row[0].split(': ')[1].strip(),
                    'header': None,
                    'data': []
                }
                collecting_data = False
                continue
            
            # Skip rows until we find a customer
            if not current_customer:
                continue
            
            # Check for header row (comes after blank line following customer name)
            if not collecting_data and row[0] == '':
                collecting_data = True  # Next non-empty row will be header
                continue
                
            if collecting_data and current_customer['header'] is None:
                # Capture header row
                current_customer['header'] = row.tolist()
                collecting_data = True  # Next rows will be data
                continue
            
            # Collect data rows until we hit a blank line
            if collecting_data and current_customer['header'] is not None:
                if row[0] == '':  # Blank line marks end of data
                    # Save completed customer
                    all_customers.append(current_customer)
                    current_customer = None
                    collecting_data = False
                    continue
                
                # Add data row with customer name
                data_row = row.tolist()
                data_row.append(current_customer['name'])  # Add customer name as last column
                current_customer['data'].append(data_row)
        
        # Create final DataFrame for this file
        final_df = pd.DataFrame()
        for customer in all_customers:
            # Add header with 'Customer Name' as last column
            customer_header = customer['header'] + ['Customer Name']
            customer_df = pd.DataFrame(customer['data'], columns=customer_header)
            final_df = pd.concat([final_df, customer_df], ignore_index=True)
        
        # Save to output
        output_path = output_dir / excel_path.name
        final_df.to_excel(output_path, index=False, engine='openpyxl')

if __name__ == '__main__':
    process_excel_files()
