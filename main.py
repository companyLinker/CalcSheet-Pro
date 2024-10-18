import os
import zipfile
from io import BytesIO
import pandas as pd
import datetime
from flask import Flask, request, send_file, render_template


app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    # Read uploaded mapping and excel files
    mappings_df = pd.read_csv(request.files['mapping_file'],encoding='ISO-8859-1')
    bank_mappings_df = pd.read_csv(request.files['bank_mapping_file'],encoding='ISO-8859-1')
    print(mappings_df.head)
    print(bank_mappings_df.head)
    excel_file = pd.read_csv(request.files['excel_file'])

    # Create a directory to store the output files
    output_dir = 'output'
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Generate output files
    excel_to_iif(excel_file, output_dir, mappings_df, bank_mappings_df)

    # Create a zip file
    zip_filename = 'output_files.zip'
    zip_path = os.path.join(output_dir, zip_filename)
    with zipfile.ZipFile(zip_path, 'w') as zip_ref:
        for root, dirs, files in os.walk(output_dir):
            for file in files:
                if file not in [zip_filename, 'output.zip']:  # exclude the zip file itself and output.zip
                    file_path = os.path.join(root, file)
                    zip_ref.write(file_path, os.path.relpath(file_path, output_dir))

    # Send the zip file as a response
    return send_file(zip_path, as_attachment=True, download_name=zip_filename)

def excel_to_iif(excel_file, output_dir, mappings_df, bank_mappings_df):
    df = excel_file
    column_name = 'Store Name'
    unique_store_names = df[column_name].unique()
    print(unique_store_names)

    for desired_store_name in unique_store_names:
        filtered_df = df[(df[column_name] == desired_store_name)]
        iif_filename = f"{desired_store_name}.iif"
        iif_path = os.path.join(output_dir, iif_filename)
        iif_content = '!TRNS\tTRNSID\tTRNSTYPE\tDATE\tACCNT\tCLASS\tAMOUNT\tDOCNUM\tMEMO\tCLEAR\n!SPL\tSPLID\tTRNSTYPE\tDATE\tACCNT\tCLASS\tAMOUNT\tDOCNUM\tMEMO\tCLEAR\n!ENDTRNS\n'
        
        for _, row in filtered_df.iterrows():
            for new_col_name, cell_value in row.items():
                if new_col_name != 'Store Name' and new_col_name != 'Customer Cou' and new_col_name == 'Date':
                    if cell_value != 0 and cell_value != '&nbsp;':
                        custom_fix_memo = get_custom_fix_memo(new_col_name, row)
                        total1_value = total1(row)
                        store_name = row['Store Name']
                        new_col_name = replace_col_name(new_col_name, store_name, bank_mappings_df)
                        cell_value1 = datetime.datetime.strptime(str(row['Date']), '%m/%d/%Y').date()
                        value1 = float(row['#1          '])
                        value2 = float(row['#2          '])
                        value3 = float(row['#3          '])
                        value4 = float(row['#4          '])
                        value5 = float(row['#5          '])
                        Total_val = round(value1 + value2 + value3 + value4 + value5, 2)
                        iif_content += f"TRNS\t\tDEPOSIT\t{cell_value1}\t{new_col_name}\t\t{Total_val}\t\t{total1_value}\tN\n"
                        break

            for col_name, cell_value in row.items():
                if col_name != 'Store Name' and col_name != 'Date' and col_name != 'Customer Cou' and col_name not in ['#1          ', '#2          ', '#3          ', '#4          ', '#5          ']:
                    if cell_value != 0 or col_name == 'Cash Shortage':
                        custom_fix_memo = get_custom_fix_memo(col_name, row)
                        col_name = replace_col_name1(col_name, mappings_df)
                        cell_value1 = datetime.datetime.strptime(str(row['Date']), '%m/%d/%Y').date()
                        iif_content += f'SPL\t\tDEPOSIT\t{cell_value1}\t{col_name}\t\t{cell_value}\t\t{custom_fix_memo}\tN\n'   
            iif_content += 'ENDTRNS\n'
        
        with open(iif_path, 'w') as f:
            f.write(iif_content)

def total1(row):
    val1 = round(float(row['#1          ']), 2)
    val2 = round(float(row['#2          ']), 2)
    val3 = round(float(row['#3          ']), 2)
    val4 = round(float(row['#4          ']), 2)
    val5 = round(float(row['#5          ']), 2)
    
    if val2 == 0 and val3 == 0 and val4 == 0 and val5 == 0:
        return str(val1)
    elif val3 == 0 and val4 == 0 and val5 == 0:
        return str(val1) + '+' + str(val2)
    elif val4 == 0 and val5 == 0:
        return str(val1) + '+' + str(val2) + '+' + str(val3)
    elif val5 == 0:
        return str(val1) + '+' + str(val2) + '+' + str(val3) + '+' + str(val4)
    else:
        return str(val1) + '+' + str(val2) + '+' + str(val3) + '+' + str(val4) + '+' + str(val5)

def get_custom_fix_memo(col_name, row):
    if col_name == 'MDA Donate     ':
        return 'MDA Donate'
    elif col_name == 'Taxable Sales_I':
        return 'FOOD SALES:Food Sale - Store:Taxable'
    elif col_name == 'Non-Taxable Sal_I':
        return 'FOOD SALES:Food Sale - Store:Non Taxable'
    elif col_name == 'Taxable Sales_D':
        return ''
    elif col_name == 'Taxable Sales_G':
        return ''
    elif col_name == 'Taxable Sales_U':
        return ''
    elif col_name == 'Non-Taxable Sal_D':
        return ''
    elif col_name == 'Non-Taxable Sal_G':
        return ''
    elif col_name == 'Non-Taxable Sal_U':
        return ''
    elif col_name == 'Surcharge DLV':
        return 'Surcharge DLV'
    elif col_name == 'Smart Cart F':
        return 'Smart Cart F'
    elif col_name == 'Unknown Sales':
        return 'Unknown Sales'
    elif col_name == 'Cents Sale':
        return 'Service Fee/Discounts'
    elif col_name == 'Instore Mobile':
        return 'Instore Mobile GC'
    elif col_name == 'EBT Total   ':
        return 'EBT Card ask for clarification'
    elif col_name == 'Taxable Sales  ':
        return row['Customer Cou']
    elif 'Taxable Sales  ' not in row and col_name =='Non-Taxable Sal':
        return row['Customer Cou']
    elif col_name == 'Cash Shortage':
        return total1(row) 
    else:
        return ''      

def replace_col_name1(col_name, mappings_df):
    if 'original_name' in mappings_df.columns and 'new_name' in mappings_df.columns:
        column_mappings = dict(zip(mappings_df['original_name'], mappings_df['new_name']))
        return column_mappings.get(col_name.strip(), col_name)  
    else:
        return col_name  # or some other default value

def replace_col_name(new_col_name, store_name, bank_mappings_df):
    # Read the mappings from a CSV file
    # Create a dataframe with the input values
    input_df = pd.DataFrame({'new_col_name': [new_col_name], 'store_name': [store_name]})

    # Merge the input dataframe with the mappings dataframe
    merged_df = pd.merge(input_df, bank_mappings_df, on=['new_col_name', 'store_name'], how='left')

    # Return the mapped column name if it exists, otherwise return an empty string
    return merged_df['mapped_col_name'].iloc[0] if not merged_df['mapped_col_name'].isnull().iloc[0] else ''


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0')
