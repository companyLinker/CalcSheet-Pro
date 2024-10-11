from flask import Flask, request, send_file, render_template, flash
import os
import pandas as pd
import io
import zipfile
import datetime
import logging

app = Flask(__name__)
app.secret_key = 'your_secret_key_here'

# Set up logging
logging.basicConfig(level=logging.DEBUG)

def excel_to_iif(excel_file, coa_mapping_file, output_dir):
    app.logger.info(f"Starting excel_to_iif function with files: {excel_file}, {coa_mapping_file}")
    try:
        # Read the Excel file
        df = pd.read_csv(excel_file, encoding='ISO-8859-1')
        app.logger.info(f"Excel file read successfully. Shape: {df.shape}")
        
        # Read the COA mapping file
        mappings_df = pd.read_csv(coa_mapping_file)
        column_mappings = dict(zip(mappings_df['original_name'], mappings_df['new_name']))
        app.logger.info(f"COA mapping file read successfully. Mappings: {column_mappings}")
        
        # Group by Store Name
        grouped = df.groupby('Store Name')
        if grouped.ngroups == 0:
            app.logger.error("No groups found. Check if 'Store Name' column exists and is correctly filled.")
            return
        
        for store_name, store_data in grouped:
            app.logger.info(f"Processing store: {store_name}")
            iif_filename = f"{store_name}.iif"
            iif_path = os.path.join(output_dir, iif_filename)
            
            with open(iif_path, 'w') as f:
                # Write IIF header
                f.write('!TRNS\tTRNSID\tTRNSTYPE\tDATE\tACCNT\tNAME\tCLASS\tAMOUNT\tDOCNUM\tMEMO\tCLEAR\tTOPRINT\n')
                f.write('!SPL\tSPLID\tTRNSTYPE\tDATE\tACCNT\tNAME\tCLASS\tAMOUNT\tDOCNUM\tMEMO\tCLEAR\tTOPRINT\n')
                f.write('!ENDTRNS\n')
                
                for _, row in store_data.iterrows():
                    date = datetime.datetime.strptime(str(row['Date']), '%m/%d/%Y').strftime('%m/%d/%Y')
                    total = sum(float(row[col]) for col in ['#1', '#2', '#3', '#4', '#5'] if col in row)
                    
                    # Write TRNS line
                    f.write(f"TRNS\t\tDEPOSIT\t{date}\t{store_name} Deposits\t\t\t{total:.2f}\t\tDeposit\tN\tN\n")
                    
                    # Write SPL lines
                    for col in row.index:
                        if col in column_mappings and col not in ['Store Name', 'Date', 'Customer Cou', '#1', '#2', '#3', '#4', '#5']:
                            amount = float(row[col])
                            if amount != 0:
                                mapped_col = column_mappings[col]
                                f.write(f"SPL\t\tDEPOSIT\t{date}\t{mapped_col}\t\t\t{amount:.2f}\t\t{col}\tN\tN\n")
                
                # Write ENDTRNS
                f.write('ENDTRNS\n')
            
            app.logger.info(f"Finished processing store: {store_name}")
        
        app.logger.info(f"Finished excel_to_iif function")
    except Exception as e:
        app.logger.error(f"An error occurred in excel_to_iif function: {str(e)}")

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Check if the post request has the file parts
        if 'excel_file' not in request.files or 'coa_mapping' not in request.files:
            flash('One or more required files are missing', 'error')
            return render_template('index.html')
        
        excel_file = request.files['excel_file']
        coa_mapping = request.files['coa_mapping']
        
        if excel_file.filename == '' or coa_mapping.filename == '':
            flash('No selected file', 'error')
            return render_template('index.html')
        
        if excel_file and coa_mapping:
            try:
                # Save uploaded files temporarily
                excel_file.save('temp_excel_file.csv')
                coa_mapping.save('temp_coa_mapping.csv')

                # Process files
                output_dir = 'temp_output'
                if not os.path.exists(output_dir):
                    os.mkdir(output_dir)

                excel_to_iif('temp_excel_file.csv', 'temp_coa_mapping.csv', output_dir)

                # Create a zip file of all output files
                memory_file = io.BytesIO()
                with zipfile.ZipFile(memory_file, 'w') as zf:
                    for root, dirs, files in os.walk(output_dir):
                        for file in files:
                            zf.write(os.path.join(root, file), file)
                memory_file.seek(0)

                # Clean up temporary files
                os.remove('temp_excel_file.csv')
                os.remove('temp_coa_mapping.csv')
                for file in os.listdir(output_dir):
                    os.remove(os.path.join(output_dir, file))
                os.rmdir(output_dir)

                return send_file(memory_file, download_name='output.zip', as_attachment=True)
            except Exception as e:
                flash(f'An error occurred: {str(e)}', 'error')
                return render_template('index.html')

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)