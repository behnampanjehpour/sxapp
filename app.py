from flask import Flask, request, send_file 
import pandas as pd
import os
from io import BytesIO
from datetime import datetime

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

@app.route('/')
def upload_page():
    return '''
    <!doctype html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Upload File</title>
        <style>
            body {
                font-family: Arial, sans-serif;
                background-color: #0b1623;
                color: white;
                text-align: center;
                padding: 50px;
            }
            h2 {
                color: #ccfc4c;
            }
            form {
                background-color: #1e2a38;
                padding: 20px;
                border-radius: 10px;
                display: inline-block;
                margin-top: 20px;
            }
            input, button {
                padding: 10px;
                margin: 10px;
                border: none;
                border-radius: 5px;
                font-size: 16px;
            }
            button {
                background-color: #ccfc4c;
                color: #0b1623;
                font-weight: bold;
                cursor: pointer;
            }
            button:hover {
                background-color: #b6e640;
            }
            img {
                max-width: 500px;
                margin-bottom: 50px;
            }
        </style>
    </head>
    <body>
        <img src="https://www.solidxperts.com/wp-content/uploads/2024/05/Logo-lime.png" alt="SolidXperts Logo">
        <h2>Upload the SW Leads Excel File(xlsx) for Processing</h2>
        <form action="/upload" method="post" enctype="multipart/form-data">
            <input type="file" name="file" required>
            <button type="submit">Upload and Process</button>
        </form>
    </body>
    </html>
    '''

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "No file part"
    
    file = request.files['file']
    if file.filename == '':
        return "No selected file"
    
    if file:
        df = pd.read_excel(file)

        # 1. Remove "/" and "Region" from column name "Country/Region"
        df.columns = [col.replace("Region", "").replace("/", "").strip() if "Country/Region" in col else col for col in df.columns]
        
        # 2. Replace "PQ" with "QC" in "State" column
        if 'State' in df.columns:
            df['State'] = df['State'].replace('PQ', 'QC')
        
        # 3. Add "Lead Source" column with value "Lead - Solidworks"
        df['Lead Source'] = "Lead - Solidworks"
        
        # 4. Fill missing values and clean up specific columns
        if 'ContactPhone' in df.columns:
            df['ContactPhone'] = df['ContactPhone'].fillna("1111111111")
            df['ContactPhone'] = df['ContactPhone'].astype(str).str.replace(r'[()\s-]', '', regex=True)
        if 'EmailAddress' in df.columns:
            df['EmailAddress'] = df['EmailAddress'].fillna("noemail@gmail.com")
        if 'Country' in df.columns:
            df['Country'] = df['Country'].fillna("Canada")
            df['Country'] = df['Country'].apply(lambda x: "Canada" if x not in ["Canada", "USA"] else x)
        if 'State' in df.columns:
            df['State'] = df['State'].fillna("QC")
        if 'CompanyName' in df.columns:
            df['CompanyName'] = df['CompanyName'].fillna("NA")
        
        # 5. Assign "Cold" if Source is "eDrawings Activation"
        if 'Source' in df.columns and 'Lead Rating' in df.columns:
            df.loc[df['Source'] == 'eDrawings Activation', 'Lead Rating'] = 'Cold'
        if 'Lead Rating' in df.columns:
            df['Lead Rating'] = df['Lead Rating'].replace("5 - very cold", "cold")
            df['Lead Rating'] = df['Lead Rating'].replace("very cold", "cold")
            df['Lead Rating'] = df['Lead Rating'].replace("5 - Very Cold", "cold")

        
        # 6. Assign "Hot" if Source contains RAQ, RAC, RAD, or SWOPT
        keywords = ['RAQ', 'RAC', 'RAD', 'SWOPT']
        if 'Source' in df.columns and 'Lead Rating' in df.columns:
            df.loc[df['Source'].astype(str).str.contains('|'.join(keywords), case=False, na=False), 'Lead Rating'] = 'Hot'
        
        # 7. Assign "Warm" if Lead Rating is still empty
        if 'Lead Rating' in df.columns:
            df['Lead Rating'] = df['Lead Rating'].fillna('Warm')

        # 8. If "Prospect Lead Reassignment" == "Y", add "Reassigned Lead" to Notes
        if 'Prospect Lead Reassignment' in df.columns:
            mask = df['Prospect Lead Reassignment'].astype(str).str.upper().str.strip() == 'Y'
            # Ensure Notes column exists
            if 'Notes' not in df.columns:
                df['Notes'] = ''
            # Normalize Notes on masked rows and append without duplicating
            df.loc[mask, 'Notes'] = df.loc[mask, 'Notes'].fillna('')
            df.loc[mask, 'Notes'] = df.loc[mask, 'Notes'].apply(
                lambda x: 'Reassigned Lead' if x.strip() == '' 
                else (x if 'Reassigned Lead' in x else f"{x}; Reassigned Lead")
            )

        # Generate dynamic filename with today's date
        today_date = datetime.today().strftime('%Y-%m-%d')
        filename = f"lead SW {today_date}.xlsx"
        
        # Save processed file to memory as XLSX
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Processed Data')
        output.seek(0)
        
        return send_file(output, as_attachment=True, download_name=filename, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000)
