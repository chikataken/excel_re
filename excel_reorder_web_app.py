#!/usr/bin/env python3
"""
Flask Web App for Excel Column Processor with CSV Export

This app allows two workflows:
1) Upload an origin Excel to clean/reorder per preference and download a new Excel ending in `..._readable.xlsx`.
2) Upload a processed Excel to map headers to the import-template and download CSV ending in `..._superdispatch.csv`.

Dependencies:
    pip install flask pandas xlsxwriter openpyxl
"""
import os
import io
import datetime
import pandas as pd
from flask import Flask, request, send_file, render_template_string

app = Flask(__name__)

# Path to header template CSV
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
IMPORT_TEMPLATE_CSV = os.path.join(BASE_DIR, 'import-template.csv')

# Constants for origin cleaning
PREFERENCE_COLUMNS = [
    'ShipmentNumber','Vin','OriginState','OriginCity',
    'OriginAddress','OriginZip','OriginContactPhone',
    'DestinationState','DestinationCity','DestinationAddress',
    'DestinationZip','DestinationContactPhone'
]
FIRST_COLUMNS = ['Vin','OriginState','OriginCity','DestinationState','DestinationCity','Price']
OTHER_COLUMNS = [c for c in PREFERENCE_COLUMNS if c not in FIRST_COLUMNS]
SORT_ORDER = ['OriginState','OriginCity','DestinationState','DestinationCity','Vin']
FIRST_COLUMN_WIDTH = 25
OTHER_COLUMN_WIDTH = 17

# HTML form with two file inputs
HTML = '''
<!doctype html>
<title>Excel Column Processor</title>
<h1>Excel Column Processor</h1>
<form method=post enctype=multipart/form-data>
  <p><strong>1) Upload Original Excel to clean (Excel output):</strong><br>
     <input type=file name=origin accept=".xls,.xlsx"></p>
  <p><strong>2) OR Upload Processed Excel for template CSV output:</strong><br>
     <input type=file name=processed_csv accept=".xls,.xlsx"></p>
  <p><input type=submit value=Process></p>
</form>
'''

def clean_origin(df):
    # Add Price column if missing
    if 'Price' not in df.columns:
        df['Price'] = ''
    # Filter and reorder by preference
    cols = [c for c in PREFERENCE_COLUMNS + ['Price'] if c in df.columns]
    df = df[cols]
    final_order = [c for c in FIRST_COLUMNS if c in df.columns] + [c for c in OTHER_COLUMNS if c in df.columns]
    df = df[final_order]
    # Multi-level sort
    sort_cols = [c for c in SORT_ORDER if c in df.columns]
    if sort_cols:
        df = df.sort_values(by=sort_cols)
    return df


def map_to_import_template(df):
    # 1) Read in the template headers, stripping whitespace
    tpl = pd.read_csv(IMPORT_TEMPLATE_CSV, nrows=0)
    desired_headers = [h.strip() for h in tpl.columns.tolist()]

    # 2) Explicit map from your cleaned Excel columns → template header names
    col_map = {
        'ShipmentNumber':           'Order ID',
        'Vin':                      'VIN',
        'OriginAddress':            'Pickup Street',
        'OriginCity':               'Pickup City',
        'OriginState':              'Pickup State',
        'OriginZip':                'Pickup Zip Code',
        'OriginContactPhone':       'Pickup Contact Phone',
        'DestinationAddress':       'Delivery Street',
        'DestinationCity':          'Delivery City',
        'DestinationState':         'Delivery State',
        'DestinationZip':           'Delivery Zip Code',
        'DestinationContactPhone':  'Delivery Contact Phone',
        'Price':                    'Carrier Price per Vehicle',
    }

    # 3) Build new DataFrame in template order
    new_df = pd.DataFrame(index=df.index)
    for header in desired_headers:
        # find which cleaned‐Excel column maps here (if any)
        pref_col = next((k for k,v in col_map.items() if v == header), None)
        if pref_col and pref_col in df.columns:
            # pull in your data
            new_df[header] = df[pref_col]
        else:
            # otherwise blank
            new_df[header] = [''] * len(df)

    return new_df



def to_excel_bytes(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        ws = writer.sheets['Sheet1']
        max_row, max_col = df.shape
        ws.autofilter(0, 0, max_row, max_col-1)
        ws.freeze_panes(1, 0)
        ws.set_column(0, 0, FIRST_COLUMN_WIDTH)
        if max_col > 1:
            ws.set_column(1, max_col-1, OTHER_COLUMN_WIDTH)
    buf.seek(0)
    return buf


def to_csv_bytes(df):
    csv_str = df.to_csv(index=False)
    buf = io.BytesIO(csv_str.encode('utf-8'))
    buf.seek(0)
    return buf

@app.route('/', methods=['GET','POST'])
def upload():
    if request.method == 'POST':
        f_origin = request.files.get('origin')
        f_csv = request.files.get('processed_csv')

        if f_origin and f_origin.filename:
            df = pd.read_excel(f_origin)
            df_out = clean_origin(df)
            out_buf = to_excel_bytes(df_out)
            ext = 'xlsx'
            suffix = 'readable'
        elif f_csv and f_csv.filename:
            df = pd.read_excel(f_csv)
            df_out = map_to_import_template(df)
            out_buf = to_csv_bytes(df_out)
            ext = 'csv'
            suffix = 'superdispatch'
        else:
            return 'No file uploaded', 400

        # Build dynamic filename
        now = datetime.datetime.now()
        date_str = f"{now.strftime('%B')}_{now.day}"
        filename = f"{date_str}_{suffix}.{ext}"

        mimetype = 'text/csv' if ext == 'csv' else 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        return send_file(
            out_buf,
            as_attachment=True,
            download_name=filename,
            mimetype=mimetype
        )
    return render_template_string(HTML)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
