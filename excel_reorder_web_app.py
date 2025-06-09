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
import traceback
import pandas as pd
from flask import Flask, request, send_file, render_template, url_for

# Initialize Flask with static and template folders
app = Flask(__name__, static_folder='static', template_folder='templates')
app.config['PROPAGATE_EXCEPTIONS'] = True

@app.errorhandler(Exception)
def handle_all_exceptions(e):
    trace = traceback.format_exc()
    app.logger.error("Uncaught exception:\n%s", trace)
    return f"<pre>{trace}</pre>", 500

# Path to header template CSV
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
IMPORT_TEMPLATE_CSV = os.path.join(BASE_DIR, 'import-template.csv')

# Constants for origin cleaning
PREFERENCE_COLUMNS = [
    'ShipmentNumber', 'Vin', 'OriginState', 'OriginCity',
    'OriginAddress', 'OriginZip', 'OriginContactPhone',
    'DestinationState', 'DestinationCity', 'DestinationAddress',
    'DestinationZip', 'DestinationContactPhone'
]
FIRST_COLUMNS = ['Vin', 'OriginState', 'OriginCity', 'DestinationState', 'DestinationCity', 'Price']
OTHER_COLUMNS = [c for c in PREFERENCE_COLUMNS if c not in FIRST_COLUMNS]
SORT_ORDER = ['OriginState', 'OriginCity', 'DestinationState', 'DestinationCity', 'Vin']
FIRST_COLUMN_WIDTH = 25
OTHER_COLUMN_WIDTH = 17
# Extra columns to include at end
EXTRA_COLUMNS = ['ExpectedPickupDate', 'NeedByDate']

# Mapping from CSV headers to readable-file columns
CSV_TO_SOURCE_MAP = {
    'VIN': 'Vin',
    'Order ID': 'ShipmentNumber',
    'Pickup State': 'OriginState',
    'Pickup City': 'OriginCity',
    'Pickup Street': 'OriginAddress',
    'Pickup Zip Code': 'OriginZip',
    'Pickup Contact Phone': 'OriginContactPhone',
    'Delivery State': 'DestinationState',
    'Delivery City': 'DestinationCity',
    'Delivery Street': 'DestinationAddress',
    'Delivery Zip Code': 'DestinationZip',
    'Delivery Contact Phone': 'DestinationContactPhone',
    'Carrier Price per Vehicle': 'Price',
    'Carrier Pickup Scheduled At': 'ExpectedPickupDate',
    'Carrier Delivery Scheduled At': 'NeedByDate'
}

# Predefined static values for certain CSV headers
PREDEFINED_VALUES = {
    'Carrier Payment Method': 'check',
    'Carrier Payment Terms': '10_days',
    'Pickup Date Type': 'estimated',
    'Delivery Date Type': 'estimated'
}


def clean_origin(df):
    if 'Price' not in df.columns:
        df['Price'] = ''
    cols = [c for c in PREFERENCE_COLUMNS + ['Price'] + EXTRA_COLUMNS if c in df.columns]
    df = df[cols]
    extras = [c for c in EXTRA_COLUMNS if c in df.columns]
    final_order = [c for c in FIRST_COLUMNS if c in df.columns] + \
                  [c for c in OTHER_COLUMNS if c in df.columns] + extras
    sort_cols = [c for c in SORT_ORDER if c in df.columns]
    if sort_cols:
        df = df.sort_values(by=sort_cols)
    df = df[final_order]
    return df


def map_to_import_template(df):
    # Load desired headers from template
    tpl = pd.read_csv(IMPORT_TEMPLATE_CSV, nrows=0)
    desired_headers = [h.strip() for h in tpl.columns.tolist()]
    df_out = pd.DataFrame(index=df.index)
    # Map or fill for each header
    for header in desired_headers:
        if header in PREDEFINED_VALUES:
            df_out[header] = [PREDEFINED_VALUES[header]] * len(df)
        elif header in CSV_TO_SOURCE_MAP:
            src_col = CSV_TO_SOURCE_MAP[header]
            if src_col in df.columns:
                if header in ('Pickup State', 'Delivery State'):
                    series = df[src_col].astype(str).str.upper().str.slice(0, 2)
                elif header in ('Carrier Pickup Scheduled At', 'Carrier Delivery Scheduled At'):
                    dt = pd.to_datetime(df[src_col], errors='coerce')
                    series = dt.dt.strftime('%m/%d/%Y').fillna('')
                else:
                    series = df[src_col].astype(str)
                df_out[header] = series
            else:
                df_out[header] = [''] * len(df)
        else:
            df_out[header] = [''] * len(df)
    return df_out


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

        now = datetime.datetime.now()
        date_str = f"{now.strftime('%B')}_{now.day}"
        filename = f"{date_str}_{suffix}.{ext}"

        mimetype = 'text/csv' if ext == 'csv' else \
                   'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        return send_file(
            out_buf,
            as_attachment=True,
            download_name=filename,
            mimetype=mimetype
        )
    # Render the styled template on GET
    return render_template('index.html')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
