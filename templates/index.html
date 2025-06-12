#!/usr/bin/env python3
"""
Flask Web App for Excel → Readable CSV Export

This app allows one workflow:
1) Upload an origin Excel to rename/reorder per predefined mappings,
   add new "model" and "Carrier Price per Vehicle" columns, and download a CSV
   ending in ..._superdispatch.csv.

Dependencies:
    pip install flask pandas openpyxl
"""
import io
import datetime
import pandas as pd
from flask import Flask, request, send_file, render_template

app = Flask(__name__, static_folder='static', template_folder='templates')
app.config['PROPAGATE_EXCEPTIONS'] = True

# === Static lookups ===
# USPS state name → two-letter code
STATE_ABBR = {
    'ALABAMA':'AL','ALASKA':'AK','ARIZONA':'AZ','ARKANSAS':'AR','CALIFORNIA':'CA',
    'COLORADO':'CO','CONNECTICUT':'CT','DELAWARE':'DE','FLORIDA':'FL','GEORGIA':'GA',
    'HAWAII':'HI','IDAHO':'ID','ILLINOIS':'IL','INDIANA':'IN','IOWA':'IA',
    'KANSAS':'KS','KENTUCKY':'KY','LOUISIANA':'LA','MAINE':'ME','MARYLAND':'MD',
    'MASSACHUSETTS':'MA','MICHIGAN':'MI','MINNESOTA':'MN','MISSISSIPPI':'MS','MISSOURI':'MO',
    'MONTANA':'MT','NEBRASKA':'NE','NEVADA':'NV','NEW HAMPSHIRE':'NH','NEW JERSEY':'NJ',
    'NEW MEXICO':'NM','NEW YORK':'NY','NORTH CAROLINA':'NC','NORTH DAKOTA':'ND','OHIO':'OH',
    'OKLAHOMA':'OK','OREGON':'OR','PENNSYLVANIA':'PA','RHODE ISLAND':'RI','SOUTH CAROLINA':'SC',
    'SOUTH DAKOTA':'SD','TENNESSEE':'TN','TEXAS':'TX','UTAH':'UT','VERMONT':'VT',
    'VIRGINIA':'VA','WASHINGTON':'WA','WEST VIRGINIA':'WV','WISCONSIN':'WI','WYOMING':'WY'
}

# VIN fourth-character → model name
MODEL_MAP = {
    'C': 'cyber',
    '3': '3\'',
    'R': 'ROADSTER',
    'T': 'SEMI',
    'S': 'S',
    'X': 'X',
    'Y': 'Y'
}

# === Constant columns with predetermined values ===
CONST_COLS = {
    'Carrier Payment Method':   'check',
    'Carrier Payment Terms':    '10_days',
    'Pickup Date Type':         'estimated',
    'Delivery Date Type':       'estimated'
}

# === Mapping & Order: output header → source column in input Excel ===
CSV_TO_SOURCE_MAP = {
    '     ':                     None,
    'model':                     None,  # decoded from VIN
    'Carrier Price per Vehicle': None,
    'Pickup State':              'OriginState',
    'Pickup City':               'OriginCity',
    'Delivery State':            'DestinationState',
    'Delivery City':             'DestinationContact',
    'Pickup Zip Code':           'OriginZip',
    'VIN':                       'Vin',
    'Order ID':                  'ShipmentNumber',
    'Pickup Street':             'OriginAddress',
    'Pickup Contact Phone':      'OriginContactPhone',
    'Delivery Street':           'DestinationAddress',
    'Delivery City':             'DestinationCity',
    'Delivery Zip Code':         'DestinationZip',
    'Delivery Contact Phone':    'DestinationContactPhone',
    # — your new constant columns —
    'Carrier Payment Method':    None,
    'Carrier Payment Terms':     None,
    'Pickup Date Type':          None,
    'Delivery Date Type':        None
}

@app.route('/', methods=['GET', 'POST'])
def upload():
    if request.method == 'POST':
        file = request.files.get('file')
        if not file or not file.filename:
            return 'No file uploaded', 400

        # read the raw Excel
        df_raw = pd.read_excel(file)

        # build output DataFrame in desired order
        df_out = pd.DataFrame(index=df_raw.index)
        for out_col, src_col in CSV_TO_SOURCE_MAP.items():

            # 1) decoded from VIN
            if out_col == 'model':
                vin = df_raw.get('Vin', pd.Series([''] * len(df_raw))) \
                           .astype(str).str.upper()
                key = vin.str.slice(3, 4)
                series = key.map(MODEL_MAP).fillna('')

            # 2) constant‐value columns
            elif out_col in CONST_COLS:
                series = pd.Series([CONST_COLS[out_col]] * len(df_raw))

            # 3) mapped from an existing source column
            elif src_col and src_col in df_raw.columns:
                # Convert to string, drop any trailing ".0" from floats
                series = (
                    df_raw[src_col]
                    .fillna('')
                    .astype(str)
                    .str.replace(r'\.0$', '', regex=True)
                )
                # Zero-pad ZIP codes back to 5 digits
                if out_col in ('Pickup Zip Code', 'Delivery Zip Code'):
                    series = series.str.zfill(5)
                # State abbreviations logic
                if out_col in ('Pickup State', 'Delivery State'):
                    temp = series.str.strip().str.upper()
                    # if already abbreviated, keep; else map full state names
                    series = temp.where(temp.str.len() == 2, temp.map(STATE_ABBR)).fillna('')

            # 4) default: empty column
            else:
                series = pd.Series([''] * len(df_raw))

            df_out[out_col] = series

        # export to CSV bytes
        buf = io.BytesIO(df_out.to_csv(index=False).encode('utf-8'))
        buf.seek(0)

        # filename with current date
        now = datetime.datetime.now()
        filename = f"{now.strftime('%B')}_{now.day}_superdispatch.csv"

        return send_file(
            buf,
            as_attachment=True,
            download_name=filename,
            mimetype='text/csv'
        )

    return render_template('index.html')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5001)
