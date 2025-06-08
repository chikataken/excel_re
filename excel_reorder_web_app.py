#!/usr/bin/env python3
"""
Flask Web App for Excel Column Reorderer

This script provides a lightweight web interface for users to upload an Excel file,
reorder and filter columns based on your predefined preference, apply multi-level
sorting, freeze panes, set column widths, and download the processed file.

---
**Quickstart & Deployment Guide**

1. **Clone or copy this script** to your Linux server, e.g. `/opt/excel_app/`:
   ```bash
   mkdir -p /opt/excel_app && cd /opt/excel_app
   # Copy excel_reorder_web_app.py here
   ```

2. **Create and activate a Python virtual environment**:
   ```bash
   python3 -m venv venv
   source venv/bin/activate
   ```

3. **Install runtime dependencies**:
   ```bash
   pip install flask pandas xlsxwriter openpyxl gunicorn
   ```

4. **Test locally on port 5000**:
   ```bash
   python excel_reorder_web_app.py
   # Visit http://localhost:5000 in your browser
   ```

5. **Configure Gunicorn** to serve the app on an internal port (e.g. 8000):
   ```bash
   source venv/bin/activate
   gunicorn --workers 3 --bind 127.0.0.1:8000 excel_reorder_web_app:app
   ```

6. **Set up a systemd service** (optional, for automatic start/restart):
   Create `/etc/systemd/system/excel_reorder.service`:
   ```ini
   [Unit]
   Description=Excel Reorder Web App
   After=network.target

   [Service]
   User=www-data
   WorkingDirectory=/opt/excel_app
   ExecStart=/opt/excel_app/venv/bin/gunicorn \
       --workers 3 \
       --bind 127.0.0.1:8000 \
       excel_reorder_web_app:app
   Restart=always

   [Install]
   WantedBy=multi-user.target
   ```
   ```bash
   systemctl daemon-reload
   systemctl enable excel_reorder.service
   systemctl start excel_reorder.service
   ```

7. **Configure Nginx as a reverse proxy** for `wastake.com`:
   - Create `/etc/nginx/sites-available/wastake.com`:
     ```nginx
     server {
         listen 80;
         server_name wastake.com;

         location / {
             proxy_pass http://127.0.0.1:8000;
             proxy_set_header Host $host;
             proxy_set_header X-Real-IP $remote_addr;
         }
     }
     ```
   - Enable and reload:
     ```bash
     ln -s /etc/nginx/sites-available/wastake.com /etc/nginx/sites-enabled/
     nginx -t && systemctl reload nginx
     ```

8. **Enable HTTPS with Let’s Encrypt**:
   ```bash
   apt update && apt install certbot python3-certbot-nginx
   certbot --nginx -d wastake.com
   ```

9. **Verify in browser**:
   Navigate to `https://wastake.com/`, upload an Excel file, and download the reordered output.

---
"""
import os
import io
import pandas as pd
from flask import Flask, request, send_file, render_template_string

app = Flask(__name__)

# Preference and formatting settings (same as GUI version)
PREFERENCE_COLUMNS = [
    "ShipmentNumber","Vin","OriginState","OriginCity",
    "OriginAddress","OriginZip","OriginContactPhone",
    "DestinationState","DestinationCity","DestinationAddress",
    "DestinationZip","DestinationContactPhone"
]
FIRST_COLUMNS = ["Vin","OriginState","OriginCity","DestinationState","DestinationCity","Price"]
OTHER_COLUMNS = [col for col in PREFERENCE_COLUMNS if col not in FIRST_COLUMNS]
SORT_ORDER = ["OriginState","OriginCity","DestinationState","DestinationCity","Vin"]
FIRST_COLUMN_WIDTH = 25
OTHER_COLUMN_WIDTH = 17

# HTML template
HTML = '''
<!doctype html>
<title>Excel Reorderer</title>
<h1>Upload Excel File</h1>
<form method=post enctype=multipart/form-data>
  <input type=file name=file accept=".xls,.xlsx">
  <input type=submit value=Process>
</form>
'''  

def process_df(df):
    # Add Price column
    if "Price" not in df.columns:
        df["Price"] = ""
    # Filter to preference + Price
    cols = [c for c in PREFERENCE_COLUMNS + ["Price"] if c in df.columns]
    df = df[cols]
    # Reorder columns
    order = [c for c in FIRST_COLUMNS if c in df.columns] + [c for c in OTHER_COLUMNS if c in df.columns]
    df = df[order]
    # Sort
    sort_cols = [c for c in SORT_ORDER if c in df.columns]
    if sort_cols:
        df = df.sort_values(by=sort_cols)
    return df


def excel_bytes(df):
    # Write df to Excel in-memory with formatting
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        ws = writer.sheets['Sheet1']
        max_row, max_col = df.shape
        ws.autofilter(0,0,max_row,max_col-1)
        ws.freeze_panes(1,0)
        ws.set_column(0,0,FIRST_COLUMN_WIDTH)
        if max_col>1:
            ws.set_column(1, max_col-1, OTHER_COLUMN_WIDTH)
    output.seek(0)
    return output

@app.route('/', methods=['GET','POST'])
def upload():
    if request.method == 'POST':
        f = request.files.get('file')
        if not f:
            return 'No file uploaded', 400
        df = pd.read_excel(f)
        df_out = process_df(df)
        output = excel_bytes(df_out)
        return send_file(
            output,
            as_attachment=True,
            download_name='processed.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    return render_template_string(HTML)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
