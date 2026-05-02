from flask import Flask, request, send_file, render_template_string
from pypdf import PdfReader
import pandas as pd
import openpyxl
from openpyxl.styles import Font
import io
import re
import traceback
from datetime import datetime

app = Flask(__name__)

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>BCA PDF to Excel Converter</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; display: flex; align-items: center; justify-content: center; height: 100vh; background-color: #f0f2f5; }
        .container { background: white; padding: 30px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); text-align: center; width: 400px; }
        h2 { color: #0056b3; margin-bottom: 20px; }
        input[type="file"] { margin: 20px 0; border: 1px dashed #0056b3; padding: 10px; width: 100%; border-radius: 5px; }
        button { background-color: #0056b3; color: white; padding: 12px 24px; border: none; border-radius: 6px; cursor: pointer; font-weight: bold; width: 100%; }
        button:hover { background-color: #004494; }
        .footer { margin-top: 20px; font-size: 12px; color: #666; }
    </style>
</head>
<body>
    <div class="container">
        <h2>BCA Converter</h2>
        <form action="/convert" method="post" enctype="multipart/form-data">
            <input type="file" name="file" accept=".pdf" required>
            <button type="submit">Konversi Sekarang</button>
        </form>
        <div class="footer">Unggah file e-statement BCA (.pdf)</div>
    </div>
</body>
</html>
"""

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return "Tidak ada file", 400
    
    file = request.files['file']
    try:
        pdf_bytes = io.BytesIO(file.read())
        extracted_data = parse_bca_pdf_robust(pdf_bytes)
        
        excel_io = create_excel_template(extracted_data)
        
        return send_file(
            excel_io,
            as_attachment=True,
            download_name=f"MUTASI_BCA_{datetime.now().strftime('%d%m%Y')}.xlsx",
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        error_details = traceback.format_exc()
        return f"<h3>Terjadi Kesalahan</h3><pre>{error_details}</pre>", 500

def parse_bca_pdf_robust(pdf_stream):
    data = []
    saldo_awal = 0
    tahun = str(datetime.now().year)
    
    reader = PdfReader(pdf_stream)
    
    for page in reader.pages:
        text = page.extract_text()
        if not text:
            continue
            
        lines = text.split('\n')
        for line in lines:
            line = line.strip()
            
            # Cari Tahun
            if "PERIODE" in line.upper():
                match_tahun = re.search(r'20\d{2}', line)
                if match_tahun:
                    tahun = match_tahun.group(0)

            # Cari Saldo Awal
            if "SALDO AWAL" in line.upper():
                # Mencari angka di akhir baris (format: 1.234.567,89)
                match_saldo = re.search(r'([\d\.,]+)$', line)
                if match_saldo:
                    try:
                        saldo_awal = float(match_saldo.group(1).replace('.', '').replace(',', '.'))
                    except: pass
            
            # Cari baris transaksi (Regex disesuaikan untuk output pypdf)
            # Pola: Tgl(DD/MM) Keterangan (Spasi) Nominal (Spasi) Saldo
            # Contoh: "01/01 KREDIT OTOMATIS 500.000,00 1.500.000,00"
            match_trx = re.search(r'^(\d{2}/\d{2})\s+(.*?)\s+([\d\.,]+)\s+([\d\.,]+)$', line)
            
            if match_trx:
                tanggal = match_trx.group(1) + f"/{tahun}"
                keterangan = match_trx.group(2).strip()
                # Ganti format ribuan titik ke standar python (1.000,00 -> 1000.00)
                mutasi_str = match_trx.group(3).replace('.', '').replace(',', '.')
                saldo_str = match_trx.group(4).replace('.', '').replace(',', '.')
                
                try:
                    mutasi = float(mutasi_str)
                    saldo = float(saldo_str)
                    
                    # Logika Debet/Kredit sederhana
                    kredit = mutasi if "CR" in keterangan or "DB" not in keterangan else 0
                    debet = mutasi if kredit == 0 else 0

                    data.append({
                        "tanggal": tanggal,
                        "keterangan": keterangan,
                        "debet": debet,
                        "kredit": kredit,
                        "saldo": saldo
                    })
                except:
                    continue
    
    return {"saldo_awal": saldo_awal, "tahun": tahun, "trx": data}

def create_excel_template(data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Mutasi"

    # Penulisan Header
    ws['A1'] = "BCA 346-8383111"
    ws['A2'] = "CV. MITRA JAYA ANUGERAH"
    ws['A3'] = f"TAHUN {data.get('tahun', '2025')}"
    
    headers = ["NO", "TANGGAL", "NAMA", "KETERANGAN", "KODE", "DEBET", "KREDIT", "SALDO"]
    ws.append([]) # Baris 4
    ws.append(headers) # Baris 5
    
    for col in range(1, 9):
        ws.cell(row=5, column=col).font = Font(bold=True)

    ws.append(["", "", "", "SALDO AWAL", "", "", "", data['saldo_awal']])
    
    for idx, row in enumerate(data['trx'], start=1):
        kode = 5 if row['kredit'] > 0 else ""
        ws.append([
            idx,
            row['tanggal'],
            "", 
            row['keterangan'],
            kode,
            row['debet'] if row['debet'] > 0 else 0,
            row['kredit'] if row['kredit'] > 0 else 0,
            row['saldo']
        ])

    excel_io = io.BytesIO()
    wb.save(excel_io)
    excel_io.seek(0)
    return excel_io

if __name__ == '__main__':
    app.run(debug=True)
