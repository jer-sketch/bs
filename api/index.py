from flask import Flask, request, send_file, render_template_string
from pypdf import PdfReader
import openpyxl
from openpyxl.styles import Font, Alignment
import io
import re
import traceback
from datetime import datetime

# INISIALISASI APP (PENTING UNTUK VERCEL)
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
    
    full_text = ""
    for page in reader.pages:
        full_text += page.extract_text() + "\n"

    # Cari Tahun
    match_tahun = re.search(r'PERIODE\s*:\s*.*\s+(20\d{2})', full_text, re.IGNORECASE)
    if match_tahun:
        tahun = match_tahun.group(1)

    # Cari Saldo Awal
    match_saldo_awal = re.search(r'SALDO AWAL\s+([\d\.,]+)', full_text, re.IGNORECASE)
    if match_saldo_awal:
        saldo_awal = float(match_saldo_awal.group(1).replace('.', '').replace(',', '.'))

    lines = full_text.split('\n')
    for line in lines:
        line = line.strip()
        # Pola BCA: Tanggal (DD/MM) - Keterangan - Nominal Mutasi - Saldo Akhir
        match_trx = re.search(r'^(\d{2}/\d{2})\s+(.*?)\s+([\d\.,]+)\s+([\d\.,]+)$', line)
        
        if match_trx:
            tgl_short = match_trx.group(1)
            keterangan = match_trx.group(2).strip()
            mutasi_raw = match_trx.group(3)
            saldo_raw = match_trx.group(4)
            
            try:
                mutasi = float(mutasi_raw.replace('.', '').replace(',', '.'))
                saldo = float(saldo_raw.replace('.', '').replace(',', '.'))
                
                # Cek jika ada CR (Credit) atau tanda uang masuk
                is_kredit = "CR" in keterangan.upper() or "CR" in mutasi_raw.upper()
                kredit = mutasi if is_kredit else 0
                debet = mutasi if not is_kredit else 0

                data.append({
                    "tanggal": f"{tgl_short}/{tahun}",
                    "keterangan": keterangan.replace(" CR", "").replace(" DB", "").replace("CR", "").replace("DB", "").strip(),
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

    ws['A1'] = "BCA 346-8383111"
    ws['A2'] = "CV. MITRA JAYA ANUGERAH"
    ws['A3'] = f"TAHUN {data['tahun']}"
    
    headers = ["NO", "TANGGAL", "NAMA", "KETERANGAN", "KODE", "DEBET", "KREDIT", "SALDO"]
    ws.append([]) 
    ws.append(headers) 
    
    for col in range(1, 9):
        cell = ws.cell(row=5, column=col)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')

    ws.append(["", "", "", "SALDO AWAL", "", "", "", data['saldo_awal']])
    
    for idx, row in enumerate(data['trx'], start=1):
        ws.append([
            idx,
            row['tanggal'],
            "", 
            row['keterangan'],
            5 if row['kredit'] > 0 else "",
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