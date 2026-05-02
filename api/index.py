from flask import Flask, request, send_file, render_template_string
import pdfplumber
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
    <title>Konverter PDF BCA ke Excel</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 50px; text-align: center; }
        .container { max-width: 500px; margin: auto; padding: 20px; border: 1px solid #ccc; border-radius: 10px; }
        input[type="file"] { margin: 20px 0; }
        button { background-color: #0066cc; color: white; padding: 10px 20px; border: none; border-radius: 5px; cursor: pointer; }
        button:hover { background-color: #004c99; }
        .error-box { background-color: #ffe6e6; color: #cc0000; padding: 15px; border-radius: 5px; text-align: left; overflow-x: auto; margin-top: 20px; }
    </style>
</head>
<body>
    <div class="container">
        <h2>Upload Mutasi Rekening BCA (PDF)</h2>
        <form action="/convert" method="post" enctype="multipart/form-data">
            <input type="file" name="file" accept=".pdf" required>
            <br>
            <button type="submit">Konversi ke Excel</button>
        </form>
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
        return "Tidak ada file yang diunggah", 400
    
    file = request.files['file']
    if file.filename == '':
        return "File tidak valid", 400

    try:
        # Proses PDF
        pdf_bytes = file.read()
        extracted_data = parse_bca_pdf(io.BytesIO(pdf_bytes))
        
        # Buat Excel
        excel_io = create_excel_template(extracted_data)
        
        # Kirim file kembali ke pengguna
        return send_file(
            excel_io,
            as_attachment=True,
            download_name=f"BANK_BCA_{datetime.now().strftime('%Y_%m')}.xlsx",
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        # Menangkap error lengkap (traceback) agar mudah di-debug
        error_details = traceback.format_exc()
        error_html = f"""
        <div style="font-family: Arial, sans-serif; padding: 20px;">
            <h2 style="color: red;">Terjadi Kesalahan!</h2>
            <p>Sistem gagal memproses PDF. Berikut adalah detail errornya:</p>
            <div style="background-color: #f8f9fa; padding: 15px; border: 1px solid #ccc; border-radius: 5px; overflow-x: auto;">
                <pre style="margin: 0;">{error_details}</pre>
            </div>
            <br>
            <a href="/" style="text-decoration: none; background-color: #0066cc; color: white; padding: 10px 15px; border-radius: 5px;">Kembali</a>
        </div>
        """
        return error_html, 500

def parse_bca_pdf(pdf_stream):
    data = []
    saldo_awal = 0
    tahun = str(datetime.now().year)
    
    with pdfplumber.open(pdf_stream) as pdf:
        if not pdf.pages:
            raise ValueError("File PDF kosong atau tidak terbaca.")
            
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
                
            lines = text.split('\n')
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                    
                # Cari Periode untuk mendapatkan Tahun
                if "PERIODE" in line.upper():
                    match_tahun = re.search(r'20\d{2}', line)
                    if match_tahun:
                        tahun = match_tahun.group(0)

                # Cari Saldo Awal
                if "SALDO AWAL" in line.upper():
                    match_saldo = re.search(r'([\d,]+\.\d{2})', line)
                    if match_saldo:
                        saldo_awal = float(match_saldo.group(1).replace(',', ''))
                
                # Cari baris transaksi (Sangat ketat dan aman dari IndexError)
                match_trx = re.match(r'^(\d{2}/\d{2})\s+(.*?)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})$', line)
                if match_trx:
                    tanggal = match_trx.group(1) + f"/{tahun}"
                    keterangan = match_trx.group(2).strip()
                    mutasi_str = match_trx.group(3).replace(',', '')
                    saldo_str = match_trx.group(4).replace(',', '')
                    
                    try:
                        mutasi = float(mutasi_str)
                        saldo = float(saldo_str)
                    except ValueError:
                        continue # Lewati jika gagal diubah ke angka
                    
                    kredit = mutasi if "CR" in keterangan or "PENERIMAAN" in keterangan.upper() else 0
                    debet = mutasi if kredit == 0 else 0
                    
                    try:
                        tgl_obj = datetime.strptime(tanggal, "%d/%m/%Y")
                        tgl_str = tgl_obj.strftime("%Y-%m-%d")
                    except ValueError:
                        tgl_str = tanggal

                    data.append({
                        "tanggal": tgl_str,
                        "keterangan": keterangan,
                        "debet": debet,
                        "kredit": kredit,
                        "saldo": saldo
                    })
    
    return {"saldo_awal": saldo_awal, "tahun": tahun, "trx": data}

def create_excel_template(data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Mutasi"

    ws['A1'] = "BCA 346-8383111"
    ws['A2'] = "CV. MITRA JAYA ANUGERAH"
    ws['A3'] = f"TAHUN {data.get('tahun', '2025')}"
    
    headers = ["NO", "TANGGAL", "NAMA", "KETERANGAN", "KODE", "DEBET", "KREDIT", "SALDO"]
    ws.append([])
    ws.append(headers)
    
    for col in range(1, 9):
        ws.cell(row=5, column=col).font = Font(bold=True)

    ws.append(["", "", "", "SALDO AWAL", "", "", "", data['saldo_awal']])
    ws.append([])
    
    for idx, row in enumerate(data['trx'], start=1):
        kode = 5 if row['kredit'] > 0 else ""
        ws.append([
            idx,
            row['tanggal'],
            "", 
            row['keterangan'],
            kode,
            row['debet'] if row['debet'] > 0 else "",
            row['kredit'] if row['kredit'] > 0 else "",
            row['saldo']
        ])

    excel_io = io.BytesIO()
    wb.save(excel_io)
    excel_io.seek(0)
    
    return excel_io

if __name__ == '__main__':
    app.run(debug=True, port=5000)