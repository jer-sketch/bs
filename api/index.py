from flask import Flask, request, send_file, render_template_string
from pypdf import PdfReader
import openpyxl
from openpyxl.styles import Font, Alignment
import io
import re
import traceback
from datetime import datetime

# INISIALISASI APP
app = Flask(__name__)

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>BCA PDF to Excel Converter - Pro</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; display: flex; align-items: center; justify-content: center; height: 100vh; background-color: #eef2f7; }
        .container { background: white; padding: 40px; border-radius: 15px; box-shadow: 0 10px 25px rgba(0,0,0,0.1); text-align: center; width: 450px; }
        h2 { color: #0056b3; margin-bottom: 10px; }
        p { color: #666; font-size: 14px; margin-bottom: 30px; }
        .upload-area { border: 2px dashed #0056b3; padding: 20px; border-radius: 10px; background: #f8fbff; margin-bottom: 20px; }
        input[type="file"] { margin: 10px 0; width: 100%; }
        button { background-color: #0056b3; color: white; padding: 14px; border: none; border-radius: 8px; cursor: pointer; font-weight: bold; width: 100%; font-size: 16px; transition: 0.3s; }
        button:hover { background-color: #004494; transform: translateY(-2px); }
        .footer { margin-top: 20px; font-size: 11px; color: #999; }
    </style>
</head>
<body>
    <div class="container">
        <h2>BCA Converter Pro</h2>
        <p>Ekstrak Mutasi PDF ke Excel dengan Cepat</p>
        <form action="/convert" method="post" enctype="multipart/form-data">
            <div class="upload-area">
                <input type="file" name="file" accept=".pdf" required>
            </div>
            <button type="submit">Konversi ke Excel</button>
        </form>
        <div class="footer">E-Statement BCA Safe Processor</div>
    </div>
</body>
</html>
"""

def clean_bca_money(value_str):
    """
    Menghapus semua karakter non-angka dan mengubahnya menjadi float.
    Sangat ampuh untuk PDF yang karakter titik/komanya berantakan.
    """
    if not value_str:
        return 0.0
    # Ambil hanya angka saja
    only_digits = re.sub(r'[^\d]', '', value_str)
    if not only_digits:
        return 0.0
    # BCA selalu punya 2 digit desimal (sen) di akhir
    return float(only_digits) / 100

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return "File tidak ditemukan", 400
    
    file = request.files['file']
    try:
        pdf_bytes = io.BytesIO(file.read())
        extracted_data = parse_bca_pdf_robust(pdf_bytes)
        
        excel_io = create_excel_template(extracted_data)
        
        return send_file(
            excel_io,
            as_attachment=True,
            download_name=f"MUTASI_BCA_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        error_details = traceback.format_exc()
        return f"<h3>Terjadi Kesalahan System</h3><p>Detail:</p><pre>{error_details}</pre>", 500

def parse_bca_pdf_robust(pdf_stream):
    data = []
    saldo_awal = 0.0
    tahun = str(datetime.now().year)
    
    reader = PdfReader(pdf_stream)
    full_text = ""
    for page in reader.pages:
        full_text += page.extract_text() + "\n"

    # Cari Tahun
    match_tahun = re.search(r'PERIODE\s*:\s*.*\s+(20\d{2})', full_text, re.IGNORECASE)
    if match_tahun:
        tahun = match_tahun.group(1)

    # Cari Saldo Awal - Menggunakan fungsi pembersih baru
    match_saldo_awal = re.search(r'SALDO AWAL\s+([\d\.,]+)', full_text, re.IGNORECASE)
    if match_saldo_awal:
        saldo_awal = clean_bca_money(match_saldo_awal.group(1))

    lines = full_text.split('\n')
    for line in lines:
        line = line.strip()
        # Regex untuk menangkap baris transaksi BCA
        match_trx = re.search(r'^(\d{2}/\d{2})\s+(.*?)\s+([\d\.,]+)\s+([\d\.,]+)$', line)
        
        if match_trx:
            tgl_short = match_trx.group(1)
            keterangan = match_trx.group(2).strip()
            mutasi_raw = match_trx.group(3)
            saldo_raw = match_trx.group(4)
            
            try:
                # Bersihkan angka dengan fungsi sakti clean_bca_money
                mutasi = clean_bca_money(mutasi_raw)
                saldo = clean_bca_money(saldo_raw)
                
                # Cek Kredit/Debet
                is_kredit = "CR" in keterangan.upper() or "CR" in mutasi_raw.upper()
                kredit = mutasi if is_kredit else 0
                debet = mutasi if not is_kredit else 0

                data.append({
                    "tanggal": f"{tgl_short}/{tahun}",
                    "keterangan": re.sub(r'\s+(CR|DB)$', '', keterangan, flags=re.IGNORECASE).strip(),
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

    # Header Atas
    ws['A1'] = "BCA 346-8383111"
    ws['A2'] = "CV. MITRA JAYA ANUGERAH"
    ws['A3'] = f"TAHUN {data['tahun']}"
    
    headers = ["NO", "TANGGAL", "NAMA", "KETERANGAN", "KODE", "DEBET", "KREDIT", "SALDO"]
    ws.append([]) 
    ws.append(headers) 
    
    for col in range(1, 9):
        cell = ws.cell(row=5, column=col)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = openpyxl.styles.PatternFill(start_color="0056b3", end_color="0056b3", fill_type="solid")
        cell.alignment = Alignment(horizontal='center')

    # Baris Saldo Awal
    ws.append(["", "", "", "SALDO AWAL", "", "", "", data['saldo_awal']])
    ws.cell(row=6, column=8).number_format = '#,##0.00'
    
    # Isi Data
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
        
        # Format angka agar ada pemisah ribuan di Excel
        curr_row = 6 + idx
        ws.cell(row=curr_row, column=6).number_format = '#,##0.00'
        ws.cell(row=curr_row, column=7).number_format = '#,##0.00'
        ws.cell(row=curr_row, column=8).number_format = '#,##0.00'

    # Auto-size kolom (sederhana)
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 15

    excel_io = io.BytesIO()
    wb.save(excel_io)
    excel_io.seek(0)
    return excel_io

if __name__ == '__main__':
    app.run(debug=True)