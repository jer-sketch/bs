from flask import Flask, request, send_file, render_template_string
import pdfplumber
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment
import io
import re
from datetime import datetime

app = Flask(__name__)

# Template HTML sederhana untuk UI Upload
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
        return f"Terjadi kesalahan saat memproses: {str(e)}", 500

def parse_bca_pdf(pdf_stream):
    """
    Logika ekstraksi baris demi baris dari PDF Rekening Koran BCA.
    Perlu disesuaikan jika format PDF memiliki variasi.
    """
    data = []
    saldo_awal = 0
    tahun = datetime.now().year # Default tahun
    
    with pdfplumber.open(pdf_stream) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
                
            lines = text.split('\n')
            for line in lines:
                # Cari Periode untuk mendapatkan Tahun (contoh: "PERIODE : JANUARI 2025")
                if "PERIODE" in line:
                    match_tahun = re.search(r'20\d{2}', line)
                    if match_tahun:
                        tahun = match_tahun.group(0)

                # Cari Saldo Awal
                if "SALDO AWAL" in line:
                    match_saldo = re.search(r'([\d,]+\.\d{2})$', line)
                    if match_saldo:
                        saldo_awal = float(match_saldo.group(1).replace(',', ''))
                
                # Cari baris transaksi (dimulai dengan tanggal DD/MM)
                match_trx = re.match(r'^(\d{2}/\d{2})\s+(.*?)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})$', line)
                if match_trx:
                    tanggal = match_trx.group(1) + f"/{tahun}" # DD/MM/YYYY
                    keterangan = match_trx.group(2).strip()
                    mutasi = float(match_trx.group(3).replace(',', ''))
                    saldo = float(match_trx.group(4).replace(',', ''))
                    
                    # Tentukan Debit / Kredit (Simplifikasi)
                    # Bisa disempurnakan dengan mengecek selisih saldo sebelumnya vs sekarang
                    kredit = mutasi if "CR" in keterangan or "PENERIMAAN" in keterangan.upper() else 0
                    debet = mutasi if kredit == 0 else 0
                    
                    # Parse tanggal ke format YYYY-MM-DD
                    try:
                        tgl_obj = datetime.strptime(tanggal, "%d/%m/%Y")
                        tgl_str = tgl_obj.strftime("%Y-%m-%d")
                    except:
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
    """
    Membuat file Excel persis dengan format CSV/XLSX yang diminta.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Mutasi"

    # Header Template Atas
    ws['A1'] = "BCA 346-8383111"
    ws['A2'] = "CV. MITRA JAYA ANUGERAH"
    ws['A3'] = f"TAHUN {data.get('tahun', '2025')}"
    
    # Border pseudo-tabel (opsional, disesuaikan dengan permintaan Anda yang seperti tabel teks)
    headers = ["NO", "TANGGAL", "NAMA", "KETERANGAN", "KODE", "DEBET", "KREDIT", "SALDO"]
    ws.append([]) # Baris kosong (baris 4)
    ws.append(headers) # Baris 5
    
    # Bold header
    for col in range(1, 9):
        ws.cell(row=5, column=col).font = Font(bold=True)

    # Baris Saldo Awal
    ws.append(["", "", "", "SALDO AWAL", "", "", "", data['saldo_awal']])
    ws.append([]) # Baris kosong
    
    # Isi data transaksi
    for idx, row in enumerate(data['trx'], start=1):
        # Default kode = 5 untuk penerimaan (CR), bisa disesuaikan dengan aturan bisnis Anda
        kode = 5 if row['kredit'] > 0 else ""
        
        ws.append([
            idx,
            row['tanggal'],
            "", # NAMA dibiarkan kosong sesuai contoh
            row['keterangan'],
            kode,
            row['debet'] if row['debet'] > 0 else "",
            row['kredit'] if row['kredit'] > 0 else "",
            row['saldo']
        ])

    # Simpan ke byte stream
    excel_io = io.BytesIO()
    wb.save(excel_io)
    excel_io.seek(0)
    
    return excel_io

# Untuk testing lokal
if __name__ == '__main__':
    app.run(debug=True, port=5000)