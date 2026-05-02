from flask import Flask, request, send_file, render_template_string
import pdfplumber
import re
import io
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from datetime import datetime

app = Flask(__name__)

MONTH_MAP = {'JANUARI':1,'FEBRUARI':2,'MARET':3,'APRIL':4,'MEI':5,'JUNI':6,
             'JULI':7,'AGUSTUS':8,'SEPTEMBER':9,'OKTOBER':10,'NOVEMBER':11,'DESEMBER':12}

def classify(line, is_credit):
    u = line.upper()
    if 'BIAYA ADM' in u: return 'BIAYA ADM BANK', '27'
    if 'PAJAK BUNGA' in u or 'PAJAK JASA' in u: return 'PAJAK JASA GIRO', ''
    if 'BUNGA' in u and not ('PAJAK' in u): return 'BUNGA', ''
    if 'TRANSFER' in u and 'BIAYA' in u: return 'BIAYA TRANSFER', '27'
    if 'TELKOM' in u or 'TELEPON' in u: return 'BIAYA TELEPON', '27'
    if 'LISTRIK' in u or 'PLN' in u: return 'BIAYA LISTRIK', '27'
    if 'GAJI' in u: return 'BIAYA GAJI PEGAWAI', ''
    if 'SETORAN TUNAI' in u: return 'SETORAN TUNAI', '1'
    
    # Dibuat lebih kebal terhadap spasi ganda
    if 'PENERIMAAN ' in u and 'NEGARA' in u: return 'PENERIMAAN NEGARA', ''
    
    if not is_credit: return 'PELUNASAN HUTANG DAGANG', ''
    
    # Jika masuk uang dan bukan hal di atas, maka Penjualan
    return 'PENERIMAAN PENJUALAN', '3' 

def extract_nama(desc):
    """
    Fungsi khusus untuk membuang keyword bank dan nomor referensi,
    sehingga hanya menyisakan Nama Orang atau Perusahaan.
    """
    # 1. Buang angka nominal di ujung kalimat
    clean = re.sub(r'[\d,]+\.\d{2}.*$', '', desc).strip()
    
    # 2. Buang tanda CR atau DB
    clean = re.sub(r'\s+(CR|DB)\b', '', clean, flags=re.IGNORECASE)
    
    # 3. Buang kata-kata mutasi sistem bank
    bank_keywords = [
        r'TRSF E-BANKING', r'SWITCHING', r'KR OTOMATIS', r'DR OTOMATIS',
        r'BIF TRANSFER DR', r'BIF TRANSFER CR', r'BIF TRANSFER',
        r'TRANSFER DR', r'TRANSFER CR', r'SETORAN TUNAI', r'BYR VIA',
        r'M-BCA', r'KARTU KREDIT', r'AUTO-DEBET'
    ]
    for kw in bank_keywords:
        clean = re.sub(kw, '', clean, flags=re.IGNORECASE)
        
    # 4. Buang kata yang mengandung angka (seperti nomor ref: 0105/FTSCY/WS9)
    clean = re.sub(r'\S*\d+\S*', '', clean)
    
    # 5. Rapikan spasi dan simbol sisa
    clean = clean.strip(' -/:')
    return re.sub(r'\s+', ' ', clean).strip()

def parse_bca_pdf_logic(pdf_stream):
    all_lines = []
    with pdfplumber.open(pdf_stream) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                all_lines.extend(text.split('\n'))

    period = ''
    saldo_awal = saldo_akhir = mut_cr = mut_db = 0
    
    for l in all_lines:
        if 'PERIODE :' in l and not period:
            period = l.split(':', 1)[1].strip()
        m = re.search(r'SALDO AWAL\s*:\s*([\d,]+\.\d+)', l)
        if m: saldo_awal = float(m.group(1).replace(',',''))
        m = re.search(r'MUTASI CR\s*:\s*([\d,]+\.\d+)', l)
        if m: mut_cr = float(m.group(1).replace(',',''))
        m = re.search(r'MUTASI DB\s*:\s*([\d,]+\.\d+)', l)
        if m: mut_db = float(m.group(1).replace(',',''))
        m = re.search(r'SALDO AKHIR\s*:\s*([\d,]+\.\d+)', l)
        if m: saldo_akhir = float(m.group(1).replace(',',''))

    parts = period.strip().split()
    year = parts[1] if len(parts) > 1 else str(datetime.now().year)

    DATE_RE = re.compile(r'^(\d{2})/(\d{2})\s+(.+)')
    header_starts = ('REKENING GIRO','KCU ','MITRA JAYA','BOJONGLOA','SITUSAEUR',
                     'JL LEUWI','BANDUNG','INDONESIA','NO. REKENING','HALAMAN',
                     'PERIODE','MATA UANG','CATATAN','Apabila','Rekening ini',
                     'telah menyetujui','BCA berhak','Laporan Mutasi',
                     'TANGGAL KETERANGAN','SALDO AWAL :','MUTASI CR','MUTASI DB','SALDO AKHIR')

    txs = []
    for l in all_lines:
        l = l.strip()
        dm = DATE_RE.match(l)
        if not dm: continue
        
        day, mon, rest = dm.group(1), dm.group(2), dm.group(3).strip()
        if 'SALDO AWAL' in rest or 'Bersambung' in rest: continue
        if any(rest.startswith(h) for h in header_starts): continue

        date = f'{day}/{mon}/{year}'
        all_amounts = re.findall(r'([\d,]+\.\d{2})', rest)
        if not all_amounts: continue

        amt = float(all_amounts[0].replace(',',''))
        u = rest.upper()
        
        is_cr = (' CR ' in u or 'SETORAN TUNAI' in u or 'KR OTOMATIS' in u or 'SWITCHING CR' in u)
        is_db = (' DB ' in u or 'BYR VIA' in u)

        if any(x in u for x in ['BIAYA ADM', 'PAJAK BUNGA', 'PAJAK JASA']):
            is_db, is_cr = True, False
        if 'BUNGA' in u and not ('PAJAK' in u):
            is_db, is_cr = False, True

        # Klasifikasi kategori
        ket, kode = classify(rest, is_cr and not is_db)
        
        # Ekstraksi Nama
        nama_orang = extract_nama(rest)
        
        debet = amt if is_db else 0
        kredit = amt if (is_cr and not is_db) else 0
        if debet == 0 and kredit == 0: kredit = amt

        txs.append({
            'date': date, 
            'nama': nama_orang, 
            'ket': ket, 
            'kode': kode, 
            'debet': debet, 
            'kredit': kredit
        })

    return {
        "txs": txs, "period": period, "saldo_awal": saldo_awal, 
        "saldo_akhir": saldo_akhir, "mut_db": mut_db, "mut_cr": mut_cr
    }

def create_excel_output(data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Mutasi BCA'
    
    # Styles
    hdr_font = Font(name='Arial', bold=True, size=10)
    data_font = Font(name='Arial', size=10)
    num_fmt = '#,##0.00'
    center_align = Alignment(horizontal='center')

    ws.append(['BCA 346-8383111'])
    ws.append(['CV. MITRA JAYA ANUGERAH'])
    ws.append([f"PERIODE: {data['period']}"])
    ws.append([]) # Baris Kosong
    
    # Header Tabel
    headers = ['NO', 'TANGGAL', 'NAMA', 'KETERANGAN', 'KODE', 'DEBET', 'KREDIT', 'SALDO']
    ws.append(headers)
    
    # Formatting Header Tabel
    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=5, column=col_num)
        cell.font = Font(name='Arial', bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="0056b3", end_color="0056b3", fill_type="solid")
        cell.alignment = center_align
    
    # Baris Saldo Awal
    ws.append(['', '', '', 'SALDO AWAL', '', '', '', data['saldo_awal']])
    ws.cell(row=6, column=8).number_format = num_fmt
    
    current_saldo = data['saldo_awal']
    for idx, tx in enumerate(data['txs'], 1):
        if tx['kredit'] > 0:
            current_saldo = round(current_saldo + tx['kredit'], 2)
        else:
            current_saldo = round(current_saldo - tx['debet'], 2)
            
        ws.append([
            idx, 
            tx['date'], 
            tx['nama'],  # <--- Nama masuk ke kolom C (NAMA)
            tx['ket'],   # <--- Klasifikasi masuk ke kolom D (KETERANGAN)
            tx['kode'],
            tx['debet'] if tx['debet'] > 0 else '', 
            tx['kredit'] if tx['kredit'] > 0 else '',
            current_saldo
        ])

    # Footer Tabel
    ws.append([])
    ws.append(['', '', '', 'TOTAL MUTASI', '', data['mut_db'], data['mut_cr'], data['saldo_akhir']])
    
    # Menerapkan Font dan Format Angka pada data
    for row in ws.iter_rows(min_row=7):
        for cell in row:
            cell.font = data_font
            # Jika sel ada di kolom Debet (6), Kredit (7), Saldo (8) dan berupa angka
            if cell.column in [6, 7, 8] and isinstance(cell.value, (int, float)):
                cell.number_format = num_fmt

    # Lebar Kolom yang Nyaman
    widths = [5, 12, 25, 28, 8, 16, 16, 18]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    excel_io = io.BytesIO()
    wb.save(excel_io)
    excel_io.seek(0)
    return excel_io

@app.route('/')
def index():
    return render_template_string("""
    <!DOCTYPE html>
    <html lang="id">
    <head>
        <meta charset="UTF-8">
        <title>Konverter BCA ke Excel</title>
        <style>
            body { font-family: 'Segoe UI', Tahoma, sans-serif; background-color: #f4f7f6; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; }
            .card { background: white; padding: 40px; border-radius: 12px; box-shadow: 0 4px 20px rgba(0,0,0,0.1); text-align: center; width: 350px; }
            h2 { color: #0056b3; margin-top: 0; }
            input[type="file"] { margin: 20px 0; width: 100%; border: 2px dashed #ccc; padding: 10px; box-sizing: border-box; }
            button { background: #0056b3; color: white; border: none; padding: 12px; width: 100%; border-radius: 6px; cursor: pointer; font-weight: bold; font-size: 16px; }
            button:hover { background: #004494; }
        </style>
    </head>
    <body>
        <div class="card">
            <h2>Konverter BCA</h2>
            <form action="/convert" method="post" enctype="multipart/form-data">
                <input type="file" name="file" accept=".pdf" required>
                <button type="submit">Konversi ke Excel</button>
            </form>
        </div>
    </body>
    </html>
    """)

@app.route('/convert', methods=['POST'])
def convert():
    file = request.files['file']
    if not file: return "File tidak ditemukan", 400
    
    pdf_stream = io.BytesIO(file.read())
    data = parse_bca_pdf_logic(pdf_stream)
    excel_file = create_excel_output(data)
    
    return send_file(
        excel_file,
        as_attachment=True,
        download_name=f"MUTASI_BCA_{datetime.now().strftime('%d%m%Y')}.xlsx",
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

if __name__ == '__main__':
    app.run(debug=True)
