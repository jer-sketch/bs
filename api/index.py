from flask import Flask, request, send_file, render_template_string
import pdfplumber
import re
import io
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from datetime import datetime

app = Flask(__name__)

def classify(full_text, is_credit):
    u = full_text.upper()
    
    # 1. Prioritas: Penerimaan Negara
    if 'PENERIMAAN NEGARA' in u: 
        return 'PENERIMAAN NEGARA', ''
    
    # 2. Bunga dan Pajak (Sesuai Permintaan: 28 dan 29)
    if 'PAJAK BUNGA' in u or 'PAJAK JASA' in u: 
        return 'PAJAK JASA GIRO', '29'
    if 'BUNGA' in u: 
        return 'BUNGA', '28'
        
    # 3. Biaya-biaya lain
    if 'BIAYA ADM' in u: return 'BIAYA ADM BANK', '27'
    if 'TRANSFER' in u and 'BIAYA' in u: return 'BIAYA TRANSFER', '27'
    if 'TELKOM' in u or 'TELEPON' in u: return 'BIAYA TELEPON', '27'
    if 'LISTRIK' in u or 'PLN' in u: return 'BIAYA LISTRIK', '27'
    if 'GAJI' in u: return 'BIAYA GAJI PEGAWAI', ''
    if 'SETORAN TUNAI' in u: return 'SETORAN TUNAI', '1'
    
    # 4. Default berdasarkan arus kas
    if not is_credit: 
        return 'PELUNASAN HUTANG DAGANG', ''
    
    return 'PENERIMAAN PENJUALAN', '3' 

def extract_nama_after_nominal(sub_lines, amount):
    """
    Mencari baris nominal, mengambil baris setelahnya, 
    lalu menghapus semua angka agar hanya tersisa TEXT saja.
    """
    amt_str = "{:,.2f}".format(amount).replace(',', '') 
    amt_comma = "{:,.2f}".format(amount) 
    
    found_idx = -1
    for i, line in enumerate(sub_lines):
        clean_line = line.replace(',', '')
        if amt_str in clean_line or amt_comma in line:
            found_idx = i
            break
    
    if found_idx != -1 and found_idx < len(sub_lines) - 1:
        # Ambil baris-baris setelah nominal
        nama_parts = sub_lines[found_idx+1:]
        text_raw = " ".join(nama_parts).upper()
        
        # Bersihkan kata kunci perbankan
        text_raw = re.sub(r'\b(TRSF|E-BANKING|DB|CR|BIF|SWITCHING|WS)\b', '', text_raw, flags=re.IGNORECASE)
        
        # HAPUS SEMUA ANGKA (Agar hanya teks saja yang tersisa)
        text_only = re.sub(r'\d+', '', text_raw)
        
        # Bersihkan simbol sisa (titik, koma, garing, spasi ganda)
        text_only = re.sub(r'[./-]', '', text_only)
        text_only = re.sub(r'\s+', ' ', text_only)
        
        return text_only.strip()
    
    return ""

def parse_bca_pdf_logic(pdf_stream):
    all_lines = []
    with pdfplumber.open(pdf_stream) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                all_lines.extend(text.split('\n'))

    period = ''
    saldo_awal = 0
    
    # Ambil Metadata Header
    for l in all_lines:
        if 'PERIODE :' in l and not period:
            period = l.split(':', 1)[1].strip()
        m = re.search(r'SALDO AWAL\s*:\s*([\d,]+\.\d+)', l)
        if m: saldo_awal = float(m.group(1).replace(',',''))

    year = period.strip().split()[-1] if period else str(datetime.now().year)

    # Kelompokkan baris berdasarkan Tanggal (Blok Transaksi)
    DATE_RE = re.compile(r'^(\d{2})/(\d{2})\s+')
    tx_blocks = []
    current_block = []

    for l in all_lines:
        l = l.strip()
        if DATE_RE.match(l):
            if current_block:
                tx_blocks.append(current_block)
            current_block = [l]
        elif current_block:
            current_block.append(l)
    if current_block:
        tx_blocks.append(current_block)

    txs = []
    for block in tx_blocks:
        full_block_text = " ".join(block).upper()
        
        # FILTER: Jangan masukkan baris saldo awal mutasi ke dalam daftar transaksi
        if 'SALDO AWAL' in full_block_text and len(block) < 3:
            continue

        m = DATE_RE.match(block[0])
        date_str = f"{m.group(1)}/{m.group(2)}/{year}"

        # Cari nominal transaksi (format 00.00)
        amounts_found = re.findall(r'([\d,]+\.\d{2})', full_block_text)
        if not amounts_found:
            continue
        
        amt = float(amounts_found[0].replace(',', ''))
        
        # Tentukan Debet/Kredit
        is_cr = (' CR' in full_block_text or 'SETORAN TUNAI' in full_block_text or 'KR OTOMATIS' in full_block_text or 'BUNGA' in full_block_text)
        is_db = (' DB' in full_block_text or 'BYR VIA' in full_block_text or 'BIAYA ADM' in full_block_text or 'PAJAK' in full_block_text)
        
        # Jika bunganya kredit tapi pajaknya debet
        if 'PAJAK' in full_block_text: is_db, is_cr = True, False
        if 'BUNGA' in full_block_text and 'PAJAK' not in full_block_text: is_db, is_cr = False, True

        # Klasifikasi kategori dan kode
        ket, kode = classify(full_block_text, is_cr and not is_db)
        
        # Ekstraksi Nama (Hanya teks setelah nominal)
        nama_orang = extract_nama_after_nominal(block, amt)
        
        # Bersihkan kolom nama jika itu pajak/bunga/adm agar Excel rapi
        if kode in ['27', '28', '29'] or ket == 'PENERIMAAN NEGARA':
            nama_orang = ''

        debet = amt if is_db else 0
        kredit = amt if (is_cr and not is_db) else 0
        if debet == 0 and kredit == 0: kredit = amt 

        txs.append({
            'date': date_str,
            'nama': nama_orang,
            'ket': ket,
            'kode': kode,
            'debet': debet,
            'kredit': kredit
        })

    return {"txs": txs, "period": period, "saldo_awal": saldo_awal}

def create_excel_output(data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Mutasi BCA'
    
    num_fmt = '#,##0.00'
    ws.append(['BCA 346-8383111'])
    ws.append(['CV. MITRA JAYA ANUGERAH'])
    ws.append([f"PERIODE: {data['period']}"])
    ws.append([]) 
    
    headers = ['NO', 'TANGGAL', 'NAMA', 'KETERANGAN', 'KODE', 'DEBET', 'KREDIT', 'SALDO']
    ws.append(headers)
    
    # Styling Header
    for col in range(1, 9):
        cell = ws.cell(row=5, column=col)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="0056b3", end_color="0056b3", fill_type="solid")
        cell.alignment = Alignment(horizontal='center')
    
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
            idx, tx['date'], tx['nama'], tx['ket'], tx['kode'],
            tx['debet'] if tx['debet'] > 0 else '', 
            tx['kredit'] if tx['kredit'] > 0 else '',
            current_saldo
        ])

    # Styling data
    for row in ws.iter_rows(min_row=6):
        for cell in row:
            cell.font = Font(name='Arial', size=10)
            if cell.column in [6, 7, 8] and isinstance(cell.value, (int, float)):
                cell.number_format = num_fmt

    widths = [5, 12, 30, 30, 8, 16, 16, 18]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    excel_io = io.BytesIO()
    wb.save(excel_io)
    excel_io.seek(0)
    return excel_io

@app.route('/')
def index():
    return render_template_string("""
    <body style="font-family:sans-serif; text-align:center; padding-top:50px; background:#f4f7f6;">
        <div style="display:inline-block; background:white; padding:40px; border-radius:15px; box-shadow:0 4px 15px rgba(0,0,0,0.1)">
            <h2 style="color:#0056b3;">BCA PDF to Excel</h2>
            <p>Pembersihan Nama (Teks Saja) & Kode Akuntansi</p>
            <form action="/convert" method="post" enctype="multipart/form-data">
                <input type="file" name="file" accept=".pdf" required style="margin-bottom:20px;"><br>
                <button type="submit" style="padding:12px 25px; background:#0056b3; color:white; border:none; border-radius:5px; cursor:pointer; font-weight:bold;">Download Excel</button>
            </form>
        </div>
    </body>
    """)

@app.route('/convert', methods=['POST'])
def convert():
    file = request.files['file']
    if not file: return "Error: File tidak ditemukan"
    pdf_stream = io.BytesIO(file.read())
    data = parse_bca_pdf_logic(pdf_stream)
    excel_file = create_excel_output(data)
    return send_file(excel_file, as_attachment=True, download_name="Mutasi_BCA_Final.xlsx")

if __name__ == '__main__':
    app.run(debug=True)