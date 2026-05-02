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
    # PENTING: Hanya mengacu pada kata "PENERIMAAN NEGARA"
    if 'PENERIMAAN NEGARA' in u: return 'PENERIMAAN NEGARA', ''
    
    if 'BIAYA ADM' in u: return 'BIAYA ADM BANK', '27'
    if 'PAJAK BUNGA' in u or 'PAJAK JASA' in u: return 'PAJAK JASA GIRO', ''
    if 'BUNGA' in u and not ('PAJAK' in u): return 'BUNGA', ''
    if 'TRANSFER' in u and 'BIAYA' in u: return 'BIAYA TRANSFER', '27'
    if 'TELKOM' in u or 'TELEPON' in u: return 'BIAYA TELEPON', '27'
    if 'LISTRIK' in u or 'PLN' in u: return 'BIAYA LISTRIK', '27'
    if 'GAJI' in u: return 'BIAYA GAJI PEGAWAI', ''
    if 'SETORAN TUNAI' in u: return 'SETORAN TUNAI', '1'
    
    if not is_credit: return 'PELUNASAN HUTANG DAGANG', ''
    return 'PENERIMAAN PENJUALAN', '3' 

def extract_nama_after_nominal(sub_lines, amount):
    """
    Mencari baris yang mengandung angka nominal, 
    lalu mengambil baris-baris setelahnya sebagai nama.
    """
    amt_str = "{:,.2f}".format(amount).replace(',', '') # Format: 41800.00
    amt_comma = "{:,.2f}".format(amount) # Format: 41,800.00
    
    found_idx = -1
    for i, line in enumerate(sub_lines):
        # Cek apakah baris ini mengandung nominal transaksi
        clean_line = line.replace(',', '')
        if amt_str in clean_line or amt_comma in line:
            found_idx = i
            break
    
    if found_idx != -1 and found_idx < len(sub_lines) - 1:
        # Ambil semua teks setelah baris nominal
        nama_parts = sub_lines[found_idx+1:]
        # Bersihkan dari kata-kata sampah bank jika masih terbawa
        nama_res = " ".join(nama_parts).strip()
        nama_res = re.sub(r'\b(TRSF|E-BANKING|DB|CR|BIF|SWITCHING)\b', '', nama_res, flags=re.IGNORECASE)
        return nama_res.strip(' -:').upper()
    
    return ""

def parse_bca_pdf_logic(pdf_stream):
    all_lines = []
    with pdfplumber.open(pdf_stream) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                all_lines.extend(text.split('\n'))

    period = ''
    saldo_awal = saldo_akhir = mut_cr = mut_db = 0
    
    # Ambil Metadata
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

    year = period.strip().split()[-1] if period else str(datetime.now().year)

    # Identifikasi baris transaksi
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
        # 1. Skip jika ini baris Saldo Awal di dalam mutasi
        full_block_text = " ".join(block).upper()
        if 'SALDO AWAL' in full_block_text:
            continue

        first_line = block[0]
        m = DATE_RE.match(first_line)
        day, mon = m.group(1), m.group(2)
        date_str = f"{day}/{mon}/{year}"

        # Cari nominal di blok ini
        # Kita cari angka dengan format desimal .00
        amounts_found = re.findall(r'([\d,]+\.\d{2})', " ".join(block))
        if not amounts_found:
            continue
        
        amt = float(amounts_found[0].replace(',', ''))
        
        # Tentukan Debet/Kredit
        is_cr = (' CR' in full_block_text or 'SETORAN TUNAI' in full_block_text or 'KR OTOMATIS' in full_block_text)
        is_db = (' DB' in full_block_text or 'BYR VIA' in full_block_text or 'BIAYA ADM' in full_block_text)
        
        # Klasifikasi
        ket, kode = classify(full_block_text, is_cr and not is_db)
        
        # Ekstraksi Nama: Cari baris setelah baris yang berisi nominal
        nama_orang = extract_nama_after_nominal(block, amt)
        
        # Jika Penerimaan Negara, kolom nama dikosongkan agar rapi
        if ket == 'PENERIMAAN NEGARA':
            nama_orang = ''

        debet = amt if is_db else 0
        kredit = amt if (is_cr and not is_db) else 0
        if debet == 0 and kredit == 0: kredit = amt # Default jika tidak terdeteksi DB/CR

        txs.append({
            'date': date_str,
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
    
    num_fmt = '#,##0.00'
    
    ws.append(['BCA 346-8383111'])
    ws.append(['CV. MITRA JAYA ANUGERAH'])
    ws.append([f"PERIODE: {data['period']}"])
    ws.append([]) 
    
    headers = ['NO', 'TANGGAL', 'NAMA', 'KETERANGAN', 'KODE', 'DEBET', 'KREDIT', 'SALDO']
    ws.append(headers)
    
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

    ws.append([])
    ws.append(['', '', '', 'TOTAL MUTASI', '', data['mut_db'], data['mut_cr'], data['saldo_akhir']])
    
    # Styling
    for row in ws.iter_rows(min_row=5):
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
    <body style="font-family:sans-serif; text-align:center; padding-top:50px; background:#f0f2f5;">
        <div style="display:inline-block; background:white; padding:30px; border-radius:10px; shadow:0 2px 10px rgba(0,0,0,0.1)">
            <h2>BCA Mutasi Fix (Nama & Penerimaan Negara)</h2>
            <form action="/convert" method="post" enctype="multipart/form-data">
                <input type="file" name="file" accept=".pdf" required><br><br>
                <button type="submit" style="padding:10px 20px; background:#0056b3; color:white; border:none; border-radius:5px; cursor:pointer;">Proses PDF</button>
            </form>
        </div>
    </body>
    """)

@app.route('/convert', methods=['POST'])
def convert():
    file = request.files['file']
    if not file: return "File tidak ada"
    pdf_stream = io.BytesIO(file.read())
    data = parse_bca_pdf_logic(pdf_stream)
    excel_file = create_excel_output(data)
    return send_file(excel_file, as_attachment=True, download_name="Mutasi_BCA_Updated.xlsx")

if __name__ == '__main__':
    app.run(debug=True)