from flask import Flask, request, send_file, render_template_string
import pdfplumber
import re
import io
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from datetime import datetime

app = Flask(__name__)

# --- LOGIKA KLASIFIKASI KODE ---
def classify(full_text, is_credit):
    u = full_text.upper()
    if 'PENERIMAAN NEGARA' in u: return 'PENERIMAAN NEGARA', ''
    if 'PAJAK BUNGA' in u or 'PAJAK JASA' in u: return 'PAJAK JASA GIRO', '29'
    if 'BUNGA' in u: return 'BUNGA', '28'
    if 'BIAYA ADM' in u: return 'BIAYA ADM BANK', '27'
    if 'TRANSFER' in u and 'BIAYA' in u: return 'BIAYA TRANSFER', '27'
    if 'TELKOM' in u or 'TELEPON' in u: return 'BIAYA TELEPON', '27'
    if 'LISTRIK' in u or 'PLN' in u: return 'BIAYA LISTRIK', '27'
    if 'GAJI' in u: return 'BIAYA GAJI PEGAWAI', ''
    if 'SETORAN TUNAI' in u: return 'SETORAN TUNAI', '1'
    
    if not is_credit: return 'PELUNASAN HUTANG DAGANG', ''
    return 'PENERIMAAN PENJUALAN', '3' 

# --- LOGIKA EKSTRAKSI NAMA (REFINED) ---
def is_strictly_capital(text):
    """Memastikan teks hanya berisi huruf kapital, angka, atau simbol (Tanpa Huruf Kecil)."""
    if any(c.islower() for c in text):
        return False
    if not any(c.isupper() for c in text):
        return False
    return True

def extract_nama_refined(block, amount):
    """
    Mengambil semua baris KAPITAL setelah nominal ditemukan.
    Mengabaikan baris yang memiliki huruf kecil (pembayaran sage, dll).
    """
    amt_str = "{:.2f}".format(amount)
    
    found_idx = -1
    for i, line in enumerate(block):
        if amt_str in line.replace(',', ''):
            found_idx = i
            break
            
    if found_idx == -1:
        return ""

    candidates = block[found_idx+1:]
    name_parts = []
    
    garbage_keywords = ['TANGGAL', 'KETERANGAN', 'MUTASI', 'SALDO', 'HALAMAN', 'REKENING', 'CBG']
    
    for cand in candidates:
        cand_clean = cand.strip()
        
        # 1. Harus Kapital Semua (Abaikan 'pembayaran sage')
        if not is_strictly_capital(cand_clean):
            continue
            
        # 2. Abaikan Header/Footer yang Kapital
        if any(g in cand_clean.upper() for g in garbage_keywords):
            continue
            
        # 3. Jika lolos, masukkan ke list (mendukung nama 2 baris)
        name_parts.append(cand_clean)

    # Gabungkan baris-baris kapital yang ditemukan
    full_raw_name = " ".join(name_parts)
    
    # Pembersihan final kode perbankan
    patterns = [
        r'\bTRSF\b', r'\bWS\b', r'\bFTSCY\b', r'\bDB\b', r'\bCR\b', r'\bDR\b',
        r'\bE-BANKING\b', r'\bM-BCA\b', r'\bBIF\b', r'\bSWITCHING\b'
    ]
    for p in patterns:
        full_raw_name = re.sub(p, '', full_raw_name, flags=re.IGNORECASE)
    
    # Hapus angka yang berdiri sendiri (kode referensi)
    full_raw_name = re.sub(r'\b\d+\b', '', full_raw_name)
    
    # Bersihkan simbol sisa
    full_raw_name = re.sub(r'[.:/]', ' ', full_raw_name)
    
    return " ".join(full_raw_name.split()).strip().upper()

# --- LOGIKA PARSING PDF ---
def parse_bca_pdf_logic(pdf_stream):
    all_lines = []
    with pdfplumber.open(pdf_stream) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                for line in text.split('\n'):
                    l_up = line.upper()
                    # Filter awal sampah halaman
                    if "BERSAMBUNG KE HALAMAN" in l_up: continue
                    if "TGL. CETAK" in l_up: continue
                    if "REKENING INI" in l_up: continue
                    all_lines.append(line.strip())

    period = next((l.split(':', 1)[1].strip() for l in all_lines if 'PERIODE :' in l), "")
    year = period.split()[-1] if period else str(datetime.now().year)
    
    saldo_awal = 0
    for l in all_lines:
        m = re.search(r'SALDO AWAL\s*:\s*([\d,]+\.\d+)', l)
        if m: 
            saldo_awal = float(m.group(1).replace(',',''))
            break

    DATE_RE = re.compile(r'^(\d{2})/(\d{2})\s+')
    tx_blocks = []
    current_block = []

    for l in all_lines:
        if DATE_RE.match(l):
            if current_block: tx_blocks.append(current_block)
            current_block = [l]
        elif current_block:
            current_block.append(l)
    if current_block: tx_blocks.append(current_block)

    txs = []
    total_db = total_cr = 0
    
    for block in tx_blocks:
        full_text = " ".join(block).upper()
        if 'SALDO AWAL' in full_text and len(block) < 3: continue

        m = DATE_RE.match(block[0])
        date_str = f"{m.group(1)}/{m.group(2)}/{year}"

        amounts = re.findall(r'([\d,]+\.\d{2})', full_text)
        if not amounts: continue
        amt = float(amounts[0].replace(',', ''))
        
        is_cr = any(x in full_text for x in [' CR', 'SETORAN TUNAI', 'KR OTOMATIS', 'BUNGA'])
        is_db = any(x in full_text for x in [' DB', 'BYR VIA', 'BIAYA ADM', 'PAJAK'])
        
        if 'PAJAK' in full_text: is_db, is_cr = True, False
        if 'BUNGA' in full_text and 'PAJAK' not in full_text: is_db, is_cr = False, True

        ket, kode = classify(full_text, is_cr and not is_db)
        
        # GUNAKAN LOGIKA REFINED DI SINI
        nama_orang = extract_nama_refined(block, amt)

        if kode in ['27', '28', '29'] or ket == 'PENERIMAAN NEGARA':
            nama_orang = ''

        debet = amt if is_db else 0
        kredit = amt if (is_cr and not is_db) else 0
        total_db += debet
        total_cr += kredit
        
        txs.append({'date':date_str, 'nama':nama_orang, 'ket':ket, 'kode':kode, 'debet':debet, 'kredit':kredit})

    return {"txs":txs, "period":period, "saldo_awal":saldo_awal, "total_db":total_db, "total_cr":total_cr}

# --- LOGIKA EXCEL ---
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
    for c in range(1, 9):
        ws.cell(5, c).font = Font(bold=True, color="FFFFFF")
        ws.cell(5, c).fill = PatternFill("solid", start_color="0056b3")

    ws.append(['', '', '', 'SALDO AWAL', '', '', '', data['saldo_awal']])
    ws.cell(6, 8).number_format = num_fmt
    
    curr_saldo = data['saldo_awal']
    row_idx = 7
    for idx, tx in enumerate(data['txs'], 1):
        curr_saldo = round(curr_saldo + tx['kredit'] - tx['debet'], 2)
        ws.append([idx, tx['date'], tx['nama'], tx['ket'], tx['kode'], 
                   tx['debet'] or '', tx['kredit'] or '', curr_saldo])
        row_idx += 1

    # BARIS TOTAL MUTASI
    ws.append(['', '', '', 'TOTAL MUTASI', '', data['total_db'], data['total_cr'], curr_saldo])
    for c in range(4, 9):
        ws.cell(row_idx, c).font = Font(bold=True)
        if c >= 6: ws.cell(row_idx, c).number_format = num_fmt

    widths = [5, 12, 30, 30, 8, 16, 16, 18]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

# --- FLASK APP ---
@app.route('/')
def index():
    return render_template_string("""
    <body style="font-family:sans-serif; text-align:center; padding-top:50px; background:#f4f7f6;">
        <div style="display:inline-block; background:white; padding:40px; border-radius:15px; box-shadow:0 4px 15px rgba(0,0,0,0.1)">
            <h2 style="color:#0056b3;">BCA Converter (Nama Multi-Line Kapital)</h2>
            <form action="/convert" method="post" enctype="multipart/form-data">
                <input type="file" name="file" accept=".pdf" required style="margin-bottom:20px;"><br>
                <button type="submit" style="padding:12px 25px; background:#0056b3; color:white; border:none; border-radius:5px; cursor:pointer; font-weight:bold;">Proses & Download Excel</button>
            </form>
        </div>
    </body>
    """)

@app.route('/convert', methods=['POST'])
def convert():
    f = request.files['file']
    if not f: return "File tidak ditemukan"
    data = parse_bca_pdf_logic(io.BytesIO(f.read()))
    excel_file = create_excel_output(data)
    return send_file(excel_file, as_attachment=True, download_name="Mutasi_BCA_Final.xlsx")

if __name__ == '__main__':
    app.run(debug=True)