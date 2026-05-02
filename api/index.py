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

# --- LOGIKA EKSTRAKSI NAMA & PENJELASAN (V5) ---
def extract_nama_refined_v5(block, amount):
    amt_str = "{:.2f}".format(amount)
    found_idx = -1
    for i, line in enumerate(block):
        if amt_str in line.replace(',', ''):
            found_idx = i
            break
            
    if found_idx == -1: return "", ""

    candidates = block[found_idx+1:]
    name_list = []
    desc_list = []
    
    stoppers = ['KCU', 'PERIODE', 'MATA UANG', 'IDR', 'INDONESIA', 'CATATAN', 'JL ', 'BANDUNG']
    garbage = ['TANGGAL', 'KETERANGAN', 'MUTASI', 'SALDO', 'HALAMAN', 'REKENING', 'CBG']
    
    stop_processing = False
    for cand in candidates:
        if stop_processing: break
        c_strip = cand.strip()
        if not c_strip: continue
        c_up = c_strip.upper()

        # 1. Cek Stopper
        for s in stoppers:
            if s in c_up:
                split_part = re.split(s, c_up, flags=re.IGNORECASE)[0].strip()
                if split_part: name_list.append(split_part)
                stop_processing = True
                break
        if stop_processing: break

        # 2. Filter Garbage Header
        if any(g in c_up for g in garbage): continue

        # 3. Logika Pemisahan (Nama vs Penjelasan)
        # Jika mengandung huruf kecil -> Masuk Penjelasan (misal: 'baud')
        if any(c.islower() for c in c_strip):
            desc_list.append(c_strip)
        # Jika diawali '/' -> Masuk Penjelasan (misal: '/9938-KCP RA')
        elif c_strip.startswith('/'):
            desc_list.append(c_strip.replace('/', '').strip())
        # Jika Kapital semua -> Masuk Nama (misal: 'RANI KHOERUN NISA')
        elif c_strip.isupper():
            # Abaikan jika hanya kode teknis singkat seperti 'DR 002'
            if not re.match(r'^[A-Z]{2}\s\d+$', c_strip):
                name_list.append(c_strip)

    # Clean Up Nama
    raw_name = " ".join(name_list)
    patterns = [r'\bTRSF\b', r'\bWS\b', r'\bFTSCY\b', r'\bDB\b', r'\bCR\b', r'\bDR\b', r'\bSWITCHING\b']
    for p in patterns:
        raw_name = re.sub(p, '', raw_name, flags=re.IGNORECASE)
    
    final_name = " ".join(raw_name.split()).strip().upper()
    final_desc = " ".join(desc_list).strip()
    
    return final_name, final_desc

# --- LOGIKA PARSING PDF ---
def parse_bca_pdf_logic(pdf_stream):
    all_lines = []
    with pdfplumber.open(pdf_stream) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                for line in text.split('\n'):
                    l_up = line.upper()
                    if any(x in l_up for x in ["BERSAMBUNG KE HALAMAN", "TGL. CETAK", "REKENING INI"]): continue
                    all_lines.append(line.strip())

    period_line = next((l for l in all_lines if 'PERIODE :' in l.upper()), "")
    period = period_line.split(':', 1)[1].strip() if ":" in period_line else ""
    year = period.split()[-1] if period else str(datetime.now().year)
    
    saldo_awal = 0
    for l in all_lines:
        m = re.search(r'SALDO AWAL\s*:\s*([\d,]+\.\d+)', l.upper())
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
        # Output dua nilai dari fungsi baru
        nama_orang, penjelasan = extract_nama_refined_v5(block, amt)

        if kode in ['27', '28', '29'] or ket == 'PENERIMAAN NEGARA':
            nama_orang = penjelasan = ''

        debet = amt if is_db else 0
        kredit = amt if (is_cr and not is_db) else 0
        total_db += debet
        total_cr += kredit
        
        txs.append({
            'date': date_str, 'nama': nama_orang, 'ket': ket, 
            'kode': kode, 'debet': debet, 'kredit': kredit, 'penjelasan': penjelasan
        })

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
    
    # Tambah kolom PENJELASAN di ujung (Kolom I)
    headers = ['NO', 'TANGGAL', 'NAMA', 'KETERANGAN', 'KODE', 'DEBET', 'KREDIT', 'SALDO', 'PENJELASAN']
    ws.append(headers)
    for c in range(1, 10):
        ws.cell(5, c).font = Font(bold=True, color="FFFFFF")
        ws.cell(5, c).fill = PatternFill("solid", start_color="0056b3")

    ws.append(['', '', '', 'SALDO AWAL', '', '', '', data['saldo_awal'], ''])
    ws.cell(6, 8).number_format = num_fmt
    
    curr_saldo = data['saldo_awal']
    row_idx = 7
    for idx, tx in enumerate(data['txs'], 1):
        curr_saldo = round(curr_saldo + tx['kredit'] - tx['debet'], 2)
        ws.append([
            idx, tx['date'], tx['nama'], tx['ket'], tx['kode'], 
            tx['debet'] or '', tx['kredit'] or '', curr_saldo, tx['penjelasan']
        ])
        row_idx += 1

    ws.append(['', '', '', 'TOTAL MUTASI', '', data['total_db'], data['total_cr'], curr_saldo, ''])
    for c in range(4, 9):
        ws.cell(row_idx, c).font = Font(bold=True)
        if c >= 6: ws.cell(row_idx, c).number_format = num_fmt

    widths = [5, 12, 25, 25, 8, 15, 15, 18, 30]
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
            <h2 style="color:#0056b3;">BCA Converter v5</h2>
            <p>Fitur: Kolom Penjelasan & Pemisah Nama Kapital</p>
            <form action="/convert" method="post" enctype="multipart/form-data">
                <input type="file" name="file" accept=".pdf" required style="margin-bottom:20px;"><br>
                <button type="submit" style="padding:12px 25px; background:#0056b3; color:white; border:none; border-radius:5px; cursor:pointer; font-weight:bold;">Download Excel</button>
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
    return send_file(excel_file, as_attachment=True, download_name="Mutasi_BCA_Penjelasan.xlsx")

if __name__ == '__main__':
    app.run(debug=True)