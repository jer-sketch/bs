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

def is_strict_uppercase(text):
    """Hanya mengembalikan True jika teks benar-benar huruf besar (mengabaikan angka/simbol)."""
    # Mencari apakah ada huruf kecil di dalam teks
    if any(c.islower() for c in text):
        return False
    # Harus mengandung setidaknya satu huruf (bukan cuma angka/simbol)
    if not any(c.isalpha() for c in text):
        return False
    return True

def clean_name_logic(text):
    """Membersihkan kode-kode transaksi agar tersisa Nama saja."""
    # List kata yang sering muncul tapi bukan bagian dari nama asli
    patterns = [
        r'\bTRSF\b', r'\bE-BANKING\b', r'\bM-BCA\b', r'\bDB\b', r'\bCR\b', 
        r'\bBIF\b', r'\bSWITCHING\b', r'\bWS\b', r'\bBCA\b', r'\bKCU\b',
        r'\bPT BANK CENTRAL ASIA\b', r'\bINDONESIA\b'
    ]
    res = text.upper()
    for p in patterns:
        res = re.sub(p, '', res, flags=re.IGNORECASE)
    
    res = re.sub(r'\d+', '', res) # Hapus angka
    res = re.sub(r'[./\-_:+]', '', res) # Hapus simbol
    return re.sub(r'\s+', ' ', res).strip()

def parse_bca_pdf_logic(pdf_stream):
    all_lines = []
    with pdfplumber.open(pdf_stream) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                # Bersihkan baris footer halaman yang jelas-jelas sampah sebelum diproses
                for line in text.split('\n'):
                    l_up = line.upper()
                    # Footer statis BCA yang sering mengganggu
                    if "BERSAMBUNG KE HALAMAN" in l_up: continue
                    if "TGL. CETAK" in l_up: continue
                    if "REKENING INI" in l_up: continue
                    all_lines.append(line.strip())

    # Ambil periode untuk tahun
    period = next((l.split(':', 1)[1].strip() for l in all_lines if 'PERIODE :' in l), "")
    year = period.split()[-1] if period else str(datetime.now().year)
    
    # Ambil saldo awal
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
        
        # Koreksi Pajak/Bunga
        if 'PAJAK' in full_text: is_db, is_cr = True, False
        if 'BUNGA' in full_text and 'PAJAK' not in full_text: is_db, is_cr = False, True

        ket, kode = classify(full_text, is_cr and not is_db)
        
        # LOGIKA NAMA: 
        # 1. Cari dari baris terbawah di dalam blok ini.
        # 2. Harus Strict Uppercase (Tanpa huruf kecil).
        # 3. Jika "PT BANK CENTRAL ASIA" adalah satu-satunya teks, kita ambil baris di atasnya.
        nama_orang = ""
        for i in range(len(block)-1, 0, -1):
            candidate = block[i].strip()
            # Cek jika baris berisi huruf kapital semua & bukan footer umum
            if is_strict_uppercase(candidate):
                # Jika baris ini berisi "PT BANK CENTRAL ASIA" atau "INDONESIA" 
                # seringkali ini adalah keterangan bank, bukan nama pengirim.
                # Kita coba ambil teksnya, tapi jika ada baris lain yang lebih spesifik, itu lebih baik.
                temp_name = clean_name_logic(candidate)
                if temp_name and temp_name not in ["PT BANK CENTRAL ASIA", "INDONESIA"]:
                    nama_orang = temp_name
                    break
                elif temp_name: # Jika hanya ada itu, simpan dulu sebagai cadangan
                    nama_orang = temp_name

        if kode in ['27', '28', '29'] or ket == 'PENERIMAAN NEGARA':
            nama_orang = ''

        debet = amt if is_db else 0
        kredit = amt if (is_cr and not is_db) else 0
        total_db += debet
        total_cr += kredit
        
        txs.append({'date':date_str, 'nama':nama_orang, 'ket':ket, 'kode':kode, 'debet':debet, 'kredit':kredit})

    return {"txs":txs, "period":period, "saldo_awal":saldo_awal, "total_db":total_db, "total_cr":total_cr}

# --- Fungsi create_excel_output tetap sama seperti sebelumnya (termasuk Baris Total) ---
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

# --- Flask Routes ---
@app.route('/')
def index():
    return render_template_string("""
    <body style="font-family:sans-serif; text-align:center; padding:50px; background:#f4f7f6;">
        <div style="display:inline-block; background:white; padding:40px; border-radius:15px; shadow:0 4px 15px rgba(0,0,0,0.1)">
            <h2 style="color:#0056b3;">BCA PDF Matcher v3</h2>
            <p>Fix: Nama Kapital & Filter Footer Indonesia</p>
            <form action="/convert" method="post" enctype="multipart/form-data">
                <input type="file" name="file" accept=".pdf" required><br><br>
                <button type="submit" style="padding:12px 25px; background:#0056b3; color:white; border:none; border-radius:5px; cursor:pointer;">Proses PDF</button>
            </form>
        </div>
    </body>
    """)

@app.route('/convert', methods=['POST'])
def convert():
    f = request.files['file']
    if not f: return "File Error"
    data = parse_bca_pdf_logic(io.BytesIO(f.read()))
    return send_file(create_excel_output(data), as_attachment=True, download_name="Mutasi_BCA_Terupdate.xlsx")

if __name__ == '__main__':
    app.run(debug=True)