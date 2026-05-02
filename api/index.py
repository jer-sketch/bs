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

# --- LOGIKA EKSTRAKSI NAMA & PENJELASAN ---
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

        for s in stoppers:
            if s in c_up:
                split_part = re.split(s, c_up, flags=re.IGNORECASE)[0].strip()
                if split_part: name_list.append(split_part)
                stop_processing = True
                break
        if stop_processing: break
        if any(g in c_up for g in garbage): continue

        if any(c.islower() for c in c_strip):
            desc_list.append(c_strip)
        elif c_strip.startswith('/'):
            desc_list.append(c_strip.replace('/', '').strip())
        elif c_strip.isupper():
            if not re.match(r'^[A-Z]{2}\s\d+$', c_strip):
                name_list.append(c_strip)

    raw_name = " ".join(name_list)
    patterns = [r'\bTRSF\b', r'\bWS\b', r'\bFTSCY\b', r'\bDB\b', r'\bCR\b', r'\bDR\b', r'\bSWITCHING\b']
    for p in patterns: raw_name = re.sub(p, '', raw_name, flags=re.IGNORECASE)
    
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
        elif current_block: current_block.append(l)
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
        ws.append([idx, tx['date'], tx['nama'], tx['ket'], tx['kode'], 
                   tx['debet'] or '', tx['kredit'] or '', curr_saldo, tx['penjelasan']])
        row_idx += 1
    ws.append(['', '', '', 'TOTAL MUTASI', '', data['total_db'], data['total_cr'], curr_saldo, ''])
    for c in range(4, 9):
        ws.cell(row_idx, c).font = Font(bold=True)
        if c >= 6: ws.cell(row_idx, c).number_format = num_fmt
    widths = [5, 12, 25, 25, 8, 15, 15, 18, 30]
    for i, w in enumerate(widths, 1): ws.column_dimensions[get_column_letter(i)].width = w
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

# --- FLASK APP + MODERN UI ---
@app.route('/')
def index():
    return render_template_string("""
    <!DOCTYPE html>
    <html lang="id">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>BCA PDF to Excel Converter</title>
        <script src="https://cdn.tailwindcss.com"></script>
        <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap" rel="stylesheet">
        <style>
            body { font-family: 'Inter', sans-serif; }
            .glass { background: rgba(255, 255, 255, 0.95); backdrop-filter: blur(10px); }
        </style>
    </head>
    <body class="bg-slate-50 min-h-screen flex items-center justify-center p-4">
        <div class="max-w-md w-full glass rounded-3xl shadow-2xl border border-slate-200 overflow-hidden">
            <div class="bg-blue-600 p-8 text-white text-center">
                <div class="inline-flex items-center justify-center w-16 h-16 bg-white/20 rounded-2xl mb-4">
                    <svg class="w-8 h-8" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="9 17v-2m3 2v-4m3 4v-6m2 10H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"></path>
                    </svg>
                </div>
                <h1 class="text-2xl md:text-3xl font-bold tracking-tight">BCA Converter</h1>
                <p class="text-blue-100 mt-2 text-sm md:text-base">Ekstrak mutasi PDF ke Excel secara otomatis</p>
            </div>
            
            <div class="p-8">
                <form action="/convert" method="post" enctype="multipart/form-data" id="uploadForm">
                    <div class="space-y-6">
                        <div class="relative group">
                            <label class="block text-sm font-semibold text-slate-700 mb-2">Pilih File PDF Mutasi</label>
                            <label for="file-upload" class="flex flex-col items-center justify-center w-full h-32 border-2 border-dashed border-slate-300 rounded-xl cursor-pointer bg-slate-50 hover:bg-slate-100 hover:border-blue-400 transition-all">
                                <div class="flex flex-col items-center justify-center pt-5 pb-6">
                                    <svg class="w-8 h-8 text-slate-400 mb-2" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"></path></svg>
                                    <p id="file-name" class="text-xs md:text-sm text-slate-500 text-center px-4">Klik untuk pilih file</p>
                                </div>
                                <input id="file-upload" name="file" type="file" class="hidden" accept=".pdf" required onchange="updateFileName(this)"/>
                            </label>
                        </div>

                        <button type="submit" class="w-full bg-blue-600 hover:bg-blue-700 text-white font-bold py-4 rounded-xl shadow-lg shadow-blue-200 transition-all active:scale-[0.98] flex items-center justify-center text-base md:text-lg">
                            <span>Konversi Sekarang</span>
                            <svg class="w-5 h-5 ml-2" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 7l5 5m0 0l-5 5m5-5H6"></path></svg>
                        </button>
                    </div>
                </form>
                
                <div class="mt-8 pt-6 border-t border-slate-100">
                    <div class="flex items-center text-xs md:text-sm text-slate-400 justify-center">
                        <svg class="w-4 h-4 mr-1" fill="currentColor" viewBox="0 0 20 20"><path fill-rule="evenodd" d="M5 9V7a5 5 0 0110 0v2a2 2 0 012 2v5a2 2 0 01-2 2H5a2 2 0 01-2-2v-5a2 2 0 012-2zm8-2v2H7V7a3 3 0 016 0z" clip-rule="evenodd"></path></svg>
                        Keamanan Terjamin • Pemrosesan Lokal
                    </div>
                </div>
            </div>
        </div>

        <script>
            function updateFileName(input) {
                const fileName = input.files[0]?.name || "Klik untuk pilih file";
                document.getElementById('file-name').textContent = fileName;
                document.getElementById('file-name').classList.add('text-blue-600', 'font-semibold');
            }
        </script>
    </body>
    </html>
    """)

@app.route('/convert', methods=['POST'])
def convert():
    f = request.files['file']
    if not f: return "File tidak ditemukan"
    data = parse_bca_pdf_logic(io.BytesIO(f.read()))
    excel_file = create_excel_output(data)
    return send_file(excel_file, as_attachment=True, download_name="Mutasi_BCA_V5.xlsx")

if __name__ == '__main__':
    app.run(debug=True)
