from flask import Flask, request, send_file, render_template_string
import io
import pdfplumber
import re
import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

app = Flask(__name__)

# --- LOGIKA PARSING ANDA (Disederhanakan untuk Web) ---
MONTH_MAP = {'JANUARI':1,'FEBRUARI':2,'MARET':3,'APRIL':4,'MEI':5,'JUNI':6,
             'JULI':7,'AGUSTUS':8,'SEPTEMBER':9,'OKTOBER':10,'NOVEMBER':11,'DESEMBER':12}

def classify(line, is_credit):
    u = line.upper()
    if 'BIAYA ADM' in u: return 'BIAYA ADM BANK','27'
    if 'PAJAK BUNGA' in u or 'PAJAK JASA GIRO' in u: return 'PAJAK JASA GIRO',''
    if u.strip().startswith('BUNGA'): return 'BUNGA',''
    if 'BIAYA TRANSFER' in u: return 'BIAYA TRANSFER','27'
    if 'TELKOM' in u or 'TELEPON' in u: return 'BIAYA TELEPON','27'
    if 'LISTRIK' in u or 'PLN' in u: return 'BIAYA LISTRIK','27'
    if 'GAJI' in u: return 'BIAYA GAJI PEGAWAI',''
    if 'SETORAN TUNAI' in u: return 'SETORAN TUNAI','1'
    if not is_credit: return 'PELUNASAN HUTANG DAGANG',''
    return 'PENERIMAAN PENJUALAN','5'

def process_pdf_to_excel(file_stream):
    with pdfplumber.open(file_stream) as pdf:
        all_lines = []
        for page in pdf.pages:
            text = page.extract_text()
            if text: all_lines.extend(text.split('\n'))

    # Ekstraksi Data Header
    period = ""
    saldo_awal = saldo_akhir = mut_cr = mut_db = 0
    for l in all_lines:
        if 'PERIODE :' in l: period = l.split(':',1)[1].strip()
        m = re.search(r'SALDO AWAL\s*:\s*([\d,]+\.\d+)', l)
        if m: saldo_awal = float(m.group(1).replace(',',''))
        m = re.search(r'MUTASI CR\s*:\s*([\d,]+\.\d+)', l)
        if m: mut_cr = float(m.group(1).replace(',',''))
        m = re.search(r'MUTASI DB\s*:\s*([\d,]+\.\d+)', l)
        if m: mut_db = float(m.group(1).replace(',',''))
        m = re.search(r'SALDO AKHIR\s*:\s*([\d,]+\.\d+)', l)
        if m: saldo_akhir = float(m.group(1).replace(',',''))

    # Parsing Transaksi
    DATE_RE = re.compile(r'^(\d{2})/(\d{2})\s+(.+)')
    txs = []
    year = period.split()[-1] if period else "2025"
    
    for l in all_lines:
        l = l.strip()
        dm = DATE_RE.match(l)
        if not dm: continue
        day, mon, rest = dm.group(1), dm.group(2), dm.group(3).strip()
        if 'SALDO AWAL' in rest or 'Bersambung' in rest: continue
        
        all_amounts = re.findall(r'([\d,]+\.\d{2})', rest)
        if not all_amounts: continue
        amt = float(all_amounts[0].replace(',',''))
        
        u = rest.upper()
        is_cr = any(x in u for x in [' CR ', 'SETORAN TUNAI', 'KR OTOMATIS', 'SWITCHING CR'])
        is_db = any(x in u for x in [' DB ', 'BYR VIA'])
        
        if any(x in u for x in ['BIAYA ADM', 'PAJAK BUNGA']): is_db, is_cr = True, False
        if u.startswith('BUNGA'): is_db, is_cr = False, True

        ket, kode = classify(rest, is_cr and not is_db)
        txs.append({'date': f'{year}-{mon}-{day}', 'ket': ket, 'kode': kode, 
                    'debet': amt if is_db else 0, 'kredit': amt if (is_cr and not is_db) else (amt if not is_db else 0)})

    # Membuat Excel di Memori
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'bca'
    ws.append(['BCA Statement Converter'])
    ws.append(['|','NO','TANGGAL','KETERANGAN','KODE','DEBET','KREDIT','SALDO','|'])
    
    saldo = saldo_awal
    for i, t in enumerate(txs, 1):
        saldo = round(saldo + t['kredit'] - t['debet'], 2)
        ws.append(['|', i, t['date'], t['ket'], t['kode'], t['debet'] or '', t['kredit'] or '', saldo, '|'])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --- ROUTE FLASK ---
@app.route('/')
def home():
    return render_template_string('''
        <h1>BCA PDF to Excel Converter</h1>
        <form action="/convert" method="post" enctype="multipart/form-data">
            <input type="file" name="file" accept=".pdf">
            <button type="submit">Convert Sekarang</button>
        </form>
    ''')

@app.route('/convert', methods=['POST'])
def convert():
    file = request.files['file']
    if not file: return "Pilih file dulu!"
    
    excel_file = process_pdf_to_excel(file)
    return send_file(excel_file, 
                     download_name="Laporan_BCA.xlsx", 
                     as_attachment=True)