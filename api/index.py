"""
Bank Statement PDF → Excel Converter
Mendukung: BCA (V5 Logic), BRI, Mandiri
Versi: 2.3 (Fix Compatibility & Detection)
"""

import re
import io
import traceback
from datetime import datetime

import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from flask import Flask, request, send_file, render_template_string, make_response

app = Flask(__name__)

# ─────────────────────────────────────────────
# KLASIFIKASI KETERANGAN & KODE
# ─────────────────────────────────────────────

KODE_MAP = {
    'PENERIMAAN PENJUALAN':     '3',
    'SETORAN TUNAI':            '1',
    'PELUNASAN PIUTANG DAGANG': '1',
    'PELUNASAN HUTANG DAGANG':  '',
    'BIAYA ADM BANK':           '27',
    'BIAYA TRANSFER':           '27',
    'BIAYA TELEPON':            '27',
    'BIAYA LISTRIK':            '27',
    'BIAYA GAJI PEGAWAI':       '',
    'BUNGA':                    '28',
    'PAJAK JASA GIRO':          '29',
    'PENERIMAAN NEGARA':        '',
    'KOREKSI BANK':             '',
}

def classify(text: str, is_credit: bool) -> tuple[str, str]:
    u = text.upper()
    if 'PENERIMAAN NEGARA' in u:                        ket = 'PENERIMAAN NEGARA'
    elif 'PAJAK' in u:                                  ket = 'PAJAK JASA GIRO'
    elif 'BUNGA' in u or 'INTEREST' in u:               ket = 'BUNGA'
    elif 'BIAYA ADM' in u:                              ket = 'BIAYA ADM BANK'
    elif 'TRANSFER' in u and 'BIAYA' in u:              ket = 'BIAYA TRANSFER'
    elif 'TELKOM' in u or 'TELEPON' in u:               ket = 'BIAYA TELEPON'
    elif 'LISTRIK' in u or 'PLN' in u:                  ket = 'BIAYA LISTRIK'
    elif 'GAJI' in u or 'SALARY' in u:                  ket = 'BIAYA GAJI PEGAWAI'
    elif 'SETORAN TUNAI' in u:                          ket = 'SETORAN TUNAI'
    elif 'KOREKSI' in u:                                ket = 'KOREKSI BANK'
    elif not is_credit:                                 ket = 'PELUNASAN HUTANG DAGANG'
    else:                                               ket = 'PENERIMAAN PENJUALAN'
    return ket, KODE_MAP.get(ket, '')


# ─────────────────────────────────────────────
# DETEKTOR JENIS BANK (DIPERKUAT)
# ─────────────────────────────────────────────

def detect_bank(all_lines: list[str]) -> str:
    # Ambil 50 baris pertama untuk identifikasi lebih akurat
    text = ' '.join(all_lines[:50]).upper()
    
    # Keyword BCA yang lebih luas
    if any(x in text for x in ['BCA', 'REKENING GIRO', 'E-BANKING', 'BI-FAST', 'TANGGAL KETERANGAN MUTASI']):
        return 'BCA'
    # Keyword BRI
    if any(x in text for x in ['BRI', 'BRIMTXDT', 'BRITAMA', 'LAPORAN TRANSAKSI FINANSIAL']):
        return 'BRI'
    # Keyword Mandiri
    if any(x in text for x in ['MANDIRI', 'KOPRA', 'ACCOUNT STATEMENT', 'MCM INHOUSETRF']):
        return 'MANDIRI'
        
    return 'UNKNOWN'


# ─────────────────────────────────────────────
# PARSER BCA (LOGIK V5 ASLI)
# ─────────────────────────────────────────────

def extract_nama_refined_v5(block: list[str], amount: float) -> tuple[str, str]:
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
    
    stoppers = ['KCU', 'PERIODE', 'MATA UANG', 'IDR', 'INDONESIA', 'CATATAN', 'JL ', 'BANDUNG', 'SALDO AWAL']
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

        if any(c.islower() for c in c_strip): desc_list.append(c_strip)
        elif c_strip.startswith('/'): desc_list.append(c_strip.replace('/', '').strip())
        elif c_strip.isupper():
            if not re.match(r'^[A-Z]{2}\s\d+$', c_strip):
                name_list.append(c_strip)

    raw_name = " ".join(name_list)
    patterns = [r'\bTRSF\b', r'\bWS\b', r'\bFTSCY\b', r'\bDB\b', r'\bCR\b', r'\bDR\b', r'\bSWITCHING\b']
    for p in patterns:
        raw_name = re.sub(p, '', raw_name, flags=re.IGNORECASE)
    
    final_name = " ".join(raw_name.split()).strip().upper()
    final_desc = " ".join(desc_list).strip()
    return final_name, final_desc

def parse_bca(all_lines: list[str]) -> dict:
    period = ''
    saldo_awal = saldo_akhir = mut_cr = mut_db = 0
    no_rek = ''

    cleaned_lines = []
    for l in all_lines:
        l_up = l.upper()
        if 'PERIODE :' in l_up and not period: period = l.split(':', 1)[1].strip()
        if 'NO. REKENING' in l_up and not no_rek:
            m = re.search(r':\s*([\d\s\-]+)', l)
            if m: no_rek = m.group(1).strip()
            
        m = re.search(r'SALDO AWAL\s*:\s*([\d,]+\.\d+)', l_up)
        if m: saldo_awal = float(m.group(1).replace(',', ''))
        m = re.search(r'MUTASI CR\s*:\s*([\d,]+\.\d+)', l_up)
        if m: mut_cr = float(m.group(1).replace(',', ''))
        m = re.search(r'MUTASI DB\s*:\s*([\d,]+\.\d+)', l_up)
        if m: mut_db = float(m.group(1).replace(',', ''))
        m = re.search(r'SALDO AKHIR\s*:\s*([\d,]+\.\d+)', l_up)
        if m: saldo_akhir = float(m.group(1).replace(',', ''))
        
        if any(x in l_up for x in ["BERSAMBUNG KE HALAMAN", "TGL. CETAK"]): continue
        cleaned_lines.append(l.strip())

    year = period.split()[-1] if period else str(datetime.now().year)
    DATE_RE = re.compile(r'^(\d{2})/(\d{2})\s+')
    
    tx_blocks = []
    current_block = []
    for l in cleaned_lines:
        if DATE_RE.match(l):
            if current_block: tx_blocks.append(current_block)
            current_block = [l]
        elif current_block: current_block.append(l)
    if current_block: tx_blocks.append(current_block)

    txs = []
    for block in tx_blocks:
        full_text = " ".join(block).upper()
        if 'SALDO AWAL' in full_text and len(block) < 3: continue
        
        m = DATE_RE.match(block[0])
        date_str = f"{m.group(1)}/{m.group(2)}/{year[-2:]}"
        amounts = re.findall(r'([\d,]+\.\d{2})', full_text)
        if not amounts: continue
        amt = float(amounts[0].replace(',', ''))
        
        is_cr = any(x in full_text for x in [' CR', 'SETORAN TUNAI', 'KR OTOMATIS', 'BUNGA'])
        is_db = any(x in full_text for x in [' DB', 'BYR VIA', 'BIAYA ADM', 'PAJAK'])
        if 'PAJAK' in full_text: is_db, is_cr = True, False
        if 'BUNGA' in full_text and 'PAJAK' not in full_text: is_db, is_cr = False, True

        ket, kode = classify(full_text, is_cr and not is_db)
        nama_orang, penjelasan = extract_nama_refined_v5(block, amt)

        txs.append({
            'date': date_str, 'nama': nama_orang, 'ket': ket, 'kode': kode,
            'debet': amt if is_db else 0.0, 'kredit': amt if (is_cr and not is_db) else 0.0,
            'penjelasan': penjelasan
        })

    nama_akun = 'CV. MITRA JAYA ANUGERAH'
    for l in all_lines[:15]:
        if 'CV.' in l.upper() or 'PT.' in l.upper():
            nama_akun = l.strip()
            break

    return {
        'bank': 'BCA', 'no_rek': no_rek, 'nama_akun': nama_akun, 'period': period,
        'saldo_awal': saldo_awal, 'saldo_akhir': saldo_akhir, 
        'mut_cr': mut_cr, 'mut_db': mut_db, 'txs': txs,
    }

# ─────────────────────────────────────────────
# PARSER BRI & MANDIRI (TETAP SAMA)
# ─────────────────────────────────────────────

def parse_bri(all_lines: list[str]) -> dict:
    period = ''; saldo_awal = saldo_akhir = total_db = total_cr = 0
    no_rek = nama_akun = ''
    for l in all_lines:
        m = re.search(r'(?:Periode Transaksi|Transaction Period)[^\d]*(\d{2}/\d{2}/\d{2})\s*-\s*(\d{2}/\d{2}/\d{2})', l)
        if m and not period: period = f"{m.group(1)} - {m.group(2)}"
        m = re.search(r'No\.\s*Rekening[^\d]*([\d]+)', l)
        if m and not no_rek: no_rek = m.group(1)
        if l.strip().startswith('CV ') or l.strip().startswith('PT '):
            if not nama_akun: nama_akun = l.strip()
        m = re.match(r'^([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})$', l.strip())
        if m:
            saldo_awal, total_db, total_cr, saldo_akhir = [float(x.replace(',', '')) for x in m.groups()]

    TX_RE = re.compile(r'^(\d{2}/\d{2}/\d{2})\s+\d{2}:\d{2}:\d{2}\s+(.+?)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})$')
    txs = []
    for l in all_lines:
        m = TX_RE.match(l.strip())
        if not m: continue
        date_raw, desc, db_s, cr_s, _ = m.groups()
        db, cr = float(db_s.replace(',', '')), float(cr_s.replace(',', ''))
        ket, kode = classify(desc, cr > 0)
        txs.append({'date': date_raw, 'nama': '', 'ket': ket, 'kode': kode, 'debet': db, 'kredit': cr, 'penjelasan': desc.strip()})

    return {'bank': 'BRI', 'no_rek': no_rek, 'nama_akun': nama_akun or 'NASABAH BRI', 'period': period, 'saldo_awal': saldo_awal, 'saldo_akhir': saldo_akhir, 'mut_cr': total_cr, 'mut_db': total_db, 'txs': txs}

def parse_mandiri(all_lines: list[str]) -> dict:
    _MANDIRI_MONTH = {'Jan':'01','Feb':'02','Mar':'03','Apr':'04','May':'05','Jun':'06','Jul':'07','Aug':'08','Sep':'09','Oct':'10','Nov':'11','Dec':'12'}
    period = ''; saldo_awal = saldo_akhir = total_db = total_cr = 0
    no_rek = nama_akun = ''
    for l in all_lines:
        m = re.search(r'(\d{2}\s+\w+\s+\d{4})\s*-\s*(\d{2}\s+\w+\s+\d{4})', l)
        if m and not period: period = f"{m.group(1)} - {m.group(2)}"
        m = re.match(r'^(\d{10,16})\s+(.+)', l.strip())
        if m and not no_rek:
            no_rek = m.group(1)
            nama_akun = m.group(2).split('  ')[0].strip()
        m = re.match(r'^([\d,]+\.\d{2})\s+\d+\s+([\d,]+\.\d{2})$', l.strip())
        if m:
            v1, v2 = float(m.group(1).replace(',', '')), float(m.group(2).replace(',', ''))
            if saldo_awal == 0: saldo_awal = v1; total_db = v2
            else: saldo_akhir = v1; total_cr = v2

    TX_AMT_RE = re.compile(r'^(.*?)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d.2})\s+([\d,]+\.\d{2})\s*$')
    DATE_RE = re.compile(r'(\d{2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{4})', re.IGNORECASE)
    txs = []; last_date = ''
    for i, l in enumerate(all_lines):
        ls = l.strip()
        dm = DATE_RE.search(ls)
        if dm: last_date = f"{dm.group(1)}/{_MANDIRI_MONTH.get(dm.group(2)[:3].capitalize(), '01')}/{dm.group(3)[-2:]}"
        m = TX_AMT_RE.match(ls)
        if not m or not last_date: continue
        desc = re.sub(r'\d{2}:\d{2}:\d{2}', '', m.group(1)).strip()
        db, cr = float(m.group(2).replace(',', '')), float(m.group(3).replace(',', ''))
        ket, kode = classify(desc, cr > 0)
        txs.append({'date': last_date, 'nama': '', 'ket': ket, 'kode': kode, 'debet': db, 'kredit': cr, 'penjelasan': desc})

    return {'bank': 'MANDIRI', 'no_rek': no_rek, 'nama_akun': nama_akun or 'NASABAH MANDIRI', 'period': period, 'saldo_awal': saldo_awal, 'saldo_akhir': saldo_akhir, 'mut_cr': total_cr, 'mut_db': total_db, 'txs': txs}

# ─────────────────────────────────────────────
# CORE ENGINE & EXCEL GEN
# ─────────────────────────────────────────────

def clean_excel_string(val):
    if isinstance(val, str):
        return ILLEGAL_CHARACTERS_RE.sub('', val)
    return val

def parse_pdf(pdf_bytes: bytes) -> dict:
    all_lines = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text: all_lines.extend(text.split('\n'))
    
    if not all_lines: raise ValueError("PDF kosong atau tidak terbaca.")
    
    bank = detect_bank(all_lines)
    if bank == 'BCA': return parse_bca(all_lines)
    elif bank == 'BRI': return parse_bri(all_lines)
    elif bank == 'MANDIRI': return parse_mandiri(all_lines)
    else: raise ValueError(f"Format bank tidak dikenali. Pastikan ini PDF Mutasi (Deteksi: {bank})")

def create_excel(data: dict) -> io.BytesIO:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Mutasi {data['bank']}"[:31]
    
    colors = {'header': '005BAC', 'sub': 'E8F0FA'} # Default
    if 'BRI' in data['bank']: colors = {'header': '003D7C', 'sub': 'E6EEF7'}
    elif 'MANDIRI' in data['bank']: colors = {'header': '003087', 'sub': 'E5ECF6'}

    # Header Rendering
    ws.merge_cells('A1:I1'); ws['A1'] = f"LAPORAN MUTASI – {data['bank']}"
    ws['A1'].font = Font(bold=True, color='FFFFFF', size=14); ws['A1'].fill = PatternFill('solid', start_color=colors['header'])
    ws['A1'].alignment = Alignment(horizontal='center')
    
    ws.merge_cells('A2:I2'); ws['A2'] = clean_excel_string(data['nama_akun'])
    ws['A2'].font = Font(bold=True); ws['A2'].alignment = Alignment(horizontal='center')
    
    ws.merge_cells('A3:I3'); ws['A3'] = f"Rek: {data['no_rek']} | Periode: {data['period']}"
    ws['A3'].alignment = Alignment(horizontal='center')

    headers = ['NO', 'TANGGAL', 'NAMA', 'KETERANGAN', 'KODE', 'DEBET', 'KREDIT', 'SALDO', 'PENJELASAN']
    for i, h in enumerate(headers, 1):
        cell = ws.cell(6, i, h)
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', start_color=colors['header'])
        cell.alignment = Alignment(horizontal='center')

    saldo = data['saldo_awal']
    for idx, tx in enumerate(data['txs'], 1):
        saldo = round(saldo + tx['kredit'] - tx['debet'], 2)
        row = idx + 6
        vals = [idx, tx['date'], tx['nama'], tx['ket'], tx['kode'], tx['debet'], tx['kredit'], saldo, tx['penjelasan']]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(row, ci, clean_excel_string(v))
            if ci >= 6 and isinstance(v, (int, float)): c.number_format = '#,##0.00'

    for col in ['A','B','C','D','E','F','G','H','I']:
        ws.column_dimensions[col].width = 15
    ws.column_dimensions['C'].width = 25
    ws.column_dimensions['D'].width = 25
    ws.column_dimensions['I'].width = 40

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

# ─────────────────────────────────────────────
# FLASK ROUTES
# ─────────────────────────────────────────────

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8">
<title>Bank Statement Fix V2.3</title>
<style>
    body { font-family: sans-serif; background: #f0f2f5; padding: 40px; display: flex; justify-content: center; }
    .card { background: white; padding: 30px; border-radius: 12px; box-shadow: 0 4px 20px rgba(0,0,0,0.1); width: 100%; max-width: 500px; }
    h2 { color: #1a73e8; margin-top: 0; }
    input[type=file] { margin: 20px 0; display: block; }
    button { background: #1a73e8; color: white; border: none; padding: 12px 20px; border-radius: 6px; cursor: pointer; width: 100%; font-weight: bold; }
    button:disabled { background: #ccc; }
    #status { margin-top: 20px; padding: 15px; border-radius: 6px; display: none; font-size: 14px; }
    .err { background: #fde8e8; color: #c81e1e; border: 1px solid #f8b4b4; }
    .ok { background: #def7ec; color: #03543f; border: 1px solid #bcf0da; }
</style>
</head>
<body>
<div class="card">
    <h2>Bank Statement Converter</h2>
    <p style="font-size: 13px; color: #666;">Silakan upload PDF BCA, BRI, atau Mandiri.</p>
    <input type="file" id="files" multiple accept=".pdf">
    <button id="btn" onclick="upload()">GENERATE EXCEL</button>
    <div id="status"></div>
</div>

<script>
async function upload() {
    const btn = document.getElementById('btn');
    const status = document.getElementById('status');
    const files = document.getElementById('files').files;
    if (files.length === 0) return;

    btn.disabled = true;
    status.style.display = 'block';
    status.className = 'ok';
    status.innerText = 'Sedang memproses...';

    const fd = new FormData();
    for (let f of files) fd.append('files', f);

    try {
        const res = await fetch('/convert', { method: 'POST', body: fd });
        if (!res.ok) {
            const errText = await res.text();
            throw new Error(errText);
        }
        
        const blob = await res.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `Mutasi_Rekening_${new Date().getTime()}.xlsx`;
        document.body.appendChild(a);
        a.click();
        a.remove();
        
        status.innerText = 'Berhasil! File Excel telah diunduh.';
    } catch (e) {
        status.className = 'err';
        status.innerText = 'Gagal: ' + e.message;
    } finally {
        btn.disabled = false;
    }
}
</script>
</body>
</html>"""

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/convert', methods=['POST'])
def convert():
    try:
        uploaded = request.files.getlist('files')
        if not uploaded: return 'Pilih file terlebih dahulu', 400

        all_txs = []
        bank_set = set(); period_set = set(); no_rek_set = set()
        saldo_awal = None; saldo_akhir = 0
        mut_cr = mut_db = 0; nama_akun = ''

        for f in uploaded:
            data = parse_pdf(f.read())
            if saldo_awal is None: 
                saldo_awal = data['saldo_awal']
                nama_akun = data['nama_akun']
            saldo_akhir = data['saldo_akhir']
            mut_cr += data['mut_cr']
            mut_db += data['mut_db']
            all_txs.extend(data['txs'])
            bank_set.add(data['bank'])
            period_set.add(data['period'])
            no_rek_set.add(data['no_rek'])

        merged = {
            'bank': ' + '.join(bank_set),
            'no_rek': ' / '.join(no_rek_set),
            'nama_akun': nama_akun,
            'period': ' - '.join(period_set),
            'saldo_awal': saldo_awal,
            'saldo_akhir': saldo_akhir,
            'mut_cr': mut_cr,
            'mut_db': mut_db,
            'txs': all_txs
        }

        excel_file = create_excel(merged)
        
        # FIX COMPATIBILITY: Kirim file dengan cara manual agar support Flask versi lama
        response = make_response(send_file(
            excel_file, 
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            attachment_filename='Mutasi_Bank.xlsx' # Menggunakan parameter lama untuk kompatibilitas
        ))
        return response

    except Exception as e:
        traceback.print_exc()
        return str(e), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)