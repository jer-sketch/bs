"""
Bank Statement PDF → Excel Converter
Mendukung: BCA (V5 Advanced), BRI (IBIZ), Mandiri (Kopra)
UI & Excel Styling: Modern Dark Mode
"""

import re
import io
from datetime import datetime

import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from flask import Flask, request, send_file, render_template_string

app = Flask(__name__)

# ─────────────────────────────────────────────
# KLASIFIKASI KETERANGAN & KODE (COMMON UTK BRI & MANDIRI)
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
    elif 'SETORAN' in u or 'KLIRING' in u:              ket = 'SETORAN TUNAI'
    elif 'KOREKSI' in u:                                ket = 'KOREKSI BANK'
    elif not is_credit:                                 ket = 'PELUNASAN HUTANG DAGANG'
    else:                                               ket = 'PENERIMAAN PENJUALAN'
    return ket, KODE_MAP.get(ket, '')

# ─────────────────────────────────────────────
# DETEKTOR JENIS BANK
# ─────────────────────────────────────────────
def detect_bank(all_lines: list[str]) -> str:
    text = ' '.join(all_lines[:30]).upper()
    if 'REKENING GIRO' in text or 'TRSF E-BANKING' in text or 'BI-FAST' in text: return 'BCA'
    if 'LAPORAN TRANSAKSI FINANSIAL' in text or 'BRIMTXDT' in text or 'BRITAMA' in text: return 'BRI'
    if 'KOPRA' in text or 'MANDIRI' in text or 'MCM INHOUSETRF' in text or 'ACCOUNT STATEMENT' in text: return 'MANDIRI'
    return 'UNKNOWN'

# ─────────────────────────────────────────────
# PARSER BCA (MENGGUNAKAN LOGIC V5 KITA)
# ─────────────────────────────────────────────
def extract_nama_refined_bca(block, amount):
    amt_str = "{:.2f}".format(amount)
    found_idx = -1
    for i, line in enumerate(block):
        if amt_str in line.replace(',', ''):
            found_idx = i; break
    if found_idx == -1: return ""

    candidates = block[found_idx+1:]
    desc_list = []
    stoppers = ['KCU', 'PERIODE', 'MATA UANG', 'IDR', 'INDONESIA', 'CATATAN', 'JL ', 'BANDUNG', 'TANGGAL', 'MUTASI', 'SALDO', 'HALAMAN']
    
    for cand in candidates:
        c_strip = cand.strip()
        if not c_strip: continue
        if any(s in c_strip.upper() for s in stoppers): break
        desc_list.append(c_strip)
        
    raw_desc = " ".join(desc_list).strip()
    return raw_desc

def parse_bca(all_lines: list[str]) -> dict:
    period = next((l.split(':', 1)[1].strip() for l in all_lines if 'PERIODE :' in l.upper()), "")
    year = period.split()[-1] if period else str(datetime.now().year)
    no_rek = next((re.search(r':\s*([\d]+)', l).group(1) for l in all_lines if 'NO. REKENING' in l.upper() and re.search(r':\s*([\d]+)', l)), "")
    nama_akun = next((l.strip() for l in all_lines[:15] if 'CV.' in l.upper() or 'PT.' in l.upper()), "NASABAH BCA")
    
    saldo_awal = 0
    for l in all_lines:
        m = re.search(r'SALDO AWAL\s*:\s*([\d,]+\.\d+)', l.upper())
        if m: saldo_awal = float(m.group(1).replace(',','')); break

    DATE_RE = re.compile(r'^(\d{2})/(\d{2})\s+')
    tx_blocks, current_block = [], []
    for l in all_lines:
        if DATE_RE.match(l):
            if current_block: tx_blocks.append(current_block)
            current_block = [l]
        elif current_block: current_block.append(l)
    if current_block: tx_blocks.append(current_block)

    txs, total_db, total_cr = [], 0, 0
    for block in tx_blocks:
        full_text = " ".join(block).upper()
        if 'SALDO AWAL' in full_text and len(block) < 3: continue
        
        m = DATE_RE.match(block[0])
        date_str = f"{m.group(1)}/{m.group(2)}/{year[-2:]}"
        amounts = re.findall(r'([\d,]+\.\d{2})', full_text)
        if not amounts: continue
        amt = float(amounts[0].replace(',', ''))
        
        # LOGIC BCA V5 UTK MENDETEKSI SETORAN KLIRING DLL
        is_cr = any(x in full_text for x in [' CR', 'SETORAN', 'KLIRING', 'KR OTOMATIS', 'BUNGA', 'SWITCHING CR', 'BI-FAST CR'])
        is_db = any(x in full_text for x in [' DB', 'BYR VIA', 'BIAYA ADM', 'PAJAK', 'BI-FAST DB', 'BA JASA', 'TARIKAN', 'SWITCHING DB'])
        
        if 'PAJAK' in full_text: is_db, is_cr = True, False
        if 'BUNGA' in full_text and 'PAJAK' not in full_text: is_db, is_cr = False, True

        db, cr = (amt, 0) if is_db else (0, amt if is_cr else 0)
        
        # Fallback jika tidak terdeteksi DB/CR (Asumsi berdasar pola BCA)
        if db == 0 and cr == 0: cr = amt 
            
        ket, kode = classify(full_text, cr > 0)
        penjelasan = extract_nama_refined_bca(block, amt) or full_text.split(amounts[0])[0].replace(block[0][:5], '').strip()
        
        total_db += db; total_cr += cr
        txs.append({'date': date_str, 'penjelasan': penjelasan, 'ket': ket, 'kode': kode, 'debet': db, 'kredit': cr})

    return {'bank': 'BCA', 'no_rek': no_rek, 'nama_akun': nama_akun, 'period': period, 'saldo_awal': saldo_awal, 'saldo_akhir': 0, 'mut_cr': total_cr, 'mut_db': total_db, 'txs': txs}

# ─────────────────────────────────────────────
# PARSER BRI (DARI CLAUDE)
# ─────────────────────────────────────────────
def parse_bri(all_lines: list[str]) -> dict:
    period = saldo_awal = total_db = total_cr = 0
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
            saldo_awal = float(m.group(1).replace(',', ''))
            total_db = float(m.group(2).replace(',', ''))
            total_cr = float(m.group(3).replace(',', ''))

    TX_RE = re.compile(r'^(\d{2}/\d{2}/\d{2})\s+\d{2}:\d{2}:\d{2}\s+(.+?)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})$')
    txs = []
    for l in all_lines:
        m = TX_RE.match(l.strip())
        if not m: continue
        date_raw, desc, db_s, cr_s, _ = m.groups()
        db, cr = float(db_s.replace(',', '')), float(cr_s.replace(',', ''))
        ket, kode = classify(desc, cr > 0)
        txs.append({'date': date_raw, 'penjelasan': desc.strip(), 'ket': ket, 'kode': kode, 'debet': db, 'kredit': cr})

    return {'bank': 'BRI', 'no_rek': no_rek, 'nama_akun': nama_akun or 'NASABAH BRI', 'period': period, 'saldo_awal': saldo_awal, 'saldo_akhir': 0, 'mut_cr': total_cr, 'mut_db': total_db, 'txs': txs}

# ─────────────────────────────────────────────
# PARSER MANDIRI (DARI CLAUDE - KOPRA FORMAT)
# ─────────────────────────────────────────────
_MANDIRI_MONTH = {'Jan':'01','Feb':'02','Mar':'03','Apr':'04','May':'05','Jun':'06','Jul':'07','Aug':'08','Sep':'09','Oct':'10','Nov':'11','Dec':'12'}

def parse_mandiri(all_lines: list[str]) -> dict:
    period = saldo_awal = total_db = total_cr = 0
    no_rek = nama_akun = ''

    for l in all_lines:
        m = re.search(r'(\d{2}\s+\w+\s+\d{4})\s*-\s*(\d{2}\s+\w+\s+\d{4})', l)
        if m and not period: period = f"{m.group(1)} - {m.group(2)}"
        m = re.match(r'^(\d{10,16})\s+(.+)', l.strip())
        if m and not no_rek: no_rek, nama_akun = m.group(1), m.group(2).split('  ')[0].strip()
        m = re.match(r'^([\d,]+\.\d{2})\s+\d+\s+([\d,]+\.\d{2})$', l.strip())
        if m:
            v1, v2 = float(m.group(1).replace(',', '')), float(m.group(2).replace(',', ''))
            if saldo_awal == 0: saldo_awal, total_db = v1, v2
            else: total_cr = v2

    TX_AMT_RE = re.compile(r'^(.*?)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s*$')
    DATE_RE = re.compile(r'(\d{2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{4})', re.IGNORECASE)
    SKIP = ('For further questions', 'Account Statement', 'Created', 'Posting Date', 'Opening Balance', 'Closing Balance', 'No. of Debit', 'Total Amount', 'Account Statement Summary', 'Account No.', 'Period ', 'Alias', 'Currency', 'Branch', 'kopra')

    txs, last_date = [], ''
    for i, l in enumerate(all_lines):
        ls = l.strip()
        if not ls or any(ls.lower().startswith(s.lower()) for s in SKIP): continue

        dm = DATE_RE.search(ls)
        if dm:
            mon = _MANDIRI_MONTH.get(dm.group(2)[:3].capitalize(), '01')
            last_date = f"{dm.group(1)}/{mon}/{dm.group(3)[-2:]}"

        m = TX_AMT_RE.match(ls)
        if not m: continue
        prefix = m.group(1).strip()
        db, cr = float(m.group(2).replace(',', '')), float(m.group(3).replace(',', ''))

        if any(prefix.lower().startswith(s.lower()) for s in ('Closing', 'Opening', 'Total', 'Terbilang')): continue

        desc = re.sub(r'\d{2}:\d{2}:\d{2}', '', prefix)
        desc = re.sub(r'\b\d{8,}\b', '', desc)
        desc = re.sub(r'\s{2,}', ' ', desc).strip(' -')

        for offset in [-2, -1, 1, 2]:
            ni = i + offset
            if 0 <= ni < len(all_lines):
                nb = all_lines[ni].strip()
                if nb and not TX_AMT_RE.match(nb) and not DATE_RE.search(nb) and not any(nb.lower().startswith(s.lower()) for s in SKIP) and not re.match(r'^\d{2}:\d{2}:\d{2}', nb):
                    desc = (desc + ' ' + nb).strip(); break

        if not last_date: continue
        ket, kode = classify(desc or ls, cr > 0)
        txs.append({'date': last_date, 'penjelasan': desc, 'ket': ket, 'kode': kode, 'debet': db, 'kredit': cr})

    return {'bank': 'MANDIRI', 'no_rek': no_rek, 'nama_akun': nama_akun or 'NASABAH MANDIRI', 'period': period, 'saldo_awal': saldo_awal, 'saldo_akhir': 0, 'mut_cr': total_cr, 'mut_db': total_db, 'txs': txs}

def parse_pdf(pdf_bytes: bytes) -> dict:
    all_lines = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text: all_lines.extend(text.split('\n'))

    bank = detect_bank(all_lines)
    if bank == 'BCA': return parse_bca(all_lines)
    elif bank == 'BRI': return parse_bri(all_lines)
    elif bank == 'MANDIRI': return parse_mandiri(all_lines)
    else: raise ValueError(f"Format bank tidak dikenali. Pastikan PDF adalah mutasi BCA, BRI, atau Mandiri.")

# ─────────────────────────────────────────────
# GENERATOR EXCEL (STYLE DARI CLAUDE + PERBAIKAN KOLOM)
# ─────────────────────────────────────────────
BANK_COLORS = {
    'BCA':     {'header': '005BAC', 'sub': 'E8F0FA'},
    'BRI':     {'header': '003D7C', 'sub': 'E6EEF7'},
    'MANDIRI': {'header': '003087', 'sub': 'E5ECF6'},
}

def create_excel(data: dict) -> io.BytesIO:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Mutasi {data['bank'][:30]}"

    bank = data['bank'].split(' + ')[0] if ' + ' in data['bank'] else data['bank']
    colors = BANK_COLORS.get(bank, BANK_COLORS['BCA'])
    NUM = '#,##0.00'

    def style_cell(cell, bold=False, color=None, bg=None, align='left', num_fmt=None):
        cell.font = Font(name='Calibri', bold=bold, color=color or '000000', size=10)
        cell.alignment = Alignment(horizontal=align, vertical='center', wrap_text=False)
        if bg: cell.fill = PatternFill('solid', start_color=bg)
        if num_fmt: cell.number_format = num_fmt

    thin = Side(style='thin', color='CCCCCC')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.merge_cells('A1:H1')
    ws['A1'] = f"LAPORAN MUTASI REKENING – {data['bank']}"
    style_cell(ws['A1'], bold=True, color='FFFFFF', bg=colors['header'], align='center')
    ws['A1'].font = Font(name='Calibri', bold=True, color='FFFFFF', size=13)

    ws.merge_cells('A2:H2'); ws['A2'] = data['nama_akun']
    style_cell(ws['A2'], bold=True, align='center', bg=colors['sub'])

    ws.merge_cells('A3:H3'); ws['A3'] = f"No. Rekening: {data['no_rek']}    |    Periode: {data['period']}"
    style_cell(ws['A3'], align='center', bg=colors['sub'])

    ws.row_dimensions[4].height = 14
    summaries = [('Saldo Awal', data['saldo_awal']), ('Total Kredit', data['mut_cr']), ('Total Debet', data['mut_db'])]
    
    # Header Kolom (Ditambahkan NAMA / DESKRIPSI)
    headers = ['NO', 'TANGGAL', 'NAMA / DESKRIPSI', 'KETERANGAN', 'KODE', 'DEBET', 'KREDIT', 'SALDO']
    header_cols = ['A','B','C','D','E','F','G','H']
    for col_letter, hdr in zip(header_cols, headers):
        c = ws[f'{col_letter}6']
        c.value = hdr
        style_cell(c, bold=True, color='FFFFFF', bg=colors['header'], align='center')
        c.border = border

    saldo = data['saldo_awal']
    for idx, tx in enumerate(data['txs'], 1):
        saldo = round(saldo + tx['kredit'] - tx['debet'], 2)
        row = idx + 6
        bg = 'FFFFFF' if idx % 2 == 0 else 'F8FAFF'

        values = [idx, tx['date'], tx.get('penjelasan', ''), tx['ket'], tx['kode'], tx['debet'] or '', tx['kredit'] or '', saldo]

        for col_idx, val in enumerate(values, 1):
            c = ws.cell(row, col_idx, val)
            aln = 'center' if col_idx in (1,2,5) else ('right' if col_idx >= 6 else 'left')
            num = NUM if col_idx >= 6 and isinstance(val, float) and val > 0 else None
            style_cell(c, align=aln, bg=bg, num_fmt=num)
            c.border = border

    col_widths = {'A': 6, 'B': 12, 'C': 40, 'D': 25, 'E': 7, 'F': 16, 'G': 16, 'H': 18}
    for col, width in col_widths.items(): ws.column_dimensions[col].width = width

    ws.freeze_panes = 'A7'
    out = io.BytesIO()
    wb.save(out); out.seek(0)
    return out

# ─────────────────────────────────────────────
# FLASK ROUTES & HTML UI CLAUDE
# ─────────────────────────────────────────────
HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0">
<title>Bank Statement Converter</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap');
  :root { --bg: #0a0e1a; --surface: #111827; --card: #161d2e; --border: #1e2d45; --accent: #3b82f6; --accent2: #06b6d4; --text: #e2e8f0; --muted: #64748b; --success: #10b981; --danger: #ef4444; --radius: 16px; }
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'IBM Plex Sans', sans-serif; background: var(--bg); color: var(--text); min-height: 100vh; display: flex; flex-direction: column; align-items: center; padding: 24px 16px 48px; }
  body::before { content: ''; position: fixed; inset: 0; background-image: linear-gradient(rgba(59,130,246,.04) 1px, transparent 1px), linear-gradient(90deg, rgba(59,130,246,.04) 1px, transparent 1px); background-size: 40px 40px; pointer-events: none; z-index: 0; }
  .container { position: relative; z-index: 1; width: 100%; max-width: 520px; }
  .header { text-align: center; margin-bottom: 32px; }
  .logo-wrap { display: inline-flex; align-items: center; justify-content: center; width: 64px; height: 64px; background: linear-gradient(135deg, var(--accent), var(--accent2)); border-radius: 18px; margin-bottom: 16px; box-shadow: 0 0 32px rgba(59,130,246,.35); }
  .logo-wrap svg { width: 32px; height: 32px; color: white; }
  h1 { font-size: clamp(20px,5vw,28px); font-weight: 700; letter-spacing: -0.5px; }
  .subtitle { color: var(--muted); font-size: 14px; margin-top: 6px; }
  .banks { display: flex; gap: 8px; justify-content: center; margin-top: 16px; flex-wrap: wrap; }
  .badge { font-family: 'IBM Plex Mono', monospace; font-size: 11px; font-weight: 500; padding: 4px 10px; border-radius: 6px; border: 1px solid; }
  .badge-bca { color: #60a5fa; border-color: #1e3a5f; background: #0a1929; }
  .badge-bri { color: #34d399; border-color: #1a3a2e; background: #0a1f18; }
  .badge-mandiri{ color: #fbbf24; border-color: #3a2f0a; background: #1f1800; }
  .card { background: var(--card); border: 1px solid var(--border); border-radius: var(--radius); overflow: hidden; margin-bottom: 16px; }
  .card-header { background: var(--surface); border-bottom: 1px solid var(--border); padding: 14px 20px; font-size: 12px; font-weight: 600; text-transform: uppercase; letter-spacing: .08em; color: var(--muted); display: flex; align-items: center; gap: 8px; }
  .card-header::before { content: ''; display: block; width: 3px; height: 14px; background: linear-gradient(var(--accent), var(--accent2)); border-radius: 2px; }
  .card-body { padding: 20px; }
  .drop-zone { border: 2px dashed var(--border); border-radius: 12px; padding: 32px 20px; text-align: center; cursor: pointer; transition: all .2s; background: rgba(59,130,246,.03); position: relative; }
  .drop-zone:hover, .drop-zone.dragover { border-color: var(--accent); background: rgba(59,130,246,.08); }
  .drop-zone input[type=file] { position: absolute; inset: 0; opacity: 0; cursor: pointer; }
  .drop-icon { width: 44px; height: 44px; background: linear-gradient(135deg, rgba(59,130,246,.2), rgba(6,182,212,.2)); border-radius: 12px; display: flex; align-items: center; justify-content: center; margin: 0 auto 12px; }
  .drop-icon svg { width: 22px; height: 22px; color: var(--accent); }
  .drop-label { font-size: 14px; color: var(--muted); line-height: 1.5; }
  .drop-label strong { color: var(--text); }
  .file-list { margin-top: 16px; display: flex; flex-direction: column; gap: 8px; }
  .file-item { display: flex; align-items: center; gap: 10px; background: var(--surface); border: 1px solid var(--border); border-radius: 8px; padding: 8px 12px; font-size: 13px; }
  .file-item .rm { cursor: pointer; color: var(--muted); font-size: 18px; line-height: 1; transition: color .15s; flex-shrink: 0; }
  .file-item .rm:hover { color: var(--danger); }
  .btn { width: 100%; padding: 14px; background: linear-gradient(135deg, var(--accent) 0%, var(--accent2) 100%); color: white; font-weight: 700; font-size: 15px; border: none; border-radius: 12px; cursor: pointer; transition: all .2s; display: flex; align-items: center; justify-content: center; gap: 8px; }
  .btn:hover:not(:disabled) { transform: translateY(-1px); box-shadow: 0 6px 24px rgba(59,130,246,.45); }
  .btn:disabled { opacity: .45; cursor: not-allowed; }
  .progress-wrap { height: 3px; background: var(--border); border-radius: 2px; overflow: hidden; margin-top: 12px; display: none; }
  .progress-bar { height: 100%; width: 0%; background: linear-gradient(90deg, var(--accent), var(--accent2)); transition: width .4s ease; }
  .status-row { display: flex; align-items: center; gap: 10px; font-size: 13px; color: var(--muted); margin-top: 12px; }
  .dot { width: 7px; height: 7px; border-radius: 50%; background: var(--muted); }
  .dot.proc { background: var(--accent); animation: pulse 1s infinite; }
  .dot.ok   { background: var(--success); }
  .dot.err  { background: var(--danger); }
  @keyframes pulse { 0%,100%{opacity:1} 50%{opacity:.3} }
  .result-box { background: rgba(16,185,129,.08); border: 1px solid rgba(16,185,129,.3); border-radius: 10px; padding: 14px 16px; margin-top: 12px; display: none; align-items: center; gap: 12px; font-size: 13px; }
  .result-box.show { display: flex; }
  .result-text strong { display: block; color: var(--success); font-size: 14px; margin-bottom: 2px; }
  footer { color: var(--muted); font-size: 12px; text-align: center; margin-top: 8px; line-height: 1.7; }
</style>
</head>
<body>
<div class="container">
  <div class="header">
    <div class="logo-wrap">
      <svg fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 17v-2m3 2v-4m3 4v-6m2 10H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/></svg>
    </div>
    <h1>Bank Statement Converter</h1>
    <p class="subtitle">Konversi mutasi PDF ke Excel secara otomatis</p>
    <div class="banks">
      <span class="badge badge-bca">✦ BCA (V5)</span>
      <span class="badge badge-bri">✦ BRI</span>
      <span class="badge badge-mandiri">✦ Mandiri</span>
    </div>
  </div>

  <div class="card">
    <div class="card-header">Upload File PDF</div>
    <div class="card-body">
      <div class="drop-zone" id="dropZone">
        <input type="file" id="fileInput" multiple accept=".pdf">
        <div class="drop-icon">
          <svg fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"/></svg>
        </div>
        <div class="drop-label"><strong>Klik atau seret file ke sini</strong><br>Bisa pilih banyak file PDF sekaligus</div>
      </div>
      <div class="file-list" id="fileList"></div>
    </div>
  </div>

  <div class="card">
    <div class="card-header">Proses Konversi</div>
    <div class="card-body">
      <button class="btn" id="convertBtn" disabled onclick="startConvert()">Generate Excel</button>
      <div class="progress-wrap" id="progressWrap"><div class="progress-bar" id="progressBar"></div></div>
      <div class="status-row">
        <div class="dot" id="statusDot"></div><span id="statusText">Pilih file PDF untuk memulai</span>
      </div>
      <div class="result-box" id="resultBox">
        <div class="result-icon">✅</div><div class="result-text" id="resultText"></div>
      </div>
    </div>
  </div>
</div>

<script>
const files = {};
const dz = document.getElementById('dropZone'), fi = document.getElementById('fileInput'), btn = document.getElementById('convertBtn'), list = document.getElementById('fileList');
dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('dragover'); });
dz.addEventListener('dragleave', () => dz.classList.remove('dragover'));
dz.addEventListener('drop', e => { e.preventDefault(); dz.classList.remove('dragover'); handleFiles([...e.dataTransfer.files]); });
fi.addEventListener('change', e => handleFiles([...e.target.files]));
dz.addEventListener('click', e => { if (e.target === dz || e.target.closest('.drop-icon, .drop-label')) fi.click(); });

function handleFiles(newFiles) { newFiles.filter(f => f.name.toLowerCase().endsWith('.pdf')).forEach(f => files[f.name] = f); renderList(); }
function renderList() {
  list.innerHTML = '';
  Object.keys(files).forEach(name => {
    const div = document.createElement('div'); div.className = 'file-item';
    div.innerHTML = `<span class="fname">${name}</span><span class="rm" data-name="${name}">×</span>`;
    list.appendChild(div);
  });
  list.querySelectorAll('.rm').forEach(el => el.addEventListener('click', e => { delete files[e.target.dataset.name]; renderList(); btn.disabled = Object.keys(files).length === 0; }));
  btn.disabled = Object.keys(files).length === 0;
}
function setStatus(state, text) { document.getElementById('statusDot').className = 'dot ' + state; document.getElementById('statusText').textContent = text; }
function setProgress(pct) { document.getElementById('progressWrap').style.display = pct>0?'block':'none'; document.getElementById('progressBar').style.width = pct+'%'; }

async function startConvert() {
  const flist = Object.values(files); if (!flist.length) return;
  btn.disabled = true; document.getElementById('resultBox').classList.remove('show');
  setStatus('proc', `Mengirim ${flist.length} file...`); setProgress(20);
  const fd = new FormData(); flist.forEach(f => fd.append('files', f));

  try {
    const res = await fetch('/convert', { method: 'POST', body: fd }); setProgress(80);
    if (!res.ok) throw new Error(await res.text());
    const blob = await res.blob();
    const fname = res.headers.get('X-Filename') || 'Mutasi_Bank.xlsx';
    const url = URL.createObjectURL(blob), a = document.createElement('a');
    a.href = url; a.download = fname; a.click(); URL.revokeObjectURL(url);
    setProgress(100); setStatus('ok', `Selesai!`);
    document.getElementById('resultText').innerHTML = `<strong>${fname}</strong>File Excel siap di folder unduhan.`;
    document.getElementById('resultBox').classList.add('show');
    setTimeout(() => setProgress(0), 1500);
  } catch (err) { setStatus('err', 'Gagal: ' + err.message); setProgress(0); }
  finally { btn.disabled = false; }
}
</script>
</body>
</html>"""

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/convert', methods=['POST'])
def convert():
    uploaded = request.files.getlist('files')
    if not uploaded: return 'Tidak ada file yang diunggah.', 400

    all_txs = []
    period_set, bank_set, no_rek_set = set(), set(), set()
    saldo_awal = None; saldo_akhir = mut_cr = mut_db = 0
    nama_akun = ''

    for f in uploaded:
        raw = f.read()
        try:
            data = parse_pdf(raw)
        except Exception as e:
            return f'Error memproses {f.filename}: {str(e)}', 400

        if saldo_awal is None: 
            saldo_awal = data['saldo_awal']
            nama_akun = data['nama_akun']

        saldo_akhir = data['saldo_akhir']
        mut_cr += data['mut_cr']; mut_db += data['mut_db']
        all_txs.extend(data['txs'])
        if data['period']: period_set.add(data['period'])
        if data['bank']: bank_set.add(data['bank'])
        if data['no_rek']: no_rek_set.add(data['no_rek'])

    merged = {
        'bank':       ' + '.join(sorted(bank_set)),
        'no_rek':     ' | '.join(sorted(no_rek_set)),
        'nama_akun':  nama_akun,
        'period':     ' | '.join(sorted(period_set)) if period_set else '-',
        'saldo_awal': saldo_awal or 0,
        'saldo_akhir': saldo_akhir,
        'mut_cr':     mut_cr,
        'mut_db':     mut_db,
        'txs':        all_txs,
    }

    excel = create_excel(merged)
    banks_str = '_'.join(sorted(bank_set))
    fname = f"Mutasi_{banks_str}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    
    return send_file(excel, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name=fname, headers={'X-Filename': fname})

if __name__ == '__main__':
    app.run(debug=True, port=5000)