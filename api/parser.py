import re
import pdfplumber

MONTH_MAP = {'JANUARI':1,'FEBRUARI':2,'MARET':3,'APRIL':4,'MEI':5,'JUNI':6,
             'JULI':7,'AGUSTUS':8,'SEPTEMBER':9,'OKTOBER':10,'NOVEMBER':11,'DESEMBER':12}

def classify(line, is_credit):
    u = line.upper()
    if 'BIAYA ADM' in u: return 'BIAYA ADM BANK','27'
    if 'PAJAK BUNGA' in u or 'PAJAK JASA GIRO' in u: return 'PAJAK JASA GIRO',''
    if u.strip().startswith('BUNGA'): return 'BUNGA',''
    if 'SETORAN TUNAI' in u: return 'SETORAN TUNAI','1'
    if not is_credit: return 'PELUNASAN HUTANG DAGANG',''
    return 'PENERIMAAN PENJUALAN','5'

def parse_bca_pdf(file_stream):
    with pdfplumber.open(file_stream) as pdf:
        all_lines = []
        for page in pdf.pages:
            text = page.extract_text()
            if text: all_lines.extend(text.split('\n'))

    period = ""
    saldo_awal = 0
    for l in all_lines:
        if 'PERIODE :' in l: period = l.split(':',1)[1].strip()
        m = re.search(r'SALDO AWAL\s*:\s*([\d,]+\.\d+)', l)
        if m: saldo_awal = float(m.group(1).replace(',',''))

    year = period.split()[-1] if period else "2025"
    DATE_RE = re.compile(r'^(\d{2})/(\d{2})\s+(.+)')
    txs = []

    for l in all_lines:
        l = l.strip()
        dm = DATE_RE.match(l)
        if not dm: continue
        
        day, mon, rest = dm.group(1), dm.group(2), dm.group(3).strip()
        if 'SALDO AWAL' in rest or 'Bersambung' in rest: continue
        
        amounts = re.findall(r'([\d,]+\.\d{2})', rest)
        if not amounts: continue
        
        amt = float(amounts[0].replace(',',''))
        is_cr = any(x in rest.upper() for x in [' CR ', 'SETORAN TUNAI', 'KR OTOMATIS'])
        is_db = ' DB ' in rest.upper() or 'BYR VIA' in rest.upper()
        
        if 'BIAYA ADM' in rest.upper(): is_db, is_cr = True, False
        
        ket, kode = classify(rest, is_cr and not is_db)
        txs.append({
            'date': f'{year}-{mon}-{day}',
            'ket': ket,
            'kode': kode,
            'debet': amt if is_db else 0,
            'kredit': amt if (is_cr and not is_db) else (amt if not is_db else 0)
        })
    
    return txs, saldo_awal