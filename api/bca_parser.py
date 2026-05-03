import re
from datetime import datetime
import pdfplumber

def classify_bca(full_text, is_credit):
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

def extract_nama_refined_v5(block, amount):
    amt_str = "{:.2f}".format(amount)
    found_idx = -1
    for i, line in enumerate(block):
        if amt_str in line.replace(',', ''):
            found_idx = i
            break
    if found_idx == -1: return "", ""

    candidates = block[found_idx+1:]
    name_list, desc_list = [], []
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
        if stop_processing or any(g in c_up for g in garbage): continue
        if any(c.islower() for c in c_strip): desc_list.append(c_strip)
        elif c_strip.startswith('/'): desc_list.append(c_strip.replace('/', '').strip())
        elif c_strip.isupper() and not re.match(r'^[A-Z]{2}\s\d+$', c_strip):
            name_list.append(c_strip)

    raw_name = " ".join(name_list)
    patterns = [r'\bTRSF\b', r'\bWS\b', r'\bFTSCY\b', r'\bDB\b', r'\bCR\b', r'\bDR\b', r'\bSWITCHING\b']
    for p in patterns: raw_name = re.sub(p, '', raw_name, flags=re.IGNORECASE)
    
    return " ".join(raw_name.split()).strip().upper(), " ".join(desc_list).strip()

def parse_bca(all_lines):
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
    tx_blocks, current_block = [], []
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
        date_str = f"{m.group(1)}/{m.group(2)}/{year[-2:]}"
        amounts = re.findall(r'([\d,]+\.\d{2})', full_text)
        if not amounts: continue
        amt = float(amounts[0].replace(',', ''))
        
        is_cr = any(x in full_text for x in [' CR', 'SETORAN TUNAI', 'KR OTOMATIS', 'BUNGA'])
        is_db = any(x in full_text for x in [' DB', 'BYR VIA', 'BIAYA ADM', 'PAJAK'])
        if 'PAJAK' in full_text: is_db, is_cr = True, False
        if 'BUNGA' in full_text and 'PAJAK' not in full_text: is_db, is_cr = False, True

        ket, kode = classify_bca(full_text, is_cr and not is_db)
        nama_orang, penjelasan = extract_nama_refined_v5(block, amt)

        db, cr = (amt, 0) if is_db else (0, amt if (is_cr and not is_db) else 0)
        total_db += db; total_cr += cr
        
        txs.append({'date': date_str, 'nama': nama_orang, 'ket': ket, 'kode': kode, 'debet': db, 'kredit': cr, 'penjelasan': penjelasan})

    return {'bank': 'BCA', 'no_rek': 'BCA-SET', 'nama_akun': 'NASABAH BCA', 'period': period, 'saldo_awal': saldo_awal, 'saldo_akhir': 0, 'mut_cr': total_cr, 'mut_db': total_db, 'txs': txs}