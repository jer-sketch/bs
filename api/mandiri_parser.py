import re

_MONTHS = {'Jan':'01','Feb':'02','Mar':'03','Apr':'04','May':'05','Jun':'06','Jul':'07','Aug':'08','Sep':'09','Oct':'10','Nov':'11','Dec':'12'}

def parse_mandiri(all_lines, classify_func):
    period = ''; saldo_awal = saldo_akhir = total_db = total_cr = 0
    no_rek = ''; nama_akun = ''
    for l in all_lines:
        m_p = re.search(r'(\d{2}\s+\w+\s+\d{4})\s*-\s*(\d{2}\s+\w+\s+\d{4})', l)
        if m_p and not period: period = f"{m_p.group(1)} - {m_p.group(2)}"
        m_n = re.match(r'^(\d{10,16})\s+(.+)', l.strip())
        if m_n and not no_rek: no_rek, nama_akun = m_n.group(1), m_n.group(2).split('  ')[0].strip()
        m_s = re.match(r'^([\d,]+\.\d{2})\s+\d+\s+([\d,]+\.\d{2})$', l.strip())
        if m_s:
            v1, v2 = float(m_s.group(1).replace(',', '')), float(m_s.group(2).replace(',', ''))
            if saldo_awal == 0: saldo_awal, total_db = v1, v2
            else: saldo_akhir, total_cr = v1, v2

    TX_AMT_RE = re.compile(r'^(.*?)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s*$')
    DATE_RE = re.compile(r'(\d{2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{4})', re.IGNORECASE)
    txs, last_date = [], ''
    for i, l in enumerate(all_lines):
        dm = DATE_RE.search(l)
        if dm: last_date = f"{dm.group(1)}/{_MONTHS.get(dm.group(2)[:3].capitalize(), '01')}/{dm.group(3)[-2:]}"
        m = TX_AMT_RE.match(l.strip())
        if not m or not last_date: continue
        desc = re.sub(r'\d{2}:\d{2}:\d{2}', '', m.group(1)).strip()
        db, cr = float(m.group(2).replace(',', '')), float(m.group(3).replace(',', ''))
        if any(desc.lower().startswith(s) for s in ('closing', 'opening', 'total')): continue
        ket, kode = classify_func(desc, cr > 0)
        txs.append({'date': last_date, 'nama': '', 'ket': ket, 'kode': kode, 'debet': db, 'kredit': cr, 'penjelasan': desc})
    return {'bank': 'MANDIRI', 'no_rek': no_rek, 'nama_akun': nama_akun or 'NASABAH MANDIRI', 'period': period, 'saldo_awal': saldo_awal, 'saldo_akhir': saldo_akhir, 'mut_cr': total_cr, 'mut_db': total_db, 'txs': txs}