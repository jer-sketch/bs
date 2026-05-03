import re

def parse_bri(all_lines, classify_func):
    period = ''; saldo_awal = saldo_akhir = total_db = total_cr = 0
    no_rek = ''; nama_akun = ''

    for l in all_lines:
        m_p = re.search(r'(?:Periode Transaksi|Transaction Period)[^\d]*(\d{2}/\d{2}/\d{2})\s*-\s*(\d{2}/\d{2}/\d{2})', l)
        if m_p and not period: period = f"{m_p.group(1)} - {m_p.group(2)}"
        m_n = re.search(r'No\.\s*Rekening[^\d]*([\d]+)', l)
        if m_n and not no_rek: no_rek = m_n.group(1)
        if (l.strip().startswith('CV ') or l.strip().startswith('PT ')) and not nama_akun: nama_akun = l.strip()
        m_s = re.match(r'^([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})$', l.strip())
        if m_s: saldo_awal, total_db, total_cr, saldo_akhir = [float(x.replace(',', '')) for x in m_s.groups()]

    TX_RE = re.compile(r'^(\d{2}/\d{2}/\d{2})\s+\d{2}:\d{2}:\d{2}\s+(.+?)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})$')
    txs = []
    for l in all_lines:
        m = TX_RE.match(l.strip())
        if not m: continue
        date, desc, db_s, cr_s, _ = m.groups()
        db, cr = float(db_s.replace(',', '')), float(cr_s.replace(',', ''))
        ket, kode = classify_func(desc, cr > 0)
        txs.append({'date': date, 'nama': '', 'ket': ket, 'kode': kode, 'debet': db, 'kredit': cr, 'penjelasan': desc})

    return {'bank': 'BRI', 'no_rek': no_rek, 'nama_akun': nama_akun or 'NASABAH BRI', 'period': period, 'saldo_awal': saldo_awal, 'saldo_akhir': saldo_akhir, 'mut_cr': total_cr, 'mut_db': total_db, 'txs': txs}