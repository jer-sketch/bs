import io
import openpyxl
from openpyxl.styles import Font, Alignment

def generate_excel(txs, saldo_awal):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'bca'

    # Header
    ws.append(['BCA Statement Report'])
    ws.append(['NO', 'TANGGAL', 'KETERANGAN', 'KODE', 'DEBET', 'KREDIT', 'SALDO'])
    
    # Baris Saldo Awal
    ws.append(['', '', 'SALDO AWAL', '', '', '', saldo_awal])
    
    current_saldo = saldo_awal
    for i, t in enumerate(txs, 1):
        current_saldo = round(current_saldo + t['kredit'] - t['debet'], 2)
        ws.append([i, t['date'], t['ket'], t['kode'], t['debet'] or '', t['kredit'] or '', current_saldo])

    # Styling sederhana
    for cell in ws[2]: cell.font = Font(bold=True)
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output