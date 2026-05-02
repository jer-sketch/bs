from flask import Flask, request, send_file, render_template_string
import pdfplumber
import re
import io
import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
from datetime import datetime

app = Flask(__name__)

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
    if 'PENERIMAAN NEGARA' in u: return 'PENERIMAAN NEGARA',''
    if not is_credit: return 'PELUNASAN HUTANG DAGANG',''
    return 'PENERIMAAN PENJUALAN','5'

def parse_bca_pdf_logic(pdf_stream):
    all_lines = []
    # Gunakan pdfplumber untuk membaca stream memory
    with pdfplumber.open(pdf_stream) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                all_lines.extend(text.split('\n'))

    period = ''
    saldo_awal = saldo_akhir = mut_cr = mut_db = 0
    
    # Extract Metadata
    for l in all_lines:
        if 'PERIODE :' in l and not period:
            period = l.split(':', 1)[1].strip()
        m = re.search(r'SALDO AWAL\s*:\s*([\d,]+\.\d+)', l)
        if m: saldo_awal = float(m.group(1).replace(',',''))
        m = re.search(r'MUTASI CR\s*:\s*([\d,]+\.\d+)', l)
        if m: mut_cr = float(m.group(1).replace(',',''))
        m = re.search(r'MUTASI DB\s*:\s*([\d,]+\.\d+)', l)
        if m: mut_db = float(m.group(1).replace(',',''))
        m = re.search(r'SALDO AKHIR\s*:\s*([\d,]+\.\d+)', l)
        if m: saldo_akhir = float(m.group(1).replace(',',''))

    parts = period.strip().split()
    year = parts[1] if len(parts) > 1 else str(datetime.now().year)

    DATE_RE = re.compile(r'^(\d{2})/(\d{2})\s+(.+)')
    header_starts = ('REKENING GIRO','KCU ','MITRA JAYA','BOJONGLOA','SITUSAEUR',
                     'JL LEUWI','BANDUNG','INDONESIA','NO. REKENING','HALAMAN',
                     'PERIODE','MATA UANG','CATATAN','Apabila','Rekening ini',
                     'telah menyetujui','BCA berhak','Laporan Mutasi',
                     'TANGGAL KETERANGAN','SALDO AWAL :','MUTASI CR','MUTASI DB','SALDO AKHIR')

    txs = []
    for l in all_lines:
        l = l.strip()
        dm = DATE_RE.match(l)
        if not dm: continue
        
        day, mon, rest = dm.group(1), dm.group(2), dm.group(3).strip()
        if 'SALDO AWAL' in rest or 'Bersambung' in rest: continue
        if any(rest.startswith(h) for h in header_starts): continue

        date = f'{day}/{mon}/{year}'
        all_amounts = re.findall(r'([\d,]+\.\d{2})', rest)
        if not all_amounts: continue

        amt = float(all_amounts[0].replace(',',''))
        u = rest.upper()
        
        is_cr = (' CR ' in u or 'SETORAN TUNAI' in u or 'KR OTOMATIS' in u or 'SWITCHING CR' in u)
        is_db = (' DB ' in u or 'BYR VIA' in u)

        if any(x in u for x in ['BIAYA ADM', 'PAJAK BUNGA', 'PAJAK JASA GIRO']):
            is_db, is_cr = True, False
        if u.strip().startswith('BUNGA'):
            is_db, is_cr = False, True

        ket, kode = classify(rest, is_cr and not is_db)
        debet = amt if is_db else 0
        kredit = amt if (is_cr and not is_db) else 0
        if debet == 0 and kredit == 0: kredit = amt

        txs.append({'date':date,'ket':ket,'kode':kode,'debet':debet,'kredit':kredit})

    return {
        "txs": txs, "period": period, "saldo_awal": saldo_awal, 
        "saldo_akhir": saldo_akhir, "mut_db": mut_db, "mut_cr": mut_cr
    }

def create_excel_output(data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'bca'
    
    # Styles
    hdr_font = Font(name='Arial', bold=True, size=10)
    data_font = Font(name='Arial', size=10)
    num_fmt = '#,##0.00'

    # Header Section
    ws.append(['BCA 346-8383111'])
    ws.append(['CV. MITRA JAYA ANUGERAH'])
    ws.append([f"PERIODE: {data['period']}"])
    ws.append(['+','-','-','-','-','-','-','-','-','+'])
    ws.append(['|','NO','TANGGAL','NAMA','KETERANGAN','KODE','DEBET','KREDIT','SALDO','|'])
    ws.append(['+','-','-','-','-','-','-','-','-','+'])
    
    # Saldo Awal Row
    ws.append(['|','','','','SALDO AWAL','','','', data['saldo_awal'], '|'])
    
    current_saldo = data['saldo_awal']
    for idx, tx in enumerate(data['txs'], 1):
        if tx['kredit'] > 0:
            current_saldo = round(current_saldo + tx['kredit'], 2)
        else:
            current_saldo = round(current_saldo - tx['debet'], 2)
            
        ws.append(['|', idx, tx['date'], '', tx['ket'], tx['kode'],
                   tx['debet'] if tx['debet'] > 0 else '', 
                   tx['kredit'] if tx['kredit'] > 0 else '',
                   current_saldo, '|'])

    # Footer
    ws.append(['+','-','-','-','-','-','-','-','-','+'])
    ws.append(['|','','','','TOTAL','', data['mut_db'], data['mut_cr'], data['saldo_akhir'], '|'])
    
    # Formatting
    for row in ws.iter_rows(min_row=5):
        for cell in row:
            cell.font = data_font
            if cell.column in [7, 8, 9] and isinstance(cell.value, (int, float)):
                cell.number_format = num_fmt

    # Set Column Widths
    widths = [3, 7, 14, 25, 30, 6, 16, 16, 18, 3]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    excel_io = io.BytesIO()
    wb.save(excel_io)
    excel_io.seek(0)
    return excel_io

@app.route('/')
def index():
    return render_template_string("""
    <html>
        <body style="font-family:sans-serif; text-align:center; padding-top:50px;">
            <h2>BCA Statement to Excel (Januari Logic)</h2>
            <form action="/convert" method="post" enctype="multipart/form-data">
                <input type="file" name="file" accept=".pdf" required><br><br>
                <button type="submit" style="padding:10px 20px;">Upload & Konversi</button>
            </form>
        </body>
    </html>
    """)

@app.route('/convert', methods=['POST'])
def convert():
    file = request.files['file']
    if not file: return "No file"
    
    pdf_stream = io.BytesIO(file.read())
    data = parse_bca_pdf_logic(pdf_stream)
    excel_file = create_excel_output(data)
    
    return send_file(
        excel_file,
        as_attachment=True,
        download_name=f"MUTASI_BCA_{datetime.now().strftime('%Y%m%d')}.xlsx"
    )

if __name__ == '__main__':
    app.run(debug=True)