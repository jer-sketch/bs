from flask import Flask, request, send_file, render_template_string
import io
import pdfplumber
import re
import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

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

def parse_pdf(file_stream):
    with pdfplumber.open(file_stream) as pdf:
        all_lines = []
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                all_lines.extend(text.split('\n'))
                
    period = ''
    saldo_awal = saldo_akhir = mut_cr = mut_db = 0
    for l in all_lines:
        if 'PERIODE :' in l and not period:
            period = l.split(':',1)[1].strip()
        m = re.search(r'SALDO AWAL\s*:\s*([\d,]+\.\d+)', l)
        if m: saldo_awal = float(m.group(1).replace(',',''))
        m = re.search(r'MUTASI CR\s*:\s*([\d,]+\.\d+)', l)
        if m: mut_cr = float(m.group(1).replace(',',''))
        m = re.search(r'MUTASI DB\s*:\s*([\d,]+\.\d+)', l)
        if m: mut_db = float(m.group(1).replace(',',''))
        m = re.search(r'SALDO AKHIR\s*:\s*([\d,]+\.\d+)', l)
        if m: saldo_akhir = float(m.group(1).replace(',',''))
        
    parts = period.strip().split()
    month_num = MONTH_MAP.get(parts[0].upper(), 0) if parts else 0
    year = parts[1] if len(parts) > 1 else '2025'
    DATE_RE = re.compile(r'^(\d{2})/(\d{2})\s+(.+)')
    skip_words = {'SALDO AWAL', 'Bersambung'}
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
        if any(s in rest for s in skip_words): continue
        if any(rest.startswith(h) for h in header_starts): continue
        date = f'{year}-{mon}-{day}'
        all_amounts = re.findall(r'([\d,]+\.\d{2})', rest)
        if not all_amounts: continue
        amt = float(all_amounts[0].replace(',',''))
        if amt <= 0: continue
        u = rest.upper()
        is_cr = (' CR ' in u or 'SETORAN TUNAI' in u or 'KR OTOMATIS' in u or 'SWITCHING CR' in u)
        is_db = (' DB ' in u or 'BYR VIA' in u)
        if 'BIAYA ADM' in u or 'PAJAK BUNGA' in u or 'PAJAK JASA GIRO' in u:
            is_db, is_cr = True, False
        if u.strip().startswith('BUNGA'):
            is_db, is_cr = False, True
        ket, kode = classify(rest, is_cr and not is_db)
        debet = amt if is_db else 0
        kredit = amt if (is_cr and not is_db) else 0
        if debet == 0 and kredit == 0:
            kredit = amt
        txs.append({'date':date,'ket':ket,'kode':kode,'debet':debet,'kredit':kredit})
        
    return txs, period, month_num, saldo_awal, saldo_akhir, mut_cr, mut_db, year

HTML_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
    <title>BCA PDF to Excel Converter</title>
    <style>
        body { font-family: Arial, sans-serif; padding: 40px; text-align: center; }
        .container { max-width: 500px; margin: 0 auto; border: 1px solid #ccc; padding: 20px; border-radius: 10px; }
        input[type="file"] { margin-bottom: 20px; }
        button { padding: 10px 20px; background-color: #0066cc; color: white; border: none; border-radius: 5px; cursor: pointer; }
        button:hover { background-color: #0052a3; }
    </style>
</head>
<body>
    <div class="container">
        <h2>Upload BCA E-Statement (PDF)</h2>
        <form action="/convert" method="post" enctype="multipart/form-data">
            <input type="file" name="file" accept=".pdf" required>
            <br>
            <button type="submit">Convert to Excel</button>
        </form>
    </div>
</body>
</html>
"""

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return "No file part in the request", 400
    
    file = request.files['file']
    if file.filename == '':
        return "No selected file", 400

    if file:
        try:
            # 1. Parse PDF from memory
            txs, period, month_num, saldo_awal, saldo_akhir, mut_cr, mut_db, year = parse_pdf(file)
            
            # 2. Build Excel Workbook in memory
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = 'bca'
            
            hdr_font = Font(name='Arial', bold=True, size=10)
            data_font = Font(name='Arial', size=10)
            num_fmt = '#,##0.00'
            
            ws.append(['BCA 346-8383111'])
            ws.append(['CV. MITRA JAYA ANUGERAH'])
            ws.append([f'TAHUN {year}'])
            ws.append(['+','-','-','-','-','-','-','-','-','+'])
            ws.append(['|','NO','TANGGAL','NAMA','KETERANGAN','KODE','DEBET','KREDIT','SALDO','|'])
            ws.append(['+','-','-','-','-','-','-','-','-','+'])
            ws.append(['|','','','','SALDO AWAL','','','',saldo_awal,'|'])
            ws.append(['|','','','','','','','','','|'])
            
            saldo = saldo_awal
            for no, tx in enumerate(txs, 1):
                if tx['kredit'] > 0:
                    saldo = round(saldo + tx['kredit'], 2)
                else:
                    saldo = round(saldo - tx['debet'], 2)
                ws.append(['|', no, tx['date'], '', tx['ket'], tx['kode'],
                           tx['debet'] if tx['debet'] else '', tx['kredit'] if tx['kredit'] else '',
                           saldo, '|'])
                           
            ws.append(['+','-','-','-','-','-','-','-','-','+'])
            ws.append(['|','','','','','',mut_db,mut_cr,saldo_akhir,'|'])
            ws.append(['+','-','-','-','-','-','-','-','-','+'])
            
            # Formatting
            widths = [3, 7, 14, 30, 28, 6, 16, 16, 18, 3]
            for i, w in enumerate(widths, 1):
                ws.column_dimensions[get_column_letter(i)].width = w
                
            for row in ws.iter_rows(min_row=7, max_row=ws.max_row):
                for cell in [row[6], row[7], row[8]]: 
                    if isinstance(cell.value, (int, float)) and cell.value:
                        cell.number_format = num_fmt
                    cell.font = data_font
                for cell in row:
                    cell.font = data_font
                    
            for cell in ws[5]:
                cell.font = hdr_font
            
            # 3. Save to BytesIO stream instead of a file on disk
            excel_io = io.BytesIO()
            wb.save(excel_io)
            excel_io.seek(0)
            
            filename_safe_period = period.replace(" ", "_")
            download_name = f'BANK_{filename_safe_period}.xlsx'
            
            return send_file(
                excel_io,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                as_attachment=True,
                download_name=download_name
            )
            
        except Exception as e:
            return f"An error occurred during processing: {str(e)}", 500

if __name__ == '__main__':
    app.run(debug=True)