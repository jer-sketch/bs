import streamlit as st
import pdfplumber
import re
import pandas as pd
import io
import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

# --- KONFIGURASI HALAMAN ---
st.set_page_config(page_title="BCA e-Statement Converter", page_icon="🏦")

# --- FUNGSI LOGIKA (COPY DARI SKRIP SEBELUMNYA) ---
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

def process_bca_pdf(uploaded_file):
    all_lines = []
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                all_lines.extend(text.split('\n'))

    period = ''
    saldo_awal = 0
    # Cari info periode dan saldo awal
    for l in all_lines:
        if 'PERIODE :' in l and not period:
            period = l.split(':',1)[1].strip()
        m = re.search(r'SALDO AWAL\s*:\s*([\d,]+\.\d+)', l)
        if m: saldo_awal = float(m.group(1).replace(',',''))

    # Regex untuk tanggal dan jumlah
    DATE_RE = re.compile(r'^(\d{2})/(\d{2})\s+(.+)')
    header_starts = ('REKENING GIRO','KCU ','MITRA JAYA','BOJONGLOA','SITUSAEUR',
                     'JL LEUWI','BANDUNG','INDONESIA','NO. REKENING','HALAMAN',
                     'PERIODE','MATA UANG','CATATAN','TANGGAL KETERANGAN')

    txs = []
    for l in all_lines:
        l = l.strip()
        dm = DATE_RE.match(l)
        if not dm: continue
        
        day, mon, rest = dm.group(1), dm.group(2), dm.group(3).strip()
        if any(rest.startswith(h) for h in header_starts) or 'SALDO AWAL' in rest: continue

        all_amounts = re.findall(r'([\d,]+\.\d{2})', rest)
        if not all_amounts: continue

        amt = float(all_amounts[0].replace(',',''))
        u = rest.upper()
        
        # Logika Debet/Kredit
        is_cr = (' CR ' in u or 'SETORAN TUNAI' in u or 'KR OTOMATIS' in u or 'SWITCHING CR' in u)
        is_db = (' DB ' in u or 'BYR VIA' in u)

        if 'BIAYA ADM' in u or 'PAJAK BUNGA' in u: is_db, is_cr = True, False
        if u.strip().startswith('BUNGA'): is_db, is_cr = False, True

        ket, kode = classify(rest, is_cr and not is_db)
        debet = amt if is_db else 0
        kredit = amt if (is_cr and not is_db) else 0
        if debet == 0 and kredit == 0: kredit = amt # Default
        
        txs.append({
            'Tanggal': f"2025-{mon}-{day}",
            'Keterangan': ket,
            'Kode': kode,
            'Debet': debet,
            'Kredit': kredit
        })
    
    return txs, saldo_awal, period

# --- UI STREAMLIT ---
st.title("🏦 BCA e-Statement to Excel")
st.markdown("Aplikasi khusus untuk mengonversi **PDF Mutasi BCA** ke format Excel Akuntansi.")

file_pdf = st.file_uploader("Upload PDF Mutasi", type=["pdf"])

if file_pdf:
    with st.spinner("Menganalisis data transaksi..."):
        try:
            data_txs, saldo_awal, periode = process_bca_pdf(file_pdf)
            
            if data_txs:
                df = pd.DataFrame(data_txs)
                
                # Hitung Saldo Berjalan
                df['Saldo'] = 0.0
                current_saldo = saldo_awal
                for i in range(len(df)):
                    current_saldo = current_saldo + df.loc[i, 'Kredit'] - df.loc[i, 'Debet']
                    df.loc[i, 'Saldo'] = current_saldo

                st.success(f"Berhasil mengekstrak {len(df)} transaksi periode {periode}")
                
                # Preview
                st.dataframe(df.style.format({"Debet": "{:,.2f}", "Kredit": "{:,.2f}", "Saldo": "{:,.2f}"}))

                # Tombol Download Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='BCA_Januari')
                
                st.download_button(
                    label="📥 Download File Excel",
                    data=output.getvalue(),
                    file_name=f"BCA_{periode.replace(' ','_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("Tidak ada transaksi yang terdeteksi. Pastikan file adalah e-Statement asli BCA.")
        except Exception as e:
            st.error(f"Terjadi kesalahan: {e}")

st.divider()
st.caption("Catatan: Aplikasi ini dirancang khusus untuk format e-Statement BCA KCU Soekarno Hatta.")