def parse_bca_pdf_robust(pdf_stream):
    data = []
    saldo_awal = 0
    tahun = str(datetime.now().year)
    
    reader = PdfReader(pdf_stream)
    
    # Gabungkan semua teks dari semua halaman untuk pencarian tahun & saldo awal
    full_text = ""
    for page in reader.pages:
        full_text += page.extract_text() + "\n"

    # Cari Tahun di seluruh dokumen
    match_tahun = re.search(r'PERIODE\s*:\s*.*\s+(20\d{2})', full_text, re.IGNORECASE)
    if match_tahun:
        tahun = match_tahun.group(1)

    # Cari Saldo Awal (biasanya di tabel ringkasan atau awal mutasi)
    match_saldo_awal = re.search(r'SALDO AWAL\s+([\d\.,]+)', full_text, re.IGNORECASE)
    if match_saldo_awal:
        saldo_awal = float(match_saldo_awal.group(1).replace('.', '').replace(',', '.'))

    # Proses per baris untuk mencari transaksi
    lines = full_text.split('\n')
    for line in lines:
        line = line.strip()
        
        # Regex baru yang lebih fleksibel:
        # 1. Mencari Tanggal (DD/MM) di awal
        # 2. Mencari Saldo Akhir di paling ujung (angka dengan koma desimal)
        # 3. Mencari Nominal Mutasi sebelum Saldo Akhir
        # Pola: TGL [KETERANGAN...] [NOMINAL] [SALDO]
        match_trx = re.search(r'^(\d{2}/\d{2})\s+(.*?)\s+([\d\.,]+)\s+([\d\.,]+)$', line)
        
        if match_trx:
            tgl_short = match_trx.group(1)
            keterangan = match_trx.group(2).strip()
            mutasi_raw = match_trx.group(3)
            saldo_raw = match_trx.group(4)
            
            try:
                # Konversi format IDN (1.000,00) ke float (1000.00)
                mutasi = float(mutasi_raw.replace('.', '').replace(',', '.'))
                saldo = float(saldo_raw.replace('.', '').replace(',', '.'))
                
                # Logika penentuan Kredit (CR) atau Debet
                # Di BCA PDF, biasanya ada teks 'CR' di dalam keterangan atau nominal mutasi
                is_kredit = "CR" in keterangan or "CR" in mutasi_raw
                kredit = mutasi if is_kredit else 0
                debet = mutasi if not is_kredit else 0

                data.append({
                    "tanggal": f"{tgl_short}/{tahun}",
                    "keterangan": keterangan.replace(" CR", "").replace(" DB", ""),
                    "debet": debet,
                    "kredit": kredit,
                    "saldo": saldo
                })
            except:
                continue

    return {"saldo_awal": saldo_awal, "tahun": tahun, "trx": data}

def create_excel_template(data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Mutasi"

    # Identitas Perusahaan
    ws['A1'] = "BCA 346-8383111"
    ws['A2'] = "CV. MITRA JAYA ANUGERAH"
    ws['A3'] = f"TAHUN {data['tahun']}"
    
    # Header Tabel
    headers = ["NO", "TANGGAL", "NAMA", "KETERANGAN", "KODE", "DEBET", "KREDIT", "SALDO"]
    ws.append([]) # Baris 4 kosong
    ws.append(headers) # Baris 5
    
    # Styling Header
    for col in range(1, 9):
        cell = ws.cell(row=5, column=col)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')

    # Baris Saldo Awal
    ws.append(["", "", "", "SALDO AWAL", "", "", "", data['saldo_awal']])
    
    # Isi Data Transaksi
    for idx, row in enumerate(data['trx'], start=1):
        ws.append([
            idx,
            row['tanggal'],
            "", # Kolom NAMA sengaja kosong sesuai template
            row['keterangan'],
            5 if row['kredit'] > 0 else "", # KODE 5 jika ada uang masuk
            row['debet'] if row['debet'] > 0 else 0,
            row['kredit'] if row['kredit'] > 0 else 0,
            row['saldo']
        ])

    # Simpan ke stream
    excel_io = io.BytesIO()
    wb.save(excel_io)
    excel_io.seek(0)
    return excel_io