from flask import Flask, request, send_file, jsonify
from .parser import parse_bca_pdf
from .excel_gen import generate_excel

app = Flask(__name__)

@app.route('/api/index', methods=['POST'])
def main_process():
    try:
        file = request.files.get('file')
        if not file:
            return "File tidak ditemukan", 400

        # Langkah 1: Jalankan Parser
        data_transaksi, saldo_awal = parse_bca_pdf(file)

        # Langkah 2: Jalankan Generator Excel
        excel_file = generate_excel(data_transaksi, saldo_awal)

        # Langkah 3: Kirim ke User
        return send_file(
            excel_file,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name="Laporan_Keuangan_BCA.xlsx"
        )
    except Exception as e:
        return f"Terjadi kesalahan: {str(e)}", 500

# Penting untuk Vercel
def handler(event, context):
    return app(event, context)