def ambil_laporan_harian(tanggal_pilihan):
    """Membaca transaksi dari Excel berdasarkan tanggal tertentu (YYYY-MM-DD)"""
    if not os.path.exists(DB_FILE):
        return [], 0
    
    wb = openpyxl.load_workbook(DB_FILE)
    if "Transaksi" not in wb.sheetnames:
        return [], 0
        
    ws_trans = wb["Transaksi"]
    transaksi_harian = []
    total_omset = 0
    
    tgl_str = tanggal_pilihan.strftime("%Y-%m-%d")
    
    for row in ws_trans.iter_rows(min_row=2, values_only=True):
        if row[1]:  # Kolom Tanggal
            tgl_transaksi = str(row[1]).split(" ")[0]  # Ambil bagian YYYY-MM-DD saja
            if tgl_transaksi == tgl_str:
                transaksi_harian.append({
                    "No Nota": row[0],
                    "Jam": str(row[1]).split(" ")[1] if " " in str(row[1]) else "-",
                    "Kode Barang": row[2],
                    "Nama Produk": row[3],
                    "Qty": row[4],
                    "Harga Satuan": f"Rp {row[5]:,}",
                    "Total Harga": row[6],
                    "Metode Bayar": row[8] if len(row) > 8 else "Cash / Tunai"
                })
                total_omset += int(row[6])
                
    return transaksi_harian, total_omset
