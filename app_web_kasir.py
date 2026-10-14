import os
from datetime import datetime
import streamlit as st
import openpyxl
from openpyxl import Workbook
import qrcode

from reportlab.lib.pagesizes import A6
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

DB_FILE = "database_kasir.xlsx"

def init_database():
    if not os.path.exists(DB_FILE):
        wb = Workbook()
        ws_produk = wb.active
        ws_produk.title = "Produk"
        ws_produk.append(["Kode Produk", "Nama Produk", "Harga (Rp)", "Stok"])
        
        produk_awal = [
            ["P001", "Kopi Susu Gula Aren", 18000, 50],
            ["P002", "Roti Bakar Cokelat", 15000, 30],
            ["P003", "Teh Manis Dingin", 6000, 100],
            ["P004", "Mie Goreng Spesial", 22000, 25],
            ["P005", "Air Mineral 600ml", 4000, 80],
        ]
        for p in produk_awal:
            ws_produk.append(p)
            
        ws_transaksi = wb.create_sheet(title="Transaksi")
        ws_transaksi.append(["No Nota", "Tanggal", "Kode Produk", "Nama Produk", "Qty", "Harga Satuan", "Total Harga", "Total Belanja", "Tunai", "Kembalian"])
        wb.save(DB_FILE)

def load_produk():
    init_database()
    wb = openpyxl.load_workbook(DB_FILE)
    ws = wb["Produk"]
    produk_dict = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0]:
            produk_dict[str(row[0])] = {
                "nama": row[1],
                "harga": int(row[2]),
                "stok": int(row[3])
            }
    return produk_dict

def update_stok_dan_simpan_transaksi(no_nota, keranjang, total_bayar, tunai, kembalian):
    wb = openpyxl.load_workbook(DB_FILE)
    ws_produk = wb["Produk"]
    ws_trans = wb["Transaksi"]
    tgl_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for item in keranjang:
        kode = item["kode"]
        qty = item["qty"]
        for row in ws_produk.iter_rows(min_row=2):
            if str(row[0].value) == kode:
                row[3].value = int(row[3].value) - qty
                break
                
        ws_trans.append([
            no_nota,
            tgl_str,
            kode,
            item["nama"],
            qty,
            item["harga"],
            item["subtotal"],
            total_bayar,
            tunai,
            kembalian
        ])
    wb.save(DB_FILE)

def cari_nota_di_db(no_nota_cari):
    if not os.path.exists(DB_FILE):
        return None
    wb = openpyxl.load_workbook(DB_FILE)
    if "Transaksi" not in wb.sheetnames:
        return None
    ws_trans = wb["Transaksi"]
    items = []
    header_info = {}
    for row in ws_trans.iter_rows(min_row=2, values_only=True):
        if row[0] and str(row[0]).strip().upper() == str(no_nota_cari).strip().upper():
            if not header_info:
                header_info = {
                    "no_nota": row[0],
                    "tanggal": row[1],
                    "total_belanja": row[7],
                    "tunai": row[8],
                    "kembalian": row[9]
                }
            items.append({
                "kode": row[2],
                "nama": row[3],
                "qty": row[4],
                "harga": row[5],
                "subtotal": row[6]
            })
    if items:
        return {"header": header_info, "items": items}
    return None

def generate_qr_code(text):
    qr = qrcode.QRCode(version=1, error_correction=qrcode.constants.ERROR_CORRECT_L, box_size=4, border=1)
    qr.add_data(text)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    filename = f"qr_{text}.png"
    img.save(filename)
    return filename

def buat_nota_pdf(no_nota, keranjang, total_bayar, tunai, kembalian, tgl_str=None):
    filename = f"Nota_{no_nota}.pdf"
    doc = SimpleDocTemplate(filename, pagesize=A6, rightMargin=10, leftMargin=10, topMargin=10, bottomMargin=10)
    styles = getSampleStyleSheet()
    style_center = ParagraphStyle(name='Center', parent=styles['Normal'], alignment=1, fontSize=9)
    style_title = ParagraphStyle(name='TitleStyle', parent=styles['Heading2'], alignment=1, fontSize=12, leading=14)
    style_text = ParagraphStyle(name='TextStyle', parent=styles['Normal'], fontSize=8, leading=10)
    
    elements = []
    elements.append(Paragraph("<b>TOKO KASIR DIGITAL WEB</b>", style_title))
    elements.append(Paragraph("Jl. Merdeka No. 123, Indonesia", style_center))
    elements.append(Paragraph("Telp: 0812-3456-7890", style_center))
    elements.append(Spacer(1, 6))
    elements.append(Paragraph("-" * 45, style_center))
    
    if not tgl_str:
        tgl_str = datetime.now().strftime("%d/%m/%Y %H:%M")
    elements.append(Paragraph(f"<b>Kode Nota : {no_nota}</b>", style_text))
    elements.append(Paragraph(f"Tanggal   : {tgl_str}", style_text))
    elements.append(Spacer(1, 5))
    
    table_data = [["Item", "Qty", "Harga", "Total"]]
    for item in keranjang:
        table_data.append([item["nama"], str(item["qty"]), f"{item['harga']:,}", f"{item['subtotal']:,}"])
        
    t = Table(table_data, colWidths=[95, 25, 45, 55])
    t.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
        ('FONTSIZE', (0,0), (-1,-1), 7),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
        ('TOPPADDING', (0,0), (-1,-1), 2),
        ('LINEBELOW', (0,0), (-1,0), 0.5, colors.black),
        ('ALIGN', (1,0), (-1,-1), 'RIGHT'),
    ]))
    elements.append(t)
    elements.append(Spacer(1, 5))
    elements.append(Paragraph("-" * 45, style_center))
    
    pembayaran_data = [
        ["Total Belanja:", f"Rp {total_bayar:,}"],
        ["Tunai:", f"Rp {tunai:,}"],
        ["Kembalian:", f"Rp {kembalian:,}"]
    ]
    t_pay = Table(pembayaran_data, colWidths=[120, 100])
    t_pay.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('ALIGN', (1,0), (1,-1), 'RIGHT'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
    ]))
    elements.append(t_pay)
    elements.append(Spacer(1, 8))
    
    qr_filename = generate_qr_code(no_nota)
    elements.append(RLImage(qr_filename, width=50, height=50))
    elements.append(Paragraph(f"<font size=6>Scan QR / Gunakan Kode <b>{no_nota}</b> untuk cek transaksi</font>", style_center))
    elements.append(Spacer(1, 5))
    elements.append(Paragraph("Terima Kasih Atas Kunjungan Anda!", style_center))
    
    doc.build(elements)
    if os.path.exists(qr_filename):
        os.remove(qr_filename)
    return filename

def main():
    st.set_page_config(page_title="Kasir Web & Nota Digital", page_icon="🛒", layout="centered")
    st.title("🛒 Kasir Digital Web")
    st.caption("Aplikasi Kasir Praktis dengan Database Excel & Nota Digital QR Code")
    
    if "keranjang" not in st.session_state:
        st.session_state.keranjang = []
        
    menu = st.sidebar.radio("Navigasi Menu", ["🛍️ Kasir / Transaksi Baru", "🔍 Cari & Lacak Nota", "📊 Stok & Database Produk"])
    
    if menu == "🛍️ Kasir / Transaksi Baru":
        st.subheader("Kasir Penjualan")
        produk_dict = load_produk()
        col1, col2 = st.columns([2, 1])
        
        with col1:
            pilihan_produk = st.selectbox(
                "Pilih Produk",
                options=list(produk_dict.keys()),
                format_func=lambda x: f"{x} - {produk_dict[x]['nama']} (Rp {produk_dict[x]['harga']:,} | Stok: {produk_dict[x]['stok']})"
            )
            qty = st.number_input("Jumlah (Qty)", min_value=1, value=1, step=1)
            
            if st.button("➕ Tambahkan ke Keranjang"):
                info = produk_dict[pilihan_produk]
                qty_di_keranjang = sum(i["qty"] for i in st.session_state.keranjang if i["kode"] == pilihan_produk)
                
                if (qty + qty_di_keranjang) > info["stok"]:
                    st.error(f"Stok tidak mencukupi! Stok tersisa: {info['stok']}")
                else:
                    st.session_state.keranjang.append({
                        "kode": pilihan_produk,
                        "nama": info["nama"],
                        "harga": info["harga"],
                        "qty": qty,
                        "subtotal": qty * info["harga"]
                    })
                    st.success(f"Berhasil menambahkan {info['nama']} x{qty}")
                    
        with col2:
            st.write("### Keranjang")
            if not st.session_state.keranjang:
                st.info("Keranjang masih kosong.")
            else:
                for idx, item in enumerate(st.session_state.keranjang):
                    st.write(f"**{item['nama']}** ({item['qty']}x)")
                    st.write(f"Rp {item['subtotal']:,}")
                    st.divider()
                total_belanja = sum(i["subtotal"] for i in st.session_state.keranjang)
                st.markdown(f"#### **Total: Rp {total_belanja:,}**")
                if st.button("🗑️ Kosongkan Keranjang"):
                    st.session_state.keranjang = []
                    st.experimental_rerun()

        if st.session_state.keranjang:
            st.divider()
            st.subheader("💳 Pembayaran")
            total_belanja = sum(i["subtotal"] for i in st.session_state.keranjang)
            tunai = st.number_input("Nominal Uang Pembayaran (Rp)", min_value=total_belanja, step=5000)
            kembalian = tunai - total_belanja
            st.info(f"Kembalian: **Rp {kembalian:,}**")
            
            if st.button("✅ Selesaikan & Cetak Nota Digital", type="primary"):
                no_nota = f"INV-{datetime.now().strftime('%Y%m%d%H%M%S')}"
                update_stok_dan_simpan_transaksi(no_nota, st.session_state.keranjang, total_belanja, tunai, kembalian)
                pdf_filename = buat_nota_pdf(no_nota, st.session_state.keranjang, total_bayar=total_belanja, tunai=tunai, kembalian=kembalian)
                
                st.success(f"Transaksi Berhasil! **Kode Nota Unik: {no_nota}**")
                with open(pdf_filename, "rb") as f:
                    st.download_button(
                        label="📄 Download Nota Digital (PDF)",
                        data=f,
                        file_name=pdf_filename,
                        mime="application/pdf"
                    )
                st.session_state.keranjang = []

    elif menu == "🔍 Cari & Lacak Nota":
        st.subheader("🔍 Lacak & Cari Nota Kembali")
        kode_input = st.text_input("Masukkan Kode Nota / ID Transaksi (misal: INV-20261011105022):").strip()
        
        if st.button("Cari Nota"):
            if kode_input:
                hasil = cari_nota_di_db(kode_input)
                if hasil:
                    hdr = hasil["header"]
                    st.success(f"Nota **{hdr['no_nota']}** Ditemukan!")
                    st.write(f"**Tanggal Transaksi:** {hdr['tanggal']}")
                    st.write("#### Detail Barang:")
                    st.table(hasil["items"])
                    st.write(f"**Total Belanja:** Rp {hdr['total_belanja']:,}")
                    st.write(f"**Tunai:** Rp {hdr['tunai']:,}")
                    st.write(f"**Kembalian:** Rp {hdr['kembalian']:,}")
                    
                    pdf_ulang = buat_nota_pdf(hdr['no_nota'], hasil["items"], hdr['total_belanja'], hdr['tunai'], hdr['kembalian'], str(hdr['tanggal']))
                    with open(pdf_ulang, "rb") as f:
                        st.download_button(
                            label="📄 Cetak / Download Ulang Nota PDF",
                            data=f,
                            file_name=pdf_ulang,
                            mime="application/pdf"
                        )
                else:
                    st.error(f"Nota dengan Kode '{kode_input}' tidak ditemukan di database.")

    elif menu == "📊 Stok & Database Produk":
        st.subheader("📊 Database Produk & Stok")
        produk_dict = load_produk()
        table_data = [{"Kode Produk": k, "Nama Produk": v["nama"], "Harga (Rp)": f"Rp {v['harga']:,}", "Stok Tersisa": v["stok"]} for k, v in produk_dict.items()]
        st.dataframe(table_data, use_container_width=True)

if __name__ == "__main__":
    main()
