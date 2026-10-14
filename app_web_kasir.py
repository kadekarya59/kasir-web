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

# ---------------------------------------------------------
# DATABASE FUNCTIONS
# ---------------------------------------------------------
def init_database():
    """Membuat database spreadsheet jika belum ada"""
    if not os.path.exists(DB_FILE):
        wb = Workbook()
        
        # Sheet Produk
        ws_produk = wb.active
        ws_produk.title = "Produk"
        ws_produk.append(["Kode Produk", "Nama Produk", "Harga (Rp)", "Stok", "Kategori", "Icon"])
        
        produk_awal = [
            ["P001", "Kopi Susu Gula Aren", 18000, 50, "Minuman", "☕"],
            ["P002", "Roti Bakar Cokelat", 15000, 30, "Makanan", "🍞"],
            ["P003", "Teh Manis Dingin", 6000, 100, "Minuman", "🧃"],
            ["P004", "Mie Goreng Spesial", 22000, 25, "Makanan", "🍜"],
            ["P005", "Air Mineral 600ml", 4000, 80, "Minuman", "🍼"],
        ]
        for p in produk_awal:
            ws_produk.append(p)
            
        # Sheet Transaksi
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
                "stok": int(row[3]),
                "kategori": row[4] if len(row) > 4 and row[4] else "Umum",
                "icon": row[5] if len(row) > 5 and row[5] else "📦"
            }
    return produk_dict

def tambah_atau_update_produk_db(kode, nama, harga, stok, kategori, icon):
    """Menambah produk baru atau mengubah data produk yang sudah ada di Excel"""
    wb = openpyxl.load_workbook(DB_FILE)
    ws = wb["Produk"]
    found = False
    
    for row in ws.iter_rows(min_row=2):
        if str(row[0].value) == kode:
            row[1].value = nama
            row[2].value = harga
            row[3].value = stok
            row[4].value = kategori
            row[5].value = icon
            found = True
            break
            
    if not found:
        ws.append([kode, nama, harga, stok, kategori, icon])
        
    wb.save(DB_FILE)

def hapus_produk_db(kode):
    """Menghapus produk dari Excel berdasarkan kode"""
    wb = openpyxl.load_workbook(DB_FILE)
    ws = wb["Produk"]
    
    for row in ws.iter_rows(min_row=2):
        if str(row[0].value) == kode:
            ws.delete_rows(row[0].row)
            break
            
    wb.save(DB_FILE)

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

# ---------------------------------------------------------
# GENERATE PDF & QR CODE
# ---------------------------------------------------------
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
    elements.append(Paragraph("<b>MINIMARKET MART DIGITAL</b>", style_title))
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

# ---------------------------------------------------------
# SISTEM AUTENTIKASI LOGIN
# ---------------------------------------------------------
USERS = {
    "admin": {"password": "semogasukses", "role": "Admin", "nama": "Manajer / Admin"},
    "kasir": {"password": "hidupkasir", "role": "Kasir", "nama": "Petugas Kasir"}
}

def login_page():
    st.title("🔑 Login Kasir POS Minimarket")
    st.caption("Masukkan Username dan Password Anda")
    
    col1, col2 = st.columns([1, 1])
    with col1:
        username_input = st.text_input("Username").strip().lower()
        password_input = st.text_input("Password", type="password")
        
        if st.button("Masuk / Login", type="primary"):
            if username_input in USERS and USERS[username_input]["password"] == password_input:
                st.session_state.logged_in = True
                st.session_state.user_info = USERS[username_input]
                st.success(f"Selamat datang, {USERS[username_input]['nama']}!")
                st.rerun()
            else:
                st.error("Username atau Password salah!")
    

# ---------------------------------------------------------
# MAIN APP INTERFACE
# ---------------------------------------------------------
def main():
    st.set_page_config(page_title="POS Minimarket Kasir", page_icon="🏪", layout="wide")
    
    # Custom CSS
    st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .product-card {
        background-color: #ffffff;
        border-radius: 12px;
        padding: 12px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.05);
        border: 1px solid #e9ecef;
        text-align: center;
        margin-bottom: 12px;
    }
    .product-icon { font-size: 38px; margin-bottom: 5px; }
    .product-title { font-weight: bold; font-size: 14px; color: #212529; height: 38px; }
    .product-price { color: #2b8a3e; font-weight: bold; font-size: 15px; margin-bottom: 5px; }
    .total-banner {
        background-color: #1c7ed6;
        color: white;
        padding: 15px;
        border-radius: 8px;
        text-align: center;
        font-size: 22px;
        font-weight: bold;
        margin-top: 10px;
    }
    </style>
    """, unsafe_allow_html=True)
    
    if "logged_in" not in st.session_state or not st.session_state.logged_in:
        login_page()
        return

    # User Info Sidebar
    user = st.session_state.user_info
    st.sidebar.markdown(f"### 👤 {user['nama']}")
    st.sidebar.caption(f"Akses: **{user['role']}**")
    
    if st.sidebar.button("🚪 Logout"):
        st.session_state.logged_in = False
        st.rerun()
        
    st.sidebar.divider()

    # Pilihan Menu Sesuai Role
    if user["role"] == "Admin":
        menu_options = ["🛒 Mesin Kasir POS", "🛠️ Kelola Produk & Harga", "🔍 Cari / Lacak Nota", "📊 Stok & Database Produk"]
    else:
        # User Kasir hanya bisa Transaksi dan Cari Nota
        menu_options = ["🛒 Mesin Kasir POS", "🔍 Cari / Lacak Nota"]

    menu = st.sidebar.radio("Navigasi", menu_options)

    if "keranjang" not in st.session_state:
        st.session_state.keranjang = []

    # -----------------------------------------------------
    # MENU 1: MESIN KASIR POS
    # -----------------------------------------------------
    if menu == "🛒 Mesin Kasir POS":
        st.title("🛒 Kasir Minimarket (POS)")
        produk_dict = load_produk()
        col_catalog, col_cart = st.columns([2.2, 1.3])
        
        with col_catalog:
            st.subheader("📦 Katalog Produk")
            col_search, col_filter = st.columns([2, 1])
            with col_search:
                search_query = st.text_input("🔍 Cari Barang / Kode...", "").strip().lower()
            with col_filter:
                kategori_list = ["Semua"] + list(set(v["kategori"] for v in produk_dict.values()))
                kategori_pilihan = st.selectbox("Kategori", kategori_list)
                
            st.divider()
            
            produk_filtered = {}
            for k, v in produk_dict.items():
                match_search = (search_query in k.lower()) or (search_query in v["nama"].lower())
                match_kat = (kategori_pilihan == "Semua") or (v["kategori"] == kategori_pilihan)
                if match_search and match_kat:
                    produk_filtered[k] = v

            grid_cols = st.columns(3)
            idx = 0
            for k, v in produk_filtered.items():
                with grid_cols[idx % 3]:
                    st.markdown(f"""
                    <div class="product-card">
                        <div class="product-icon">{v['icon']}</div>
                        <div class="product-title">{v['nama']}</div>
                        <div class="product-price">Rp {v['harga']:,}</div>
                        <small style="color: #6c757d;">Kode: {k} | Stok: {v['stok']}</small>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    if st.button(f"➕ Tambah", key=f"btn_{k}", use_container_width=True):
                        qty_di_keranjang = sum(i["qty"] for i in st.session_state.keranjang if i["kode"] == k)
                        if qty_di_keranjang + 1 > v["stok"]:
                            st.toast(f"⚠️ Stok {v['nama']} tidak mencukupi!", icon="❌")
                        else:
                            found = False
                            for item in st.session_state.keranjang:
                                if item["kode"] == k:
                                    item["qty"] += 1
                                    item["subtotal"] = item["qty"] * item["harga"]
                                    found = True
                                    break
                            if not found:
                                st.session_state.keranjang.append({
                                    "kode": k,
                                    "nama": v["nama"],
                                    "harga": v["harga"],
                                    "qty": 1,
                                    "subtotal": v["harga"]
                                })
                            st.toast(f"✅ {v['nama']} ditambahkan!", icon="🛒")
                            st.rerun()
                idx += 1

        with col_cart:
            st.subheader("📋 Struk Kasir")
            if not st.session_state.keranjang:
                st.info("Keranjang kosong. Klik produk di sebelah kiri untuk menambahkan.")
            else:
                total_belanja = sum(i["subtotal"] for i in st.session_state.keranjang)
                st.markdown(f'<div class="total-banner">TOTAL: Rp {total_belanja:,}</div>', unsafe_allow_html=True)
                st.write("")
                
                for idx, item in enumerate(st.session_state.keranjang):
                    c_nama, c_qty, c_sub = st.columns([2, 1, 1])
                    with c_nama:
                        st.write(f"**{item['nama']}**")
                        st.caption(f"@ Rp {item['harga']:,}")
                    with c_qty:
                        st.write(f"x {item['qty']}")
                    with c_sub:
                        st.write(f"**Rp {item['subtotal']:,}**")
                    st.divider()

                if st.button("🗑️ Batal All", use_container_width=True):
                    st.session_state.keranjang = []
                    st.rerun()

                st.subheader("💳 Pembayaran")
                tunai = st.number_input("Bayar / Tunai (Rp)", min_value=total_belanja, value=float(total_belanja), step=5000.0)
                kembalian = tunai - total_belanja
                st.success(f"Kembalian: **Rp {int(kembalian):,}**")
                
                if st.button("🔴 CETAK STRUK / BAYAR", type="primary", use_container_width=True):
                    no_nota = f"INV-{datetime.now().strftime('%Y%m%d%H%M%S')}"
                    update_stok_dan_simpan_transaksi(no_nota, st.session_state.keranjang, total_belanja, int(tunai), int(kembalian))
                    pdf_filename = buat_nota_pdf(no_nota, st.session_state.keranjang, total_belanja, int(tunai), int(kembalian))
                    
                    st.balloons()
                    st.success(f"Transaksi Selesai! Nota: **{no_nota}**")
                    with open(pdf_filename, "rb") as f:
                        st.download_button(
                            label="📄 Download Struk PDF",
                            data=f,
                            file_name=pdf_filename,
                            mime="application/pdf",
                            use_container_width=True
                        )
                    st.session_state.keranjang = []

    # -----------------------------------------------------
    # MENU 2: KELOLA PRODUK & HARGA (KHUSUS ADMIN)
    # -----------------------------------------------------
    elif menu == "🛠️ Kelola Produk & Harga":
        st.title("🛠️ Kelola Produk, Stok & Harga (Admin Only)")
        produk_dict = load_produk()
        
        tab1, tab2 = st.tabs(["➕ Tambah / Edit Produk", "❌ Hapus Produk"])
        
        with tab1:
            st.subheader("Tambah Produk Baru / Edit Produk")
            pilihan_edit = st.selectbox("Pilih Produk yang Ingin Di-edit (Atau pilih 'Baru')", ["-- Buat Baru --"] + list(produk_dict.keys()))
            
            if pilihan_edit != "-- Buat Baru --":
                info = produk_dict[pilihan_edit]
                val_kode = pilihan_edit
                val_nama = info["nama"]
                val_harga = info["harga"]
                val_stok = info["stok"]
                val_kat = info["kategori"]
                val_icon = info["icon"]
            else:
                val_kode = ""
                val_nama = ""
                val_harga = 10000
                val_stok = 10
                val_kat = "Makanan"
                val_icon = "📦"

            c1, c2 = st.columns(2)
            with c1:
                kode_p = st.text_input("Kode Produk (Unik)", value=val_kode)
                nama_p = st.text_input("Nama Produk", value=val_nama)
                harga_p = st.number_input("Harga (Rp)", min_value=100, value=val_harga, step=500)
            with c2:
                stok_p = st.number_input("Jumlah Stok", min_value=0, value=val_stok, step=1)
                kategori_p = st.selectbox("Kategori", ["Makanan", "Minuman", "Snack", "Umum"])
                icon_p = st.text_input("Ikon Emoji", value=val_icon)

            if st.button("💾 Simpan Data Produk", type="primary"):
                if kode_p and nama_p:
                    tambah_atau_update_produk_db(kode_p, nama_p, harga_p, stok_p, kategori_p, icon_p)
                    st.success(f"Produk {nama_p} berhasil disimpan ke Excel!")
                    st.rerun()
                else:
                    st.error("Kode dan Nama Produk wajib diisi!")

        with tab2:
            st.subheader("Hapus Produk dari Spreadsheet")
            kode_hapus = st.selectbox("Pilih Produk yang Akan Dihapus", list(produk_dict.keys()), format_func=lambda x: f"{x} - {produk_dict[x]['nama']}")
            if st.button("🗑️ Hapus Produk Ini", type="primary"):
                hapus_produk_db(kode_hapus)
                st.success(f"Produk {kode_hapus} berhasil dihapus dari Excel!")
                st.rerun()

    # -----------------------------------------------------
    # MENU 3: CARI NOTA
    # -----------------------------------------------------
    elif menu == "🔍 Cari / Lacak Nota":
        st.title("🔍 Lacak Transaksi Minimarket")
        kode_input = st.text_input("Masukkan Kode Nota / ID Transaksi:").strip()
        
        if st.button("Cari Nota"):
            if kode_input:
                hasil = cari_nota_di_db(kode_input)
                if hasil:
                    hdr = hasil["header"]
                    st.success(f"Nota **{hdr['no_nota']}** Ditemukan!")
                    st.write(f"**Tanggal Transaksi:** {hdr['tanggal']}")
                    st.table(hasil["items"])
                    st.write(f"**Total Belanja:** Rp {hdr['total_belanja']:,}")
                    st.write(f"**Tunai:** Rp {hdr['tunai']:,}")
                    st.write(f"**Kembalian:** Rp {hdr['kembalian']:,}")
                    
                    pdf_ulang = buat_nota_pdf(hdr['no_nota'], hasil["items"], hdr['total_belanja'], hdr['tunai'], hdr['kembalian'], str(hdr['tanggal']))
                    with open(pdf_ulang, "rb") as f:
                        st.download_button(
                            label="📄 Download Ulang Struk PDF",
                            data=f,
                            file_name=pdf_ulang,
                            mime="application/pdf"
                        )
                else:
                    st.error(f"Nota dengan Kode '{kode_input}' tidak ditemukan di database.")

    # -----------------------------------------------------
    # MENU 4: STOK PRODUK (ADMIN ONLY)
    # -----------------------------------------------------
    elif menu == "📊 Stok & Database Produk":
        st.title("📊 Database & Stok Minimarket")
        produk_dict = load_produk()
        table_data = [
            {
                "Kode": k,
                "Icon": v["icon"],
                "Nama Produk": v["nama"],
                "Kategori": v["kategori"],
                "Harga (Rp)": f"Rp {v['harga']:,}",
                "Stok": v["stok"]
            } for k, v in produk_dict.items()
        ]
        st.dataframe(table_data, use_container_width=True)

if __name__ == "__main__":
    main()
