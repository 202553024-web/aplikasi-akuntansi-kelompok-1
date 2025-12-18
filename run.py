import streamlit as st
import pandas as pd
import altair as alt
from datetime import datetime
import io
import calendar
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# ============================
# CONFIG TAMPAK APLIKASI
# ============================
st.set_page_config(
    page_title="Aplikasi Akuntansi Keuangan",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    /* Main Title */
    .main-title {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 30px;
        border-radius: 15px;
        text-align: center;
        color: white;
        margin-bottom: 30px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.2);
    }
    .main-title h1 {
        font-size: 42px;
        font-weight: 800;
        margin: 0;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
    }
    .main-title p {
        font-size: 18px;
        margin: 10px 0 0 0;
        opacity: 0.9;
    }
    
    /* Subtitle */
    .subtitle {
        background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
        padding: 20px;
        border-radius: 10px;
        color: white;
        font-size: 24px;
        font-weight: 700;
        margin: 20px 0;
        text-align: center;
        box-shadow: 0 3px 10px rgba(0,0,0,0.15);
    }
    
    /* Cards */
    .metric-card {
        background: white;
        padding: 25px;
        border-radius: 12px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        border-left: 5px solid #667eea;
        margin: 10px 0;
    }
    
    /* Buttons */
    .stButton>button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%) !important;
        color: white !important;
        padding: 12px 28px !important;
        border-radius: 8px !important;
        font-size: 16px !important;
        font-weight: 600 !important;
        border: none !important;
        box-shadow: 0 4px 12px rgba(102, 126, 234, 0.4) !important;
        transition: all 0.3s ease !important;
    }
    .stButton>button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 20px rgba(102, 126, 234, 0.6) !important;
    }
    
    /* Sidebar */
    .css-1d391kg, [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #667eea 0%, #764ba2 100%);
    }
    
    /* Form inputs */
    .stTextInput>div>div>input, .stNumberInput>div>div>input, .stSelectbox>div>div>select {
        border-radius: 8px !important;
        border: 2px solid #e0e0e0 !important;
        padding: 10px !important;
    }
    
    /* DataFrames */
    .dataframe {
        border-radius: 10px !important;
        overflow: hidden !important;
    }
    
    /* Info boxes */
    .stAlert {
        border-radius: 10px !important;
        padding: 15px !important;
    }
    
    /* Metrics */
    [data-testid="stMetricValue"] {
        font-size: 28px !important;
        font-weight: 700 !important;
        color: #667eea !important;
    }
    
    /* Success/Error messages */
    .element-container:has(.stSuccess) {
        animation: slideIn 0.5s ease;
    }
    
    @keyframes slideIn {
        from {
            opacity: 0;
            transform: translateY(-10px);
        }
        to {
            opacity: 1;
            transform: translateY(0);
        }
    }
</style>
""", unsafe_allow_html=True)

# Header
st.markdown("""
<div class='main-title'>
    <h1>💰 Aplikasi Akuntansi Keuangan</h1>
    <p>Kelola keuangan bisnis Anda dengan mudah dan efisien</p>
</div>
""", unsafe_allow_html=True)

# ============================
# SESSION DATA
# ============================
if "transaksi" not in st.session_state:
    st.session_state.transaksi = []

# ============================
# FORMAT RUPIAH
# ============================
def to_rp(n):
    try:
        return "Rp {:,}".format(int(n)).replace(",", ".")
    except:
        return "Rp 0"

# ============================
# KLASIFIKASI AKUN
# ============================
pendapatan_akun = ["Pendapatan Jasa", "Pendapatan Lainnya"]
beban_akun = ["Beban Gaji", "Beban Listrik", "Beban Sewa", "Beban Lainnya"]

# ============================
# FUNGSI AKUNTANSI
# ============================
def tambah_transaksi(tgl, akun, ket, debit, kredit, bulan, tahun):
    st.session_state.transaksi.append({
        "Tanggal": tgl,
        "Akun": akun,
        "Keterangan": ket,
        "Debit": int(debit),
        "Kredit": int(kredit),
        "Bulan": bulan,
        "Tahun": int(tahun)
    })

def hapus_transaksi(idx):
    st.session_state.transaksi.pop(idx)

def buku_besar(df):
    akun_list = df["Akun"].unique()
    buku_besar_data = {}
    for akun in akun_list:
        df_akun = df[df["Akun"] == akun].copy()
        df_akun["Saldo"] = df_akun["Debit"].cumsum() - df_akun["Kredit"].cumsum()
        buku_besar_data[akun] = df_akun
    return buku_besar_data

def neraca_saldo(df):
    grouped = df.groupby("Akun")[["Debit", "Kredit"]].sum()
    grouped["Saldo"] = grouped["Debit"] - grouped["Kredit"]
    return grouped

def laporan_laba_rugi(df):
    # PERBAIKAN: Pendapatan dicatat di DEBIT, jadi kita ambil debit untuk pendapatan
    total_pendapatan = df[df["Akun"].isin(pendapatan_akun)]["Debit"].sum()
    # PERBAIKAN: Beban dicatat di KREDIT, jadi kita ambil kredit untuk beban
    total_beban = df[df["Akun"].isin(beban_akun)]["Kredit"].sum()
    laba_rugi = total_pendapatan - total_beban
    return {
        "Total Pendapatan": total_pendapatan,
        "Total Beban": total_beban,
        "Laba/Rugi": laba_rugi
    }

# ============================
# FUNGSI EXPORT EXCEL
# ============================
def export_excel_multi(df):
    output = io.BytesIO()
    wb = Workbook()

    # =====================
    # STYLE
    # =====================
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    header_fill = PatternFill("solid", fgColor="4472C4")
    title_fill = PatternFill("solid", fgColor="B4C7E7")
    year_fill = PatternFill("solid", fgColor="D9E1F2")

    # =====================
    # SHEET 1: LAPORAN KEUANGAN
    # =====================
    ws_main = wb.active
    ws_main.title = "Laporan Keuangan"

    df_sorted = df.sort_values(["Tahun", "Bulan", "Tanggal"])

    current_row = 1
    tahun_sekarang = None

    for (tahun, bulan), grup in df_sorted.groupby(["Tahun", "Bulan"]):

        if tahun != tahun_sekarang:
            ws_main.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=5)
            cell = ws_main.cell(row=current_row, column=1, value=f"Laporan Keuangan Tahun {tahun}")
            cell.font = Font(bold=True, size=14)
            cell.alignment = Alignment(horizontal="center")
            cell.fill = year_fill
            current_row += 1
            tahun_sekarang = tahun

        nama_bulan = calendar.month_name[bulan].capitalize()
        ws_main.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=5)
        cell = ws_main.cell(row=current_row, column=1, value=f"Bulan {nama_bulan}")
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
        cell.fill = title_fill
        current_row += 1

        headers = ["Tanggal", "Akun", "Keterangan", "Debit", "Kredit"]
        for col, h in enumerate(headers, start=1):
            c = ws_main.cell(row=current_row, column=col, value=h)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = header_fill
            c.border = thin_border
        current_row += 1

        for r in dataframe_to_rows(grup[headers], index=False, header=False):
            for col, val in enumerate(r, start=1):
                cell = ws_main.cell(row=current_row, column=col, value=val)
                cell.border = thin_border
            current_row += 1

        current_row += 1

    for col, w in zip(["A", "B", "C", "D", "E"], [20, 18, 25, 20, 20]):
        ws_main.column_dimensions[col].width = w

    # =====================
    # SHEET 2: JURNAL UMUM
    # =====================
    ws_jurnal = wb.create_sheet("Jurnal Umum")
    ws_jurnal.append(["Tanggal", "Akun", "Keterangan", "Debit", "Kredit"])

    for r in dataframe_to_rows(df[["Tanggal", "Akun", "Keterangan", "Debit", "Kredit"]],
                               index=False, header=False):
        ws_jurnal.append(r)

    # =====================
    # SHEET 3: BUKU BESAR
    # =====================
    ws_bb = wb.create_sheet("Buku Besar")
    bb = buku_besar(df)
    row = 1

    for akun, data in bb.items():
        ws_bb.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        ws_bb.cell(row=row, column=1, value=f"Buku Besar - {akun}").font = Font(bold=True)
        row += 1

        ws_bb.append(["Tanggal", "Akun", "Keterangan", "Debit", "Kredit", "Saldo"])
        for r in dataframe_to_rows(data[["Tanggal", "Akun", "Keterangan", "Debit", "Kredit", "Saldo"]],
                                   index=False, header=False):
            ws_bb.append(r)
        row += len(data) + 2

    # =====================
    # SHEET 4: NERACA SALDO
    # =====================
    ws_ns = wb.create_sheet("Neraca Saldo")
    ns = neraca_saldo(df).reset_index()
    ws_ns.append(["Akun", "Debit", "Kredit", "Saldo"])
    for r in dataframe_to_rows(ns, index=False, header=False):
        ws_ns.append(r)

    # =====================
    # SHEET 5: LABA RUGI
    # =====================
    ws_lr = wb.create_sheet("Laba Rugi")
    lr = laporan_laba_rugi(df)

    ws_lr.append(["Keterangan", "Jumlah"])
    ws_lr.append(["Total Pendapatan", lr["Total Pendapatan"]])
    ws_lr.append(["Total Beban", lr["Total Beban"]])
    ws_lr.append(["Laba / Rugi", lr["Laba/Rugi"]])

    wb.save(output)
    output.seek(0)
    return output.getvalue()

# ============================
# MENU NAVIGASI
# ============================
st.sidebar.markdown("### 📋 Menu Navigasi")
menu = st.sidebar.radio(
    "",
    ["🏠 Dashboard", "📝 Input Transaksi", "📋 Lihat Transaksi", "📖 Buku Besar", 
     "⚖️ Neraca Saldo", "💰 Laporan Laba Rugi", "📈 Grafik", "📤 Export Excel", "📥 Import Excel"],
    label_visibility="collapsed"
)

# Info di sidebar
st.sidebar.markdown("---")
st.sidebar.markdown("### 📊 Statistik")
total_transaksi = len(st.session_state.transaksi)
st.sidebar.info(f"Total Transaksi: **{total_transaksi}**")

if total_transaksi > 0:
    df_temp = pd.DataFrame(st.session_state.transaksi)
    total_debit = df_temp["Debit"].sum()
    total_kredit = df_temp["Kredit"].sum()
    st.sidebar.success(f"Total Debit: **{to_rp(total_debit)}**")
    st.sidebar.warning(f"Total Kredit: **{to_rp(total_kredit)}**")

# ============================
# 0. DASHBOARD
# ============================
if menu == "🏠 Dashboard":
    st.markdown("<div class='subtitle'>🏠 Dashboard Overview</div>", unsafe_allow_html=True)
    
    if len(st.session_state.transaksi) == 0:
        st.info("👋 Selamat datang! Mulai dengan menambahkan transaksi pertama Anda.")
        st.markdown("""
        ### 📚 Panduan Penggunaan:
        1. **Input Transaksi** - Tambahkan transaksi baru
        2. **Lihat Transaksi** - Review dan hapus transaksi
        3. **Buku Besar** - Lihat detail per akun
        4. **Neraca Saldo** - Ringkasan semua akun
        5. **Laporan Laba Rugi** - Analisis profit/loss
        6. **Grafik** - Visualisasi data
        7. **Export Excel** - Download laporan lengkap
        """)
    else:
        df = pd.DataFrame(st.session_state.transaksi)
        lr = laporan_laba_rugi(df)
        
        # Metrics
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("📊 Total Transaksi", f"{len(df)}")
        with col2:
            st.metric("💵 Total Pendapatan", to_rp(lr["Total Pendapatan"]))
        with col3:
            st.metric("💸 Total Beban", to_rp(lr["Total Beban"]))
        with col4:
            if lr["Laba/Rugi"] >= 0:
                st.metric("✅ Laba Bersih", to_rp(lr["Laba/Rugi"]))
            else:
                st.metric("⚠️ Rugi Bersih", to_rp(abs(lr["Laba/Rugi"])))
        
        st.markdown("---")
        
        # Recent Transactions
        st.markdown("### 📋 Transaksi Terbaru")
        df_display = df.tail(5).copy()
        df_display["Debit"] = df_display["Debit"].apply(to_rp)
        df_display["Kredit"] = df_display["Kredit"].apply(to_rp)
        st.dataframe(df_display, use_container_width=True, hide_index=True)

# ============================
# 1. INPUT TRANSAKSI
# ============================
elif menu == "📝 Input Transaksi":
    st.markdown("<div class='subtitle'>📝 Input Transaksi Baru</div>", unsafe_allow_html=True)
    
    st.markdown("""
    <div style='background: #e3f2fd; padding: 15px; border-radius: 10px; margin-bottom: 20px;'>
        <h4 style='color: #1976d2; margin: 0;'>💡 Tips Pencatatan:</h4>
        <ul style='color: #1976d2; margin: 5px 0;'>
            <li><strong>Pendapatan</strong> → Dicatat di kolom <strong>DEBIT</strong></li>
            <li><strong>Beban</strong> → Dicatat di kolom <strong>KREDIT</strong></li>
            <li><strong>Kas Masuk</strong> → Kas di <strong>DEBIT</strong>, Pendapatan di <strong>DEBIT</strong></li>
            <li><strong>Kas Keluar</strong> → Beban di <strong>KREDIT</strong>, Kas di <strong>KREDIT</strong></li>
        </ul>
    </div>
    """, unsafe_allow_html=True)
    
    with st.form("form_transaksi", clear_on_submit=True):
        col1, col2 = st.columns(2)
        
        with col1:
            tgl = st.datetime_input(
    "📅 Tanggal & Waktu Transaksi",
    datetime.now()
)
            bulan = st.selectbox(
    "🗓️ Bulan Periode",
    list(calendar.month_name)[1:]
)

tahun = st.number_input(
    "📆 Tahun Periode",
    min_value=2000,
    max_value=2100,
    value=datetime.now().year
)
            akun = st.selectbox("🏦 Pilih Akun", 
                ["Kas", "Piutang", "Modal", "Pendapatan Jasa", "Pendapatan Lainnya", 
                 "Beban Gaji", "Beban Listrik", "Beban Sewa", "Beban Lainnya"])
            ket = st.text_input("📝 Keterangan", placeholder="Contoh: Pembayaran gaji karyawan")
        
        with col2:
            st.markdown("#### 💰 Jumlah Transaksi")
            debit = st.number_input("Debit (Rp)", min_value=0, step=10000, format="%d")
            kredit = st.number_input("Kredit (Rp)", min_value=0, step=10000, format="%d")
            
        st.markdown("---")
        col_btn1, col_btn2, col_btn3 = st.columns([1,1,3])
        with col_btn1:
            submit = st.form_submit_button("✅ Simpan Transaksi", use_container_width=True)
        with col_btn2:
            cancel = st.form_submit_button("🔄 Reset", use_container_width=True)
        
        if submit:
            if debit == 0 and kredit == 0:
                st.error("❌ Debit atau Kredit harus diisi!")
            elif ket.strip() == "":
                st.error("❌ Keterangan harus diisi!")
            else:
                tambah_transaksi(
    tgl, akun, ket, debit, kredit,
    bulan, tahun
)
                st.success("✅ Transaksi berhasil ditambahkan!")
                st.balloons()
                st.rerun()

# ============================
# 2. LIHAT TRANSAKSI
# ============================
elif menu == "📋 Lihat Transaksi":
    st.markdown("<div class='subtitle'>📋 Daftar Semua Transaksi</div>", unsafe_allow_html=True)
    
    if len(st.session_state.transaksi) == 0:
        st.info("📭 Belum ada transaksi yang tercatat.")
    else:
        df = pd.DataFrame(st.session_state.transaksi)
        
        # Filter
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            filter_akun = st.multiselect("🔍 Filter berdasarkan Akun", df["Akun"].unique())
        with col_f2:
            sort_by = st.selectbox("📊 Urutkan berdasarkan", ["Tanggal", "Akun", "Debit", "Kredit"])
        
        df_filtered = df.copy()
        if filter_akun:
            df_filtered = df_filtered[df_filtered["Akun"].isin(filter_akun)]
        
        df_filtered = df_filtered.sort_values(sort_by)
        df_display = df_filtered.copy()
        df_display["Debit"] = df_display["Debit"].apply(to_rp)
        df_display["Kredit"] = df_display["Kredit"].apply(to_rp)
        
        st.dataframe(df_display, use_container_width=True, hide_index=True)
        
        st.markdown("---")
        st.markdown("### 🗑️ Hapus Transaksi")
        col_h1, col_h2 = st.columns([3, 1])
        with col_h1:
            idx_hapus = st.number_input("Nomor indeks transaksi yang ingin dihapus", 
                                       min_value=0, max_value=len(st.session_state.transaksi)-1, step=1)
        with col_h2:
            if st.button("🗑️ Hapus", use_container_width=True):
                hapus_transaksi(idx_hapus)
                st.success("✅ Transaksi berhasil dihapus!")
                st.rerun()

# ============================
# 3. BUKU BESAR
# ============================
elif menu == "📖 Buku Besar":
    st.markdown("<div class='subtitle'>📖 Buku Besar Per Akun</div>", unsafe_allow_html=True)
    
    if len(st.session_state.transaksi) == 0:
        st.info("📭 Belum ada transaksi untuk ditampilkan.")
    else:
        df = pd.DataFrame(st.session_state.transaksi)
        bb = buku_besar(df)
        
        for idx, (akun, data) in enumerate(bb.items()):
            with st.expander(f"📊 {akun}", expanded=(idx==0)):
                data_display = data.copy()
                data_display["Debit"] = data_display["Debit"].apply(to_rp)
                data_display["Kredit"] = data_display["Kredit"].apply(to_rp)
                data_display["Saldo"] = data_display["Saldo"].apply(to_rp)
                st.dataframe(data_display, use_container_width=True, hide_index=True)

# ============================
# 4. NERACA SALDO
# ============================
elif menu == "⚖️ Neraca Saldo":
    st.markdown("<div class='subtitle'>⚖️ Neraca Saldo</div>", unsafe_allow_html=True)
    
    if len(st.session_state.transaksi) == 0:
        st.info("📭 Belum ada transaksi untuk ditampilkan.")
    else:
        df = pd.DataFrame(st.session_state.transaksi)
        ns = neraca_saldo(df)
        ns_display = ns.copy()
        ns_display["Debit"] = ns_display["Debit"].apply(to_rp)
        ns_display["Kredit"] = ns_display["Kredit"].apply(to_rp)
        ns_display["Saldo"] = ns_display["Saldo"].apply(to_rp)
        st.dataframe(ns_display, use_container_width=True)

# ============================
# 5. LAPORAN LABA RUGI
# ============================
elif menu == "💰 Laporan Laba Rugi":
    st.markdown("<div class='subtitle'>💰 Laporan Laba Rugi</div>", unsafe_allow_html=True)
    
    if len(st.session_state.transaksi) == 0:
        st.info("📭 Belum ada transaksi untuk dianalisis.")
    else:
        df = pd.DataFrame(st.session_state.transaksi)
        lr = laporan_laba_rugi(df)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("""
            <div style='background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%); 
                        padding: 25px; border-radius: 12px; color: white; text-align: center;'>
                <h3 style='margin: 0; font-size: 18px;'>💵 Total Pendapatan</h3>
                <h2 style='margin: 10px 0 0 0; font-size: 28px;'>""" + to_rp(lr["Total Pendapatan"]) + """</h2>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown("""
            <div style='background: linear-gradient(135deg, #ee0979 0%, #ff6a00 100%); 
                        padding: 25px; border-radius: 12px; color: white; text-align: center;'>
                <h3 style='margin: 0; font-size: 18px;'>💸 Total Beban</h3>
                <h2 style='margin: 10px 0 0 0; font-size: 28px;'>""" + to_rp(lr["Total Beban"]) + """</h2>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            if lr["Laba/Rugi"] >= 0:
                st.markdown("""
                <div style='background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
                            padding: 25px; border-radius: 12px; color: white; text-align: center;'>
                    <h3 style='margin: 0; font-size: 18px;'>✅ Laba Bersih</h3>
                    <h2 style='margin: 10px 0 0 0; font-size: 28px;'>""" + to_rp(lr["Laba/Rugi"]) + """</h2>
                </div>
                """, unsafe_allow_html=True)
            else:
                st.markdown("""
                <div style='background: linear-gradient(135deg, #f2709c 0%, #ff9472 100%); 
                            padding: 25px; border-radius: 12px; color: white; text-align: center;'>
                    <h3 style='margin: 0; font-size: 18px;'>⚠️ Rugi Bersih</h3>
                    <h2 style='margin: 10px 0 0 0; font-size: 28px;'>""" + to_rp(abs(lr["Laba/Rugi"])) + """</h2>
                </div>
                """, unsafe_allow_html=True)
        
        st.markdown("---")
        
        # Detail breakdown
        st.markdown("### 📊 Detail Perhitungan")
        detail_data = {
            "Keterangan": ["Total Pendapatan", "Total Beban", "Laba/Rugi Bersih"],
            "Jumlah": [to_rp(lr["Total Pendapatan"]), to_rp(lr["Total Beban"]), 
                      to_rp(lr["Laba/Rugi"]) if lr["Laba/Rugi"] >= 0 else f"({to_rp(abs(lr['Laba/Rugi']))})"]
        }
        st.table(pd.DataFrame(detail_data))

# ============================
# 6. GRAFIK
# ============================
elif menu == "📈 Grafik":
    st.markdown("<div class='subtitle'>📈 Visualisasi Data Akuntansi</div>", unsafe_allow_html=True)

    if len(st.session_state.transaksi) == 0:
        st.info("📭 Belum ada data untuk divisualisasikan.")
    else:
        df = pd.DataFrame(st.session_state.transaksi)
        
        tab1, tab2, tab3 = st.tabs(["📊 Debit per Akun", "📊 Kredit per Akun", "📊 Perbandingan"])
        
        with tab1:
            chart = alt.Chart(df).mark_bar().encode(
                x=alt.X("Akun:N", title="Akun"),
                y=alt.Y("Debit:Q", title="Debit (Rp)"),
                color=alt.Color("Akun:N", legend=None),
                tooltip=["Akun", "Debit"]
            ).properties(
                title="Grafik Total Debit per Akun",
                height=400
            )
            st.altair_chart(chart, use_container_width=True)
        
        with tab2:
            chart2 = alt.Chart(df).mark_bar().encode(
                x=alt.X("Akun:N", title="Akun"),
                y=alt.Y("Kredit:Q", title="Kredit (Rp)"),
                color=alt.Color("Akun:N", legend=None),
                tooltip=["Akun", "Kredit"]
            ).properties(
                title="Grafik Total Kredit per Akun",
                height=400
            )
            st.altair_chart(chart2, use_container_width=True)
        
        with tab3:
            df_grouped = df.groupby("Akun")[["Debit", "Kredit"]].sum().reset_index()
            df_melted = df_grouped.melt(id_vars="Akun", value_vars=["Debit", "Kredit"], 
                                        var_name="Tipe", value_name="Jumlah")
            
            chart3 = alt.Chart(df_melted).mark_bar().encode(
                x=alt.X("Akun:N", title="Akun"),
                y=alt.Y("Jumlah:Q", title="Jumlah (Rp)"),
                color="Tipe:N",
                xOffset="Tipe:N",
                tooltip=["Akun", "Tipe", "Jumlah"]
            ).properties(
                title="Perbandingan Debit vs Kredit per Akun",
                height=400
            )
            st.altair_chart(chart3, use_container_width=True)

# ============================
# 7. EXPORT EXCEL
# ============================
elif menu == "📤 Export Excel":
    st.markdown("<div class='subtitle'>📤 Export Laporan ke Excel</div>", unsafe_allow_html=True)

    if len(st.session_state.transaksi) == 0:
        st.info("📭 Belum ada transaksi untuk diekspor.")
    else:
        st.markdown("""
        <div style='background: #fff3cd; padding: 20px; border-radius: 10px; border-left: 5px solid #ffc107; margin-bottom: 20px;'>
            <h4 style='color: #856404; margin: 0 0 10px 0;'>📦 File Excel akan berisi:</h4>
            <ul style='color: #856404; margin: 0;'>
                <li>📄 Sheet 1: Laporan Keuangan (per bulan)</li>
                <li>📄 Sheet 2: Jurnal Umum</li>
                <li>📄 Sheet 3: Buku Besar</li>
                <li>📄 Sheet 4: Neraca Saldo</li>
                <li>📄 Sheet 5: Laporan Laba Rugi</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
        
        df = pd.DataFrame(st.session_state.transaksi)
        
        # Preview data
        st.markdown("### 👁️ Preview Data")
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            st.metric("Total Transaksi", len(df))
        with col_p2:
            st.metric("Total Akun Unik", df["Akun"].nunique())
        
        # Generate Excel
        try:
            excel_file = export_excel_multi(df)
            
            st.markdown("---")
            col_d1, col_d2, col_d3 = st.columns([1, 2, 1])
            with col_d2:
                st.download_button(
                    label="📥 Download Laporan Excel Lengkap",
                    data=excel_file,
                    file_name=f"laporan_akuntansi_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            st.success("✅ File Excel siap didownload!")
        except Exception as e:
            st.error(f"❌ Terjadi kesalahan saat membuat file Excel: {str(e)}")

# ============================
# 8. IMPORT EXCEL
# ============================
elif menu == "📥 Import Excel":
    st.markdown("<div class='subtitle'>📥 Import Transaksi dari Excel</div>", unsafe_allow_html=True)

    file = st.file_uploader("Upload file Excel", type=["xlsx"])

    if file:
        df_import = pd.read_excel(file)

        kolom = ["Tanggal", "Akun", "Keterangan", "Debit", "Kredit", "Bulan", "Tahun"]

        if not all(k in df_import.columns for k in kolom):
            st.error("❌ Kolom Excel tidak sesuai format")
        else:
            for _, r in df_import.iterrows():
                tambah_transaksi(
                    pd.to_datetime(r["Tanggal"]),
                    r["Akun"],
                    r["Keterangan"],
                    r["Debit"],
                    r["Kredit"],
                    r["Bulan"],
                    r["Tahun"]
                )

            st.success("✅ Import berhasil")
            st.rerun()


# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #888; padding: 20px;'>
    <p style='margin: 0;'>💰 <strong>Aplikasi Akuntansi Profesional</strong></p>
    <p style='margin: 5px 0 0 0; font-size: 14px;'>Kelola keuangan bisnis Anda dengan mudah dan efisien</p>
</div>
""", unsafe_allow_html=True)
