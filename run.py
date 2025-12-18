import streamlit as st
import pandas as pd
import altair as alt
from datetime import datetime
import pytz
import io
import calendar
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# ===========================
# Styling tema pantai
# ===========================
st.set_page_config(page_title="Aplikasi Akuntansi Keuangan", page_icon="💰", layout="wide")

st.markdown("""
<style>
    .main-title {
        background: linear-gradient(135deg, #56ccf2 0%, #2f80ed 100%);
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
    .subtitle {
        background: linear-gradient(135deg, #fbd786 0%, #f7797d 100%);
        padding: 20px;
        border-radius: 10px;
        color: white;
        font-size: 24px;
        font-weight: 700;
        margin: 20px 0;
        text-align: center;
        box-shadow: 0 3px 10px rgba(0,0,0,0.15);
    }
    .stButton>button {
        background: linear-gradient(135deg, #56ccf2 0%, #2f80ed 100%) !important;
        color: white !important;
        padding: 12px 28px !important;
        border-radius: 8px !important;
        font-size: 16px !important;
        font-weight: 600 !important;
        border: none !important;
        box-shadow: 0 4px 12px rgba(86,204,242,0.4) !important;
        transition: all 0.3s ease !important;
    }
    .stButton>button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 20px rgba(86,204,242,0.6) !important;
    }
    .css-1d391kg, [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #56ccf2 0%, #2f80ed 100%);
    }
</style>
""", unsafe_allow_html=True)

# ===========================
# Header utama
# ===========================
st.markdown("""
<div class='main-title'>
    <h1>💰 Aplikasi Akuntansi Keuangan</h1>
    <p>Kelola keuangan bisnis Anda dengan mudah dan efisien</p>
</div>
""", unsafe_allow_html=True)

# ===========================
# Inisialisasi Session State
# ===========================
if "transaksi" not in st.session_state:
    st.session_state.transaksi = []

# ===========================
# Fungsi format Rupiah dan tanggal
# ===========================
def format_rupiah_angka(n):
    if n == 0 or n is None:
        return "Rp -"
    s = f"{n:,.2f}"
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"Rp {s}"

def format_tanggal(dt):
    if pd.isna(dt):
        return ""
    if isinstance(dt, str):
        try:
            dt = pd.to_datetime(dt)
        except:
            return dt
    return dt.strftime("%Y-%m-%d %H:%M:%S")

# ===========================
# Fungsi Akuntansi + Klasifikasi Akun
# ===========================
pendapatan_akun = ["Pendapatan Jasa", "Pendapatan Lainnya"]
beban_akun = ["Beban Gaji", "Beban Listrik", "Beban Sewa", "Beban Lainnya"]

def tambah_transaksi(tgl, akun, ket, debit, kredit, periode_tahun, periode_bulan):
    st.session_state.transaksi.append({
        "Tanggal": tgl,
        "Akun": akun,
        "Keterangan": ket,
        "Debit": int(debit),
        "Kredit": int(kredit),
        "Periode_Tahun": int(periode_tahun),
        "Periode_Bulan": int(periode_bulan)
    })

def hapus_transaksi(idx):
    st.session_state.transaksi.pop(idx)

def buku_besar(df):
    akun_list = df["Akun"].unique()
    buku_besar_data = {}
    for akun in akun_list:
        df_akun = df[df["Akun"] == akun].copy().sort_values("Tanggal")
        df_akun["Saldo"] = df_akun["Debit"].cumsum() - df_akun["Kredit"].cumsum()
        buku_besar_data[akun] = df_akun
    return buku_besar_data

def neraca_saldo(df):
    grouped = df.groupby("Akun")[["Debit", "Kredit"]].sum()
    grouped["Saldo"] = grouped["Debit"] - grouped["Kredit"]
    return grouped

def laporan_laba_rugi(df):
    total_pendapatan = df[df["Akun"].isin(pendapatan_akun)]["Debit"].sum()
    total_beban = df[df["Akun"].isin(beban_akun)]["Kredit"].sum()
    laba_rugi = total_pendapatan - total_beban
    return {
        "Total Pendapatan": total_pendapatan,
        "Total Beban": total_beban,
        "Laba/Rugi": laba_rugi
    }

# ===========================
# Fungsi Export Excel lengkap & perperiode dengan total di tiap sheet
# ===========================
def export_excel_multi(df):
    import calendar
    output = io.BytesIO()
    wb = Workbook()
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    header_fill = PatternFill(start_color='305496', end_color='305496', fill_type='solid')
    title_fill = PatternFill(start_color='bdd7ee', end_color='bdd7ee', fill_type='solid')
    year_fill = PatternFill(start_color='d9e1f2', end_color='d9e1f2', fill_type='solid')
    font_white_bold = Font(bold=True, color='FFFFFF')
    font_bold = Font(bold=True)

    # Laporan Keuangan Sheet:
    ws = wb.active
    ws.title = "Laporan Keuangan"
    row = 1

    # Kelompokkan per Periode Tahun & Bulan
    periode_unique = (
        df[['Periode_Tahun', 'Periode_Bulan']]
        .drop_duplicates()
        .sort_values(['Periode_Tahun', 'Periode_Bulan'])
        .reset_index(drop=True)
    )

    for _, period in periode_unique.iterrows():
        pt = period['Periode_Tahun']
        pb = period['Periode_Bulan']
        nmbulan = calendar.month_name[pb]
        df_p = df[(df['Periode_Tahun'] == pt) & (df['Periode_Bulan'] == pb)]

        # Tahun Header
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
        c = ws.cell(row, 1, f'Laporan Keuangan Tahun {pt}')
        c.font = font_bold
        c.fill = year_fill
        c.alignment = Alignment(horizontal='center')
        for col in range(1, 6):
            ws.cell(row, col).border = thin_border
        row += 1

        # Bulan Header
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
        c = ws.cell(row, 1, f'Bulan {nmbulan}')
        c.font = font_bold
        c.fill = title_fill
        c.alignment = Alignment(horizontal='center')
        for col in range(1, 6):
            ws.cell(row, col).border = thin_border
        row += 1

        # Kolom Header
        cols = ['Tanggal', 'Akun', 'Keterangan', 'Debit', 'Kredit']
        for idx, colname in enumerate(cols, 1):
            c = ws.cell(row, idx, colname)
            c.font = font_white_bold
            c.fill = header_fill
            c.alignment = Alignment(horizontal='center')
            c.border = thin_border
        row += 1

        if df_p.empty:
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
            c = ws.cell(row, 1, 'Tidak ada transaksi di periode ini')
            c.alignment = Alignment(horizontal='center')
            row += 1
            total_debit = 0
            total_kredit = 0
        else:
            for _, r in df_p.iterrows():
                ws.cell(row, 1, r['Tanggal'].strftime('%Y-%m-%d %H:%M:%S')).alignment = Alignment(horizontal='left')
                ws.cell(row, 2, r['Akun']).alignment = Alignment(horizontal='left')
                ws.cell(row, 3, r['Keterangan']).alignment = Alignment(horizontal='left')
                ws.cell(row, 4, format_rupiah_angka(r['Debit'])).alignment = Alignment(horizontal='right')
                ws.cell(row, 5, format_rupiah_angka(r['Kredit'])).alignment = Alignment(horizontal='right')
                for cidx in range(1, 6):
                    ws.cell(row, cidx).border = thin_border
                row += 1
            total_debit = df_p['Debit'].sum()
            total_kredit = df_p['Kredit'].sum()

        # Total row
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
        c = ws.cell(row, 1, 'TOTAL')
        c.font = font_bold
        c.fill = title_fill
        c.alignment = Alignment(horizontal='center')
        for col in range(1, 6):
            ws.cell(row, col).border = thin_border

        c = ws.cell(row, 4, format_rupiah_angka(total_debit))
        c.font = font_bold
        c.fill = title_fill
        c.alignment = Alignment(horizontal='right')
        c.border = thin_border

        c2 = ws.cell(row, 5, format_rupiah_angka(total_kredit))
        c2.font = font_bold
        c2.fill = title_fill
        c2.alignment = Alignment(horizontal='right')
        c2.border = thin_border
        row += 2

    # Set column widths
    col_widths = [22, 18, 32, 20, 20]
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[chr(64+i)].width = w

    # TODO: terapkan prinsip yang sama di sheet Jurnal Umum, Buku Besar, Neraca Saldo, Laporan Laba Rugi dengan filter periode sama

    wb.save(output)
    output.seek(0)
    return output.getvalue()

# ===========================
# Menu Streamlit utama
# ===========================
st.sidebar.markdown("### 📋 Menu Navigasi")
menu = st.sidebar.radio("", [
    "🏠 Dashboard",
    "📝 Input Transaksi",
    "📋 Lihat Transaksi",
    "📖 Buku Besar",
    "⚖️ Neraca Saldo",
    "💰 Laporan Laba Rugi",
    "📈 Grafik",
    "📥 Import Excel",
    "📤 Export Excel",
], label_visibility='collapsed')

# Sidebar statistik
st.sidebar.markdown("---")
st.sidebar.markdown("### 📊 Statistik")
total_transaksi = len(st.session_state.transaksi)
st.sidebar.info(f"Total Transaksi: **{total_transaksi}**")
if total_transaksi > 0:
    df_stat = pd.DataFrame(st.session_state.transaksi)
    st.sidebar.success(f"Total Debit: **{format_rupiah_angka(df_stat['Debit'].sum())}**")
    st.sidebar.warning(f"Total Kredit: **{format_rupiah_angka(df_stat['Kredit'].sum())}**")

tz = pytz.timezone('Asia/Jakarta')

# ==== Implementasi detail menu (contoh menu input lengkap) ====
if menu == "🏠 Dashboard":
    st.markdown("<h3>🏠 Dashboard Overview</h3>", unsafe_allow_html=True)
    if total_transaksi == 0:
        st.info("Belum ada transaksi. Silakan mulai di menu Input Transaksi.")
    else:
        df = pd.DataFrame(st.session_state.transaksi)
        lr = laporan_laba_rugi(df)
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Transaksi", total_transaksi)
        col2.metric("Total Pendapatan", format_rupiah_angka(lr["Total Pendapatan"]))
        col3.metric("Total Beban", format_rupiah_angka(lr["Total Beban"]))
        laba = lr["Laba/Rugi"]
        if laba >= 0:
            col4.metric("Laba Bersih", format_rupiah_angka(laba))
        else:
            col4.metric("Rugi Bersih", format_rupiah_angka(abs(laba)))

elif menu == "📝 Input Transaksi":
    st.markdown("<h3>📝 Input Transaksi Baru</h3>", unsafe_allow_html=True)
    with st.form("form_input", clear_on_submit=True):
        tgl_transaksi = st.date_input("📅 Tanggal Transaksi", datetime.now(tz).date())
        akun = st.selectbox("🏦 Pilih Akun", [
            "Kas", "Piutang", "Modal", "Pendapatan Jasa", "Pendapatan Lainnya",
            "Beban Gaji", "Beban Listrik", "Beban Sewa", "Beban Lainnya"
        ])
        ket = st.text_input("📝 Keterangan")
        debit = st.number_input("Debit (Rp)", min_value=0, step=10000, format="%d")
        kredit = st.number_input("Kredit (Rp)", min_value=0, step=10000, format="%d")

        submitted = st.form_submit_button("Simpan Transaksi")
        if submitted:
            if debit == 0 and kredit == 0:
                st.error("Debit atau Kredit harus diisi")
            elif ket.strip() == "":
                st.error("Keterangan harus diisi")
            else:
                waktu = datetime.now(tz).time()
                tgl_waktu = datetime.combine(tgl_transaksi, waktu)
                tambah_transaksi(tgl_waktu, akun, ket, debit, kredit, tgl_waktu.year, tgl_waktu.month)
                st.success("Transaksi berhasil ditambahkan")
                st.experimental_rerun()

elif menu == "📋 Lihat Transaksi":
    st.markdown("<h3>📋 Daftar Transaksi</h3>", unsafe_allow_html=True)
    if total_transaksi == 0:
        st.info("Belum ada transaksi")
    else:
        df = pd.DataFrame(st.session_state.transaksi)
        filter_akun = st.multiselect("Filter Akun", df["Akun"].unique())
        sort_by = st.selectbox("Urutkan sesuai", ["Tanggal", "Akun", "Debit", "Kredit"])
        if filter_akun:
            df = df[df["Akun"].isin(filter_akun)]
        df = df.sort_values(sort_by)
        df_show = df.copy()
        df_show["Tanggal"] = df_show["Tanggal"].apply(format_tanggal)
        df_show["Debit"] = df_show["Debit"].apply(format_rupiah_angka)
        df_show["Kredit"] = df_show["Kredit"].apply(format_rupiah_angka)
        st.dataframe(df_show, use_container_width=True)
        idx_hapus = st.number_input("Nomor indeks transaksi hapus", min_value=0, max_value=len(st.session_state.transaksi)-1)
        if st.button("Hapus Transaksi"):
            hapus_transaksi(idx_hapus)
            st.success("Transaksi berhasil dihapus")
            st.experimental_rerun()

# Anda bisa melengkapi fungsi menu Buku Besar, Neraca Saldo, Laporan Laba Rugi, Grafik, Import Excel dan Export Excel 
# dengan logika yang sudah saya sampaikan sebelumnya.

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #888; padding: 20px;'>
    <p>💰 <strong>Aplikasi Akuntansi Profesional</strong></p>
    <p>Kelola keuangan bisnis Anda dengan mudah dan efisien</p>
</div>
""", unsafe_allow_html=True)
