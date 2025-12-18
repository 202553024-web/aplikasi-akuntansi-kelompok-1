import streamlit as st
import pandas as pd
import altair as alt
from datetime import datetime
import pytz
import io
import calendar
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# ====================
# Styling tema pantai
# ====================
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
    /* Sidebar */
    .css-1d391kg, [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #56ccf2 0%, #2f80ed 100%);
    }
</style>
""", unsafe_allow_html=True)

# ====================
# Header utama
# ====================
st.markdown("""
<div class='main-title'>
    <h1>💰 Aplikasi Akuntansi Keuangan</h1>
    <p>Kelola keuangan bisnis Anda dengan mudah dan efisien</p>
</div>
""", unsafe_allow_html=True)

# =====================
# Data storage session
# =====================
if "transaksi" not in st.session_state:
    st.session_state.transaksi = []

# =====================
# Fungsi Format Rupiah yang sesuai
# =====================
def format_rupiah_angka(n):
    """Format angka ke format Rupiah dengan titik ribuan dan koma desimal, string"""
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
    # Format lengkap dengan waktu
    return dt.strftime("%Y-%m-%d %H:%M:%S")

# =====================
# Fungsi Accounting
# =====================
pendapatan_akun = ["Pendapatan Jasa", "Pendapatan Lainnya"]
beban_akun = ["Beban Gaji", "Beban Listrik", "Beban Sewa", "Beban Lainnya"]

def tambah_transaksi(tgl, akun, ket, debit, kredit):
    st.session_state.transaksi.append({
        "Tanggal": tgl,
        "Akun": akun,
        "Keterangan": ket,
        "Debit": int(debit),
        "Kredit": int(kredit)
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

# =====================
# Export Excel - Semua Sheet dengan Format sesuai gambar Anda
# =====================
def export_excel_multi(df):
    output = io.BytesIO()
    wb = Workbook()

    thin_border = Border( left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    header_fill = PatternFill(start_color="305496", end_color="305496", fill_type="solid")   # header
    title_fill = PatternFill(start_color="bdd7ee", end_color="bdd7ee", fill_type="solid")    # bulan
    year_fill = PatternFill(start_color="d9e1f2", end_color="d9e1f2", fill_type="solid")     # tahun
    font_white_bold = Font(bold=True, color="FFFFFF")
    font_bold = Font(bold=True)

    df["Tanggal"] = pd.to_datetime(df["Tanggal"])
    df = df.sort_values("Tanggal")
    df["Tahun"] = df["Tanggal"].dt.year
    df["Bulan"] = df["Tanggal"].dt.month

    # =======================
    # Sheet 1: Laporan Keuangan
    # =======================
    ws = wb.active
    ws.title = "Laporan Keuangan"
    current_row = 1

    for tahun, df_tahun in df.groupby("Tahun"):
        # Header tahun
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=5)
        tcell = ws.cell(row=current_row, column=1, value=f"Laporan Keuangan Tahun {tahun}")
        tcell.font = Font(bold=True, size=14)
        tcell.fill = year_fill
        tcell.alignment = Alignment(horizontal="center", vertical="center")
        for col in range(1, 6):
            ws.cell(row=current_row, column=col).border = thin_border
        current_row += 1

        for bulan, df_bulan in df_tahun.groupby("Bulan"):
            nama_bulan = calendar.month_name[bulan]
            # Header bulan
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=5)
            bcell = ws.cell(row=current_row, column=1, value=f"Bulan {nama_bulan}")
            bcell.font = font_bold
            bcell.fill = title_fill
            bcell.alignment = Alignment(horizontal="center", vertical="center")
            for col in range(1, 6):
                ws.cell(row=current_row, column=col).border = thin_border

            current_row += 1

            # Header tabel
            headers = ["Tanggal", "Akun", "Keterangan", "Debit", "Kredit"]
            for idx, val in enumerate(headers, start=1):
                hcell = ws.cell(row=current_row, column=idx, value=val)
                hcell.font = font_white_bold
                hcell.fill = header_fill
                hcell.alignment = Alignment(horizontal="center", vertical="center")
                hcell.border = thin_border
            current_row += 1

            # Isi data
            for _, row in df_bulan.iterrows():
                ws.cell(row=current_row, column=1, value=row["Tanggal"].strftime("%Y-%m-%d %H:%M:%S")).alignment = Alignment(horizontal="left")
                ws.cell(row=current_row, column=2, value=row["Akun"]).alignment = Alignment(horizontal="left")
                ws.cell(row=current_row, column=3, value=row["Keterangan"]).alignment = Alignment(horizontal="left")

                # Format rupiah sesuai format di gambar
                debit_str = format_rupiah_angka(row["Debit"])
                kredit_str = format_rupiah_angka(row["Kredit"])

                ws.cell(row=current_row, column=4, value=debit_str).alignment = Alignment(horizontal="right")
                ws.cell(row=current_row, column=5, value=kredit_str).alignment = Alignment(horizontal="right")

                for col in range(1,6):
                    ws.cell(row=current_row, column=col).border = thin_border

                current_row +=1

            current_row += 1  # space antar bulan

        current_row += 1  # space antar tahun

    width_list = [22, 18, 32, 20, 20]
    for i, w in enumerate(width_list, start=1):
        ws.column_dimensions[chr(64+i)].width = w


    # =======================
    # Sheet 2 : Jurnal Umum (semua transaksi)
    # =======================
    ws2 = wb.create_sheet("Jurnal Umum")
    ws2.title = "Jurnal Umum"

    ws2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5)
    cell_title = ws2.cell(row=1, column=1, value="Jurnal Umum")
    cell_title.font = Font(bold=True, size=14)
    cell_title.alignment = Alignment(horizontal="center", vertical="center")

    headers = ["Tanggal", "Akun", "Keterangan", "Debit", "Kredit"]
    for idx, h in enumerate(headers, start=1):
        c = ws2.cell(row=2, column=idx, value=h)
        c.font = font_white_bold
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = thin_border

    rnum = 3
    for _, row in df.iterrows():
        ws2.cell(row=rnum, column=1, value=row["Tanggal"].strftime("%Y-%m-%d %H:%M:%S")).alignment = Alignment(horizontal="left")
        ws2.cell(row=rnum, column=2, value=row["Akun"]).alignment = Alignment(horizontal="left")
        ws2.cell(row=rnum, column=3, value=row["Keterangan"]).alignment = Alignment(horizontal="left")

        debit_str = format_rupiah_angka(row["Debit"])
        kredit_str = format_rupiah_angka(row["Kredit"])

        ws2.cell(row=rnum, column=4, value=debit_str).alignment = Alignment(horizontal="right")
        ws2.cell(row=rnum, column=5, value=kredit_str).alignment = Alignment(horizontal="right")

        for c in range(1,6):
            ws2.cell(row=rnum, column=c).border = thin_border

        rnum += 1

    widths = [22, 18, 30, 20, 20]
    for i, w in enumerate(widths, start=1):
        ws2.column_dimensions[chr(64 + i)].width = w

    # =======================
    # Sheet 3 : Buku Besar per akun
    # =======================
    ws3 = wb.create_sheet("Buku Besar")
    ws3.title = "Buku Besar"

    bb = buku_besar(df)
    r = 1
    for akun, data in bb.items():
        ws3.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        c = ws3.cell(row=r, column=1, value=f"Buku Besar - {akun}")
        c.font=font_bold
        c.alignment = Alignment(center=True)
        r += 1

        headers = ["Tanggal", "Akun", "Keterangan", "Debit", "Kredit", "Saldo"]
        for idx, h in enumerate(headers, start=1):
            cell = ws3.cell(row=r, column=idx, value=h)
            cell.font = font_white_bold
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
        r += 1

        for _, row in data.iterrows():
            ws3.cell(row=r, column=1, value=row["Tanggal"].strftime("%Y-%m-%d %H:%M:%S")).alignment = Alignment(horizontal="left")
            ws3.cell(row=r, column=2, value=row["Akun"]).alignment = Alignment(horizontal="left")
            ws3.cell(row=r, column=3, value=row["Keterangan"]).alignment = Alignment(horizontal="left")

            ws3.cell(row=r, column=4, value=format_rupiah_angka(row["Debit"])).alignment = Alignment(horizontal="right")
            ws3.cell(row=r, column=5, value=format_rupiah_angka(row["Kredit"])).alignment = Alignment(horizontal="right")
            ws3.cell(row=r, column=6, value=format_rupiah_angka(row["Saldo"])).alignment = Alignment(horizontal="right")

            for col in range(1,7):
                ws3.cell(row=r, column=col).border = thin_border
            r +=1
        r += 2

    widths = [22, 18, 30, 20, 20, 20]
    for i, w in enumerate(widths, start=1):
        ws3.column_dimensions[chr(64 + i)].width = w

    # =======================
    # Sheet 4 : Neraca Saldo
    # =======================
    ws4 = wb.create_sheet("Neraca Saldo")
    ws4.title = "Neraca Saldo"

    ns = neraca_saldo(df)
    ws4.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
    c = ws4.cell(row=1, column=1, value="Neraca Saldo")
    c.font = font_bold
    c.alignment = Alignment(horizontal="center")

    headers = ["Akun", "Debit", "Kredit", "Saldo"]
    for idx, h in enumerate(headers, start=1):
        cell = ws4.cell(row=2, column=idx, value=h)
        cell.font = font_white_bold
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
        cell.border = thin_border

    r = 3
    for _, row in ns.reset_index().iterrows():
        ws4.cell(row=r, column=1, value=row["Akun"]).alignment = Alignment(horizontal="left")
        ws4.cell(row=r, column=2, value=format_rupiah_angka(row["Debit"])).alignment = Alignment(horizontal="right")
        ws4.cell(row=r, column=3, value=format_rupiah_angka(row["Kredit"])).alignment = Alignment(horizontal="right")
        ws4.cell(row=r, column=4, value=format_rupiah_angka(row["Saldo"])).alignment = Alignment(horizontal="right")

        for col in range(1,5):
            ws4.cell(row=r, column=col).border = thin_border
        r += 1

    widths = [22, 20, 20, 20]
    for i, w in enumerate(widths, start=1):
        ws4.column_dimensions[chr(64 + i)].width = w

    # =======================
    # Sheet 5 : Laporan Laba Rugi
    # =======================
    ws5 = wb.create_sheet("Laporan Laba Rugi")
    ws5.title = "Laporan Laba Rugi"

    lr = laporan_laba_rugi(df)

    ws5.merge_cells(start_row=1, start_column=1, end_row=1, end_column=2)
    c = ws5.cell(row=1, column=1, value="Laporan Laba Rugi")
    c.font = font_bold
    c.alignment = Alignment(horizontal="center")
    c.fill = year_fill

    headers = ["Keterangan", "Jumlah"]
    for idx, h in enumerate(headers, start=1):
        cell = ws5.cell(row=2, column=idx, value=h)
        cell.font = font_white_bold
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
        cell.border = thin_border

    labels = ["Total Pendapatan", "Total Beban", "Laba/Rugi"]

    ns = [lr["Total Pendapatan"], lr["Total Beban"], lr["Laba/Rugi"]]

    r = 3
    for label, val in zip(labels, ns):
        ws5.cell(row=r, column=1, value=label).alignment = Alignment(horizontal="left")
        ws5.cell(row=r, column=1).border = thin_border
        # Laba rugi negatif dibungkus tanda kurung
        if label == "Laba/Rugi" and val < 0:
            val_str = f"(Rp {abs(val):,.2f})".replace(",", "X").replace(".", ",").replace("X", ".")
        else:
            val_str = format_rupiah_angka(val)
        c = ws5.cell(row=r, column=2, value=val_str)
        c.alignment = Alignment(horizontal="right")
        c.border = thin_border
        r += 1

    ws5.column_dimensions['A'].width = 25
    ws5.column_dimensions['B'].width = 20

    wb.save(output)
    output.seek(0)
    return output.getvalue()

# ====================
# Menu navigasi utama
# ====================
st.sidebar.markdown("### 📋 Menu Navigasi")
menu = st.sidebar.radio("",[
    "🏠 Dashboard",
    "📝 Input Transaksi",
    "📋 Lihat Transaksi",
    "📖 Buku Besar",
    "⚖️ Neraca Saldo",
    "💰 Laporan Laba Rugi",
    "📈 Grafik",
    "📥 Import Excel",
    "📤 Export Excel",
], label_visibility="collapsed")

st.sidebar.markdown("---")
st.sidebar.markdown("### 📊 Statistik")
total_transaksi = len(st.session_state.transaksi)
st.sidebar.info(f"Total Transaksi: **{total_transaksi}**")

if total_transaksi > 0:
    df_temp = pd.DataFrame(st.session_state.transaksi)
    total_debit = df_temp["Debit"].sum()
    total_kredit = df_temp["Kredit"].sum()
    st.sidebar.success(f"Total Debit: **{format_rupiah_angka(total_debit)}**")
    st.sidebar.warning(f"Total Kredit: **{format_rupiah_angka(total_kredit)}**")

# ====================
# Dashboard
# ====================
if menu == "🏠 Dashboard":
    st.markdown("<div class='subtitle'>🏠 Dashboard Overview</div>", unsafe_allow_html=True)
    if total_transaksi == 0:
        st.info("👋 Selamat datang! Mulai dengan menambahkan transaksi pertama Anda.")
        st.markdown("""
1. **Input Transaksi** - Tambahkan transaksi baru  
2. **Lihat Transaksi** - Review dan hapus transaksi  
3. **Buku Besar** - Lihat detail per akun  
4. **Neraca Saldo** - Ringkasan semua akun  
5. **Laporan Laba Rugi** - Analisis profit vs loss  
6. **Grafik** - Visualisasi data  
7. **Import Excel** - Import transaksi dari Excel  
8. **Export Excel** - Download laporan lengkap  
""")
    else:
        df = pd.DataFrame(st.session_state.transaksi)
        lr = laporan_laba_rugi(df)

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("📊 Total Transaksi", len(df))
        col2.metric("💵 Total Pendapatan", format_rupiah_angka(lr["Total Pendapatan"]))
        col3.metric("💸 Total Beban", format_rupiah_angka(lr["Total Beban"]))
        laba = lr["Laba/Rugi"]
        if laba >= 0:
            col4.metric("✅ Laba Bersih", format_rupiah_angka(laba))
        else:
            col4.metric("⚠️ Rugi Bersih", format_rupiah_angka(abs(laba)))

        st.markdown("---")

        st.markdown("### 📋 Transaksi Terbaru")
        df_show = df.tail(5).copy()
        df_show["Tanggal"] = df_show["Tanggal"].apply(format_tanggal)
        df_show["Debit"] = df_show["Debit"].apply(format_rupiah_angka)
        df_show["Kredit"] = df_show["Kredit"].apply(format_rupiah_angka)
        st.dataframe(df_show, use_container_width=True)

# ====================
# Input Transaksi
# ====================
elif menu == "📝 Input Transaksi":
    st.markdown("<div class='subtitle'>📝 Input Transaksi Baru</div>", unsafe_allow_html=True)

    with st.form("form_transaksi", clear_on_submit=True):
        tz = pytz.timezone('Asia/Jakarta')
        tgl_input = st.date_input("📅 Tanggal Transaksi", datetime.now(tz).date())
        akun = st.selectbox("🏦 Pilih Akun", [
            "Kas", "Piutang", "Modal", "Pendapatan Jasa", "Pendapatan Lainnya", 
            "Beban Gaji", "Beban Listrik", "Beban Sewa", "Beban Lainnya"])
        ket = st.text_input("📝 Keterangan", placeholder="Contoh: Pembayaran gaji karyawan")
        debit = st.number_input("Debit (Rp)", min_value=0, step=10000, format="%d")
        kredit = st.number_input("Kredit (Rp)", min_value=0, step=10000, format="%d")

        submitted = st.form_submit_button("✅ Simpan Transaksi")

        if submitted:
            if debit == 0 and kredit == 0:
                st.error("❌ Debit atau Kredit harus diisi!")
            elif ket.strip() == "":
                st.error("❌ Keterangan harus diisi!")
            else:
                waktu = datetime.now(tz).time()
                tgl_waktu = datetime.combine(tgl_input, waktu)
                tambah_transaksi(tgl_waktu, akun, ket, debit, kredit)
                st.success("✅ Transaksi berhasil ditambahkan!")
                st.balloons()
                st.experimental_rerun()

# ====================
# Lihat Transaksi
# ====================
elif menu == "📋 Lihat Transaksi":
    st.markdown("<div class='subtitle'>📋 Daftar Transaksi</div>", unsafe_allow_html=True)

    if len(st.session_state.transaksi) == 0:
        st.info("Belum ada transaksi.")
    else:
        df = pd.DataFrame(st.session_state.transaksi)
        col1, col2 = st.columns(2)
        with col1:
            filter_akun = st.multiselect("Filter Akun", df["Akun"].unique())
        with col2:
            sort_by = st.selectbox("Urutkan per", ["Tanggal", "Akun", "Debit", "Kredit"])

        if filter_akun:
            df = df[df["Akun"].isin(filter_akun)]
        df = df.sort_values(sort_by)

        df_show = df.copy()
        df_show["Tanggal"] = df_show["Tanggal"].apply(format_tanggal)
        df_show["Debit"] = df_show["Debit"].apply(format_rupiah_angka)
        df_show["Kredit"] = df_show["Kredit"].apply(format_rupiah_angka)
        st.dataframe(df_show, use_container_width=True)

        st.markdown("Hapus Transaksi")
        del_idx = st.number_input("Nomor indeks hapus", min_value=0, max_value=len(st.session_state.transaksi)-1, step=1)
        if st.button("🗑️ Hapus Transaksi"):
            hapus_transaksi(del_idx)
            st.success("Transaksi berhasil dihapus.")
            st.experimental_rerun()

# ====================
# Seleksi menu lainnya dan implementasi serupa, contoh:
# - Buku Besar
# - Neraca Saldo
# - Laporan Laba Rugi
# - Grafik
# - Import Excel
# - Export Excel (panggil export_excel_multi)
# ====================
# Anda bisa melanjutkan seperti contoh sebelumnya dengan logika dan tampilan sesuai kebutuhan

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #888; padding: 20px;'>
<p>💰 <strong>Aplikasi Akuntansi Profesional</strong></p>
<p>Kelola keuangan bisnis Anda dengan mudah dan efisien</p>
</div>
""", unsafe_allow_html=True)
