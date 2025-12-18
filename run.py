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
    /* Sidebar */
    .css-1d391kg, [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #56ccf2 0%, #2f80ed 100%);
    }
</style>
""", unsafe_allow_html=True)

# ===========================
# Header
# ===========================
st.markdown("""
<div class='main-title'>
    <h1>💰 Aplikasi Akuntansi Keuangan</h1>
    <p>Kelola keuangan bisnis Anda dengan mudah dan efisien</p>
</div>
""", unsafe_allow_html=True)

# ===========================
# Session state untuk simpan transaksi
# ===========================
if "transaksi" not in st.session_state:
    st.session_state.transaksi = []

# ===========================
# Fungsi format rupiah sesuai contoh
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
# Fungsi-fungsi akun
# ===========================
pendapatan_akun = ["Pendapatan Jasa", "Pendapatan Lainnya"]
beban_akun = ["Beban Gaji", "Beban Listrik", "Beban Sewa", "Beban Lainnya"]

def tambah_transaksi(data):
    st.session_state.transaksi.append(data)

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
# Fungsi export excel lengkap dan sesuai template
# ===========================
def export_excel_multi(df):
    output = io.BytesIO()
    wb = Workbook()

    thin_border = Border(left=Side(style='thin'),
                         right=Side(style='thin'),
                         top=Side(style='thin'),
                         bottom=Side(style='thin'))

    header_fill = PatternFill(start_color="305496", end_color="305496", fill_type="solid")   # biru header
    title_fill = PatternFill(start_color="bdd7ee", end_color="bdd7ee", fill_type="solid")    # biru muda
    year_fill = PatternFill(start_color="d9e1f2", end_color="d9e1f2", fill_type="solid")     # biru sangat muda

    font_white_bold = Font(bold=True, color="FFFFFF")
    font_bold = Font(bold=True)

    df["Tanggal"] = pd.to_datetime(df["Tanggal"])
    df = df.sort_values("Tanggal")
    df["Tahun"] = df["Tanggal"].dt.year
    df["Bulan"] = df["Tanggal"].dt.month

    # Sheet 1: Laporan Keuangan
    ws = wb.active
    ws.title = "Laporan Keuangan"
    current_row = 1

    for tahun, df_tahun in df.groupby("Tahun"):
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
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=5)
            bcell = ws.cell(row=current_row, column=1, value=f"Bulan {nama_bulan}")
            bcell.font = font_bold
            bcell.fill = title_fill
            bcell.alignment = Alignment(horizontal="center", vertical="center")
            for col in range(1, 6):
                ws.cell(row=current_row, column=col).border = thin_border

            current_row += 1

            headers = ["Tanggal", "Akun", "Keterangan", "Debit", "Kredit"]
            for idx, val in enumerate(headers, start=1):
                hcell = ws.cell(row=current_row, column=idx, value=val)
                hcell.font = font_white_bold
                hcell.fill = header_fill
                hcell.alignment = Alignment(horizontal="center", vertical="center")
                hcell.border = thin_border
            current_row += 1

            for _, row in df_bulan.iterrows():
                ws.cell(row=current_row, column=1, value=row["Tanggal"].strftime("%Y-%m-%d %H:%M:%S")).alignment = Alignment(horizontal="left")
                ws.cell(row=current_row, column=2, value=row["Akun"]).alignment = Alignment(horizontal="left")
                ws.cell(row=current_row, column=3, value=row["Keterangan"]).alignment = Alignment(horizontal="left")

                debit_str = format_rupiah_angka(row["Debit"])
                kredit_str = format_rupiah_angka(row["Kredit"])

                dcell = ws.cell(row=current_row, column=4, value=debit_str)
                dcell.alignment = Alignment(horizontal="right")
                dcell.border = thin_border

                kcell = ws.cell(row=current_row, column=5, value=kredit_str)
                kcell.alignment = Alignment(horizontal="right")
                kcell.border = thin_border

                for col in range(1, 6):
                    ws.cell(row=current_row, column=col).border = thin_border
                current_row += 1

            current_row += 1

        current_row += 1

    col_widths = [22, 18, 30, 20, 20]
    for i, width in enumerate(col_widths, start=1):
        ws.column_dimensions[chr(64 + i)].width = width

    # Sheet 2: Jurnal Umum
    ws2 = wb.create_sheet("Jurnal Umum")
    ws2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5)
    title_cell = ws2.cell(row=1, column=1, value="Jurnal Umum")
    title_cell.font = Font(bold=True, size=14)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    headers = ["Tanggal", "Akun", "Keterangan", "Debit", "Kredit"]
    for idx, val in enumerate(headers, start=1):
        hcell = ws2.cell(row=2, column=idx, value=val)
        hcell.font = font_white_bold
        hcell.fill = header_fill
        hcell.alignment = Alignment(horizontal="center", vertical="center")
        hcell.border = thin_border

    r = 3
    for _, row in df.iterrows():
        ws2.cell(row=r, column=1, value=row["Tanggal"].strftime("%Y-%m-%d %H:%M:%S")).alignment = Alignment(horizontal="left")
        ws2.cell(row=r, column=2, value=row["Akun"]).alignment = Alignment(horizontal="left")
        ws2.cell(row=r, column=3, value=row["Keterangan"]).alignment = Alignment(horizontal="left")

        debit_str = format_rupiah_angka(row["Debit"])
        kredit_str = format_rupiah_angka(row["Kredit"])

        ws2.cell(row=r, column=4, value=debit_str).alignment = Alignment(horizontal="right")
        ws2.cell(row=r, column=5, value=kredit_str).alignment = Alignment(horizontal="right")

        for col in range(1, 6):
            ws2.cell(row=r, column=col).border = thin_border
        r += 1

    for i, width in enumerate(col_widths, 1):
        ws2.column_dimensions[chr(64 + i)].width = width

    # Sheet 3: Buku Besar
    ws3 = wb.create_sheet("Buku Besar")
    bb = buku_besar(df)
    r = 1
    for akun, data in bb.items():
        ws3.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        a_cell = ws3.cell(row=r, column=1, value=f"Buku Besar - {akun}")
        a_cell.font = font_bold
        a_cell.alignment = Alignment(horizontal="center", vertical="center")
        r += 1

        headers = ["Tanggal", "Akun", "Keterangan", "Debit", "Kredit", "Saldo"]
        for idx, val in enumerate(headers, start=1):
            hcell = ws3.cell(row=r, column=idx, value=val)
            hcell.font = font_white_bold
            hcell.fill = header_fill
            hcell.alignment = Alignment(horizontal="center", vertical="center")
            hcell.border = thin_border
        r += 1

        for _, row in data.iterrows():
            ws3.cell(row=r, column=1, value=row["Tanggal"].strftime("%Y-%m-%d %H:%M:%S")).alignment = Alignment(horizontal="left")
            ws3.cell(row=r, column=2, value=row["Akun"]).alignment = Alignment(horizontal="left")
            ws3.cell(row=r, column=3, value=row["Keterangan"]).alignment = Alignment(horizontal="left")

            ws3.cell(row=r, column=4, value=format_rupiah_angka(row["Debit"])).alignment = Alignment(horizontal="right")
            ws3.cell(row=r, column=5, value=format_rupiah_angka(row["Kredit"])).alignment = Alignment(horizontal="right")
            ws3.cell(row=r, column=6, value=format_rupiah_angka(row["Saldo"])).alignment = Alignment(horizontal="right")

            for col in range(1, 7):
                ws3.cell(row=r, column=col).border = thin_border
            r += 1
        r += 2

    col_widths_bb = [22, 18, 30, 20, 20, 20]
    for i, width in enumerate(col_widths_bb, 1):
        ws3.column_dimensions[chr(64 + i)].width = width

    # Sheet 4: Neraca Saldo
    ws4 = wb.create_sheet("Neraca Saldo")
    ws4.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
    c = ws4.cell(row=1, column=1, value="Neraca Saldo")
    c.font = font_bold
    c.alignment = Alignment(horizontal="center", vertical="center")

    headers = ["Akun", "Debit", "Kredit", "Saldo"]
    for idx, val in enumerate(headers, start=1):
        hcell = ws4.cell(row=2, column=idx, value=val)
        hcell.font = font_white_bold
        hcell.fill = header_fill
        hcell.alignment = Alignment(horizontal="center", vertical="center")
        hcell.border = thin_border

    ns = neraca_saldo(df)
    r = 3
    for _, row in ns.reset_index().iterrows():
        ws4.cell(row=r, column=1, value=row["Akun"]).alignment = Alignment(horizontal="left")
        ws4.cell(row=r, column=2, value=format_rupiah_angka(row["Debit"])).alignment = Alignment(horizontal="right")
        ws4.cell(row=r, column=3, value=format_rupiah_angka(row["Kredit"])).alignment = Alignment(horizontal="right")
        ws4.cell(row=r, column=4, value=format_rupiah_angka(row["Saldo"])).alignment = Alignment(horizontal="right")

        for col in range(1, 5):
            ws4.cell(row=r, column=col).border = thin_border
        r += 1

    col_widths_ns = [22, 20, 20, 20]
    for i, width in enumerate(col_widths_ns, 1):
        ws4.column_dimensions[chr(64 + i)].width = width

    # Sheet 5: Laporan Laba Rugi
    ws5 = wb.create_sheet("Laporan Laba Rugi")
    ws5.merge_cells(start_row=1, start_column=1, end_row=1, end_column=2)
    c = ws5.cell(row=1, column=1, value="Laporan Laba Rugi")
    c.font = font_bold
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = year_fill

    headers = ["Keterangan", "Jumlah"]
    for idx, val in enumerate(headers, start=1):
        hcell = ws5.cell(row=2, column=idx, value=val)
        hcell.font = font_white_bold
        hcell.fill = header_fill
        hcell.alignment = Alignment(horizontal="center", vertical="center")
        hcell.border = thin_border

    lr = laporan_laba_rugi(df)
    labels = ["Total Pendapatan", "Total Beban", "Laba/Rugi"]
    values = [lr["Total Pendapatan"], lr["Total Beban"], lr["Laba/Rugi"]]

    r = 3
    for label, val in zip(labels, values):
        ws5.cell(row=r, column=1, value=label).alignment = Alignment(horizontal="left")
        ws5.cell(row=r, column=1).border = thin_border

        if label == "Laba/Rugi" and val < 0:
            val_str = f"(Rp {abs(val):,.2f})"
            val_str = val_str.replace(",", "X").replace(".", ",").replace("X", ".")
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

# ===========================
# Menu Navigasi Streamlit
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
], label_visibility="collapsed")

# Sidebar Statistik
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

# ===========================
# Semua menu dan implementasi opsi lengkap
# ===========================

if menu == "🏠 Dashboard":
    st.markdown("<div class='subtitle'>🏠 Dashboard Overview</div>", unsafe_allow_html=True)
    if total_transaksi == 0:
        st.info("👋 Mulai dengan menambahkan transaksi di menu Input Transaksi.")
    else:
        df = pd.DataFrame(st.session_state.transaksi)
        lr = laporan_laba_rugi(df)

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("📊 Total Transaksi", total_transaksi)
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

elif menu == "📝 Input Transaksi":
    st.markdown("<div class='subtitle'>📝 Input Transaksi Baru</div>", unsafe_allow_html=True)

    with st.form("form_transaksi", clear_on_submit=True):
        # Pengguna isi periode sendiri
        periode = st.text_input("🗓️ Periode (YYYY-MM)", value=datetime.now().strftime("%Y-%m"))
        try:
            tahun_input, bulan_input = map(int, periode.split("-"))
        except:
            st.warning("Format periode harus YYYY-MM, misal 2025-12")
            st.stop()

        # Input tanggal saja
        tanggal_input = st.date_input(
            "📅 Tanggal Transaksi",
            value=datetime(tahun_input, bulan_input, 1).date()
        )

        akun = st.selectbox("🏦 Pilih Akun", [
            "Kas", "Piutang", "Modal", "Pendapatan Jasa", "Pendapatan Lainnya", 
            "Beban Gaji", "Beban Listrik", "Beban Sewa", "Beban Lainnya"])
        ket = st.text_input("📝 Keterangan", "")
        debit = st.number_input("Debit (Rp)", min_value=0, step=10000, format="%d")
        kredit = st.number_input("Kredit (Rp)", min_value=0, step=10000, format="%d")
        submitted = st.form_submit_button("✅ Simpan Transaksi")

        if submitted:
            if debit == 0 and kredit == 0:
                st.error("❌ Debit atau Kredit harus diisi!")
            elif not ket.strip():
                st.error("❌ Keterangan harus diisi!")
            else:
                tgl_waktu = datetime.combine(tanggal_input, datetime.now().time())

                tambah_transaksi({
                    "Tanggal": tgl_waktu,
                    "Tahun": tahun_input,
                    "Bulan": bulan_input,
                    "Akun": akun,
                    "Keterangan": ket,
                    "Debit": int(debit),
                    "Kredit": int(kredit)
                })
                st.success("✅ Transaksi berhasil ditambahkan!")
                st.balloons()
                st.rerun()

elif menu == "📋 Lihat Transaksi":
    st.markdown("<div class='subtitle'>📋 Daftar Semua Transaksi</div>", unsafe_allow_html=True)
    if len(st.session_state.transaksi) == 0:
        st.info("Belum ada transaksi.")
    else:
        df = pd.DataFrame(st.session_state.transaksi)
        filter_akun = st.multiselect("Filter Akun", df["Akun"].unique())
        sort_by = st.selectbox("Urutkan berdasarkan", ["Tanggal", "Akun", "Debit", "Kredit"])
        if filter_akun:
            df = df[df["Akun"].isin(filter_akun)]
        df = df.sort_values(sort_by)

        df_display = df.copy()
        df_display["Tanggal"] = df_display["Tanggal"].apply(format_tanggal)
        df_display["Debit"] = df_display["Debit"].apply(format_rupiah_angka)
        df_display["Kredit"] = df_display["Kredit"].apply(format_rupiah_angka)
        st.dataframe(df_display, use_container_width=True)

        idx_hapus = st.number_input("Nomor indeks hapus transaksi", min_value=0, max_value=len(st.session_state.transaksi)-1)
        if st.button("🗑️ Hapus"):
            hapus_transaksi(idx_hapus)
            st.success("Transaksi berhasil dihapus")
            st.rerun()

elif menu == "📖 Buku Besar":
    st.markdown("<div class='subtitle'>📖 Buku Besar Per Akun</div>", unsafe_allow_html=True)
    if len(st.session_state.transaksi) == 0:
        st.info("Belum ada transaksi.")
    else:
        df = pd.DataFrame(st.session_state.transaksi)
        bb = buku_besar(df)
        for idx, (akun, data) in enumerate(bb.items()):
            with st.expander(f"📊 {akun}", expanded=(idx == 0)):
                d = data.copy()
                d["Tanggal"] = d["Tanggal"].apply(format_tanggal)
                d["Debit"] = d["Debit"].apply(format_rupiah_angka)
                d["Kredit"] = d["Kredit"].apply(format_rupiah_angka)
                d["Saldo"] = d["Saldo"].apply(format_rupiah_angka)
                st.dataframe(d, use_container_width=True, hide_index=True)

elif menu == "⚖️ Neraca Saldo":
    st.markdown("<div class='subtitle'>⚖️ Neraca Saldo</div>", unsafe_allow_html=True)
    if len(st.session_state.transaksi) == 0:
        st.info("Belum ada transaksi.")
    else:
        df = pd.DataFrame(st.session_state.transaksi)
        ns = neraca_saldo(df)
        ns_display = ns.copy()
        ns_display["Debit"] = ns_display["Debit"].apply(format_rupiah_angka)
        ns_display["Kredit"] = ns_display["Kredit"].apply(format_rupiah_angka)
        ns_display["Saldo"] = ns_display["Saldo"].apply(format_rupiah_angka)
        st.dataframe(ns_display, use_container_width=True)

elif menu == "💰 Laporan Laba Rugi":
    st.markdown("<div class='subtitle'>💰 Laporan Laba Rugi</div>", unsafe_allow_html=True)
    if len(st.session_state.transaksi) == 0:
        st.info("Belum ada transaksi.")
    else:
        df = pd.DataFrame(st.session_state.transaksi)
        lr = laporan_laba_rugi(df)
        col1, col2, col3 = st.columns(3)
        col1.markdown(f"<div style='background:#11998e; padding:25px; border-radius:12px; color:#fff; text-align:center;'>\
            <h3>💵 Total Pendapatan</h3><h2>{format_rupiah_angka(lr['Total Pendapatan'])}</h2></div>", unsafe_allow_html=True)
        col2.markdown(f"<div style='background:#ee0979; padding:25px; border-radius:12px; color:#fff; text-align:center;'>\
            <h3>💸 Total Beban</h3><h2>{format_rupiah_angka(lr['Total Beban'])}</h2></div>", unsafe_allow_html=True)
        laba = lr["Laba/Rugi"]
        if laba >= 0:
            col3.markdown(f"<div style='background:#56ccf2; padding:25px; border-radius:12px; color:#fff; text-align:center;'>\
            <h3>✅ Laba Bersih</h3><h2>{format_rupiah_angka(laba)}</h2></div>", unsafe_allow_html=True)
        else:
            col3.markdown(f"<div style='background:#f2709c; padding:25px; border-radius:12px; color:#fff; text-align:center;'>\
            <h3>⚠️ Rugi Bersih</h3><h2>{format_rupiah_angka(abs(laba))}</h2></div>", unsafe_allow_html=True)

elif menu == "📈 Grafik":
    st.markdown("<div class='subtitle'>📈 Visualisasi Data Akuntansi</div>", unsafe_allow_html=True)
    if len(st.session_state.transaksi) == 0:
        st.info("Belum ada data.")
    else:
        df = pd.DataFrame(st.session_state.transaksi)
        tab1, tab2, tab3 = st.tabs(["📊 Debit per Akun", "📊 Kredit per Akun", "📊 Perbandingan"])
        with tab1:
            chart = alt.Chart(df).mark_bar().encode(
                x=alt.X("Akun:N", title="Akun"),
                y=alt.Y("Debit:Q", title="Debit (Rp)"),
                color=alt.Color("Akun:N", legend=None),
                tooltip=["Akun", "Debit"]
            ).properties(title="Grafik Total Debit per Akun", height=400)
            st.altair_chart(chart, use_container_width=True)
        with tab2:
            chart = alt.Chart(df).mark_bar().encode(
                x=alt.X("Akun:N", title="Akun"),
                y=alt.Y("Kredit:Q", title="Kredit (Rp)"),
                color=alt.Color("Akun:N", legend=None),
                tooltip=["Akun", "Kredit"]
            ).properties(title="Grafik Total Kredit per Akun", height=400)
            st.altair_chart(chart, use_container_width=True)
        with tab3:
            df_grouped = df.groupby("Akun")[["Debit", "Kredit"]].sum().reset_index()
            df_melt = df_grouped.melt(id_vars="Akun", value_vars=["Debit", "Kredit"], var_name="Tipe", value_name="Jumlah")
            chart = alt.Chart(df_melt).mark_bar().encode(
                x=alt.X("Akun:N", title="Akun"),
                y=alt.Y("Jumlah:Q", title="Jumlah (Rp)"),
                color="Tipe:N",
                xOffset="Tipe:N",
                tooltip=["Akun", "Tipe", "Jumlah"]
            ).properties(title="Perbandingan Debit vs Kredit per Akun", height=400)
            st.altair_chart(chart, use_container_width=True)

elif menu == "📥 Import Excel":
    st.markdown("<div class='subtitle'>📥 Import Transaksi dari File Excel</div>", unsafe_allow_html=True)
    uploaded_file = st.file_uploader("Pilih file Excel", type=["xlsx"])
    if uploaded_file:
        df_import = pd.read_excel(uploaded_file)
        df_import.columns = df_import.columns.str.strip()  # bersihkan nama kolom
        expected_cols = ["Tanggal", "Akun", "Keterangan", "Debit", "Kredit"]
    
    if not all(col in df_import.columns for col in expected_cols):
        st.error(f"File harus ada kolom: {expected_cols}")
    else:
        df_import["Tanggal"] = pd.to_datetime(df_import["Tanggal"], errors="coerce")
        df_import = df_import.dropna(subset=["Tanggal"])

        # bersihkan angka
        for col in ["Debit", "Kredit"]:
            df_import[col] = (
                df_import[col].astype(str)
                .str.replace("Rp", "")
                .str.replace(".", "")
                .str.replace(",", ".")
            )
            df_import[col] = pd.to_numeric(df_import[col], errors="coerce").fillna(0).astype(int)

        # tambah Tahun & Bulan
        df_import["Tahun"] = df_import["Tanggal"].dt.year
        df_import["Bulan"] = df_import["Tanggal"].dt.month

        st.dataframe(df_import.head())

        if st.button("Tambahkan semua transaksi dari file"):
            for _, row in df_import.iterrows():
                tambah_transaksi(
                    row["Tanggal"], row["Akun"], row["Keterangan"], row["Debit"], row["Kredit"]
                )
            st.success(f"Berhasil menambahkan {len(df_import)} transaksi!")
            st.rerun()

elif menu == "📤 Export Excel":
    df = pd.DataFrame(st.session_state.transaksi)

    st.markdown("<div class='subtitle'>📤 Export Laporan ke Excel</div>", unsafe_allow_html=True)

    if len(df) == 0:
        st.info("Belum ada transaksi untuk diexport.")
    else:
        # Filter periode sebelum export
        periode_filter = st.text_input("Filter Periode (YYYY-MM)", value=datetime.now().strftime("%Y-%m"))
        try:
            tahun_filter, bulan_filter = map(int, periode_filter.split("-"))
            df_filtered = df[(df["Tahun"] == tahun_filter) & (df["Bulan"] == bulan_filter)]
        except:
            st.warning("Format periode harus YYYY-MM, misal 2025-12")
            st.stop()

        if df_filtered.empty:
            st.info("Tidak ada transaksi di periode ini.")
        else:
            total_debit = df_filtered["Debit"].sum()
            total_kredit = df_filtered["Kredit"].sum()
            saldo = total_debit - total_kredit
            st.info(f"Total Debit: {format_rupiah_angka(total_debit)} | Total Kredit: {format_rupiah_angka(total_kredit)} | Saldo: {format_rupiah_angka(saldo)}")

            try:
                excel_data = export_excel_multi(df_filtered)  # gunakan fungsi export_excel_multi yang sudah ada
                st.download_button(
                    "Download Laporan Akuntansi.xlsx",
                    excel_data,
                    file_name=f"laporan_akuntansi_{periode_filter.replace('-', '')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.success("File siap diunduh!")
            except Exception as e:
                st.error(f"Error saat generate file Excel: {e}")

# ===================
# Footer
# ===================
st.markdown("---")
st.markdown("""
<div style='text-align:center; color:#888; padding:20px;'>
    <p>💰 <strong>Aplikasi Akuntansi Profesional</strong></p>
    <p>Kelola keuangan bisnis Anda dengan mudah dan efisien</p>
</div>
""", unsafe_allow_html=True)






