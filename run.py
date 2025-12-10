import streamlit as st
import pandas as pd
import altair as alt
from datetime import datetime
import io
import calendar
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment

# ============================
# CONFIG TAMPAK APLIKASI
# ============================
st.set_page_config(
    page_title="Aplikasi Akuntansi",
    page_icon="💰",
    layout="wide"
)

# CSS untuk UI modern + warna teks terlihat
st.markdown("""
<style>
    .title { font-size: 38px; font-weight: 800; color: #1a237e; text-align:center; }
    .subtitle { font-size: 22px; font-weight: 600; color:#1a237e; margin-top: 10px; }
    .stButton>button {
        background-color: #1a237e !important;
        color: white !important;
        padding: 10px 20px;
        border-radius: 10px;
        font-size: 17px;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("<div class='title'>📊 Aplikasi Akuntansi</div>", unsafe_allow_html=True)
st.write("")

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
#  FUNGSI AKUNTANSI
# ============================
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
        df_akun = df[df["Akun"] == akun].copy()
        df_akun["Saldo"] = df_akun["Debit"].cumsum() - df_akun["Kredit"].cumsum()
        buku_besar_data[akun] = df_akun
    return buku_besar_data

def neraca_saldo(df):
    grouped = df.groupby("Akun")[["Debit", "Kredit"]].sum()
    grouped["Saldo"] = grouped["Debit"] - grouped["Kredit"]
    return grouped

# ============================
# SIDEBAR MENU
# ============================
menu = st.sidebar.radio(
    "📌 PILIH MENU",
    ["Input Transaksi", "Jurnal Umum", "Buku Besar", "Neraca Saldo", "Grafik", "Export Excel"]
)

# ============================
# 1. INPUT TRANSAKSI
# ============================
if menu == "Input Transaksi":
    st.markdown("<div class='subtitle'>📝 Input Transaksi</div>", unsafe_allow_html=True)

    akun_list = [
        "Kas", "Piutang", "Utang", "Modal", "Pendapatan Jasa",
        "Beban Gaji", "Beban Listrik", "Beban Sewa"
    ]

    tanggal = st.date_input("Tanggal", datetime.now())
    akun = st.selectbox("Akun", akun_list)
    ket = st.text_input("Keterangan")

    col1, col2 = st.columns(2)
    with col1:
        debit = st.number_input("Debit (Rp)", min_value=0, step=1000, format="%d")
    with col2:
        kredit = st.number_input("Kredit (Rp)", min_value=0, step=1000, format="%d")

    if st.button("Tambah Transaksi"):
        tambah_transaksi(str(tanggal), akun, ket, debit, kredit)
        st.success("Transaksi berhasil ditambahkan!")

    st.write("### 📄 Daftar Transaksi")

    if len(st.session_state.transaksi) > 0:
        df = pd.DataFrame(st.session_state.transaksi)
        df_display = df.copy()
        df_display["Debit"] = df_display["Debit"].apply(to_rp)
        df_display["Kredit"] = df_display["Kredit"].apply(to_rp)
        st.dataframe(df_display, use_container_width=True)

        idx = st.number_input("Hapus transaksi index", 0, len(df)-1)
        if st.button("Hapus"):
            hapus_transaksi(idx)
            st.warning("Transaksi berhasil dihapus!")
    else:
        st.info("Belum ada transaksi.")

# ============================
# 2. JURNAL UMUM
# ============================
elif menu == "Jurnal Umum":
    st.markdown("<div class='subtitle'>📘 Jurnal Umum</div>", unsafe_allow_html=True)

    if len(st.session_state.transaksi) == 0:
        st.info("Belum ada data.")
    else:
        df = pd.DataFrame(st.session_state.transaksi)
        df2 = df.copy()
        df2["Debit"] = df2["Debit"].apply(to_rp)
        df2["Kredit"] = df2["Kredit"].apply(to_rp)
        st.dataframe(df2, use_container_width=True)

# ============================
# 3. BUKU BESAR
# ============================
elif menu == "Buku Besar":
    st.markdown("<div class='subtitle'>📗 Buku Besar</div>", unsafe_allow_html=True)

    if len(st.session_state.transaksi) == 0:
        st.info("Belum ada data.")
    else:
        df = pd.DataFrame(st.session_state.transaksi)
        buku = buku_besar(df)

        for akun, data in buku.items():
            st.write(f"### ▶ {akun}")
            df2 = data.copy()
            df2["Debit"] = df2["Debit"].apply(to_rp)
            df2["Kredit"] = df2["Kredit"].apply(to_rp)
            df2["Saldo"] = df2["Saldo"].apply(to_rp)
            st.dataframe(df2, use_container_width=True)

# ============================
# 4. NERACA SALDO
# ============================
elif menu == "Neraca Saldo":
    st.markdown("<div class='subtitle'>📙 Neraca Saldo</div>", unsafe_allow_html=True)

    if len(st.session_state.transaksi) == 0:
        st.info("Belum ada data.")
    else:
        df = pd.DataFrame(st.session_state.transaksi)
        neraca = neraca_saldo(df)
        df2 = neraca.copy()
        df2["Debit"] = df2["Debit"].apply(to_rp)
        df2["Kredit"] = df2["Kredit"].apply(to_rp)
        df2["Saldo"] = df2["Saldo"].apply(to_rp)
        st.dataframe(df2, use_container_width=True)

# ============================
# 5. GRAFIK
# ============================
elif menu == "Grafik":
    st.markdown("<div class='subtitle'>📈 Grafik Akuntansi</div>", unsafe_allow_html=True)

    if len(st.session_state.transaksi) == 0:
        st.info("Belum ada data.")
    else:
        df = pd.DataFrame(st.session_state.transaksi)
        chart = alt.Chart(df).mark_bar().encode(
            x="Akun",
            y="Debit",
            color="Akun"
        ).properties(
            title="Grafik Jumlah Debit per Akun",
            width=700
        )
        st.altair_chart(chart, use_container_width=True)

# ============================
# 6. EXPORT EXCEL MULTI-SHEET DENGAN PEMBATAS BULAN
# ============================
def export_excel_multi(df):
    output = io.BytesIO()
    wb = Workbook()

    # =============================
    #   SETUP DATA BULAN & TAHUN
    # =============================
    df["Tanggal"] = pd.to_datetime(df["Tanggal"])
    df["Bulan"] = df["Tanggal"].dt.month
    df["Tahun"] = df["Tanggal"].dt.year
    df_sorted = df.sort_values("Tanggal")

    bulan_awal = calendar.month_name[int(df_sorted["Bulan"].iloc[0])]
    tahun_awal = df_sorted["Tahun"].iloc[0]
    bulan_akhir = calendar.month_name[int(df_sorted["Bulan"].iloc[-1])]
    tahun_akhir = df_sorted["Tahun"].iloc[-1]

    periode_text = f"Periode: {bulan_awal} {tahun_awal} - {bulan_akhir} {tahun_akhir}"

    # =====================================================
    #  SHEET 1: JURNAL UMUM
    # =====================================================
    ws1 = wb.active
    ws1.title = "Jurnal Umum"

    ws1.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
    header_cell = ws1.cell(row=1, column=1, value=periode_text)
    header_cell.font = Font(bold=True, size=13)
    header_cell.alignment = Alignment(horizontal="center")

    current_row = 3

    for (tahun, bulan), grup in df_sorted.groupby(["Tahun", "Bulan"]):
        nama_bulan = calendar.month_name[bulan].upper()

        ws1.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=6)
        cell = ws1.cell(row=current_row, column=1, value=f"=== {nama_bulan} {tahun} ===")
        cell.font = Font(bold=True, size=12)
        current_row += 1

        headers = list(grup.columns.drop(["Bulan", "Tahun"]))
        for col_num, header in enumerate(headers, start=1):
            ws1.cell(row=current_row, column=col_num, value=header).font = Font(bold=True)
        current_row += 1

        for r in dataframe_to_rows(grup.drop(["Bulan", "Tahun"], axis=1), index=False, header=False):
            for c_idx, val in enumerate(r, start=1):
                cell = ws1.cell(row=current_row, column=c_idx, value=val)
                if c_idx in [4, 5] and isinstance(val, (int, float)):
                    cell.number_format = '"Rp"#,##0'
            current_row += 1

        current_row += 2

    # =====================================================
    #  SHEET 2: BUKU BESAR
    # =====================================================
    ws2 = wb.create_sheet("Buku Besar")

    ws2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
    h2 = ws2.cell(row=1, column=1, value=periode_text)
    h2.font = Font(bold=True, size=13)
    h2.alignment = Alignment(horizontal="center")

    row_buku = 3
    buku = buku_besar(df)

    for akun, data in buku.items():
        ws2.merge_cells(start_row=row_buku, start_column=1, end_row=row_buku, end_column=6)
        ws2.cell(row=row_buku, column=1, value=f"== {akun.upper()} ==").font = Font(bold=True, size=12)
        row_buku += 1

        for col_num, header in enumerate(data.columns, start=1):
            ws2.cell(row=row_buku, column=col_num, value=header).font = Font(bold=True)
        row_buku += 1

        for r in dataframe_to_rows(data, index=False, header=False):
            for c_idx, val in enumerate(r, start=1):
                cell = ws2.cell(row=row_buku, column=c_idx, value=val)
                if c_idx in [4, 5, 6] and isinstance(val, (int, float)):
                    cell.number_format = '"Rp"#,##0'
            row_buku += 1

        row_buku += 2

    # =====================================================
    #  SHEET 3: NERACA SALDO
    # =====================================================
    ws3 = wb.create_sheet("Neraca Saldo")

    ws3.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
    h3 = ws3.cell(row=1, column=1, value=periode_text)
    h3.font = Font(bold=True, size=13)
    h3.alignment = Alignment(horizontal="center")

    neraca = neraca_saldo(df).reset_index()

    row = 3
    for r in dataframe_to_rows(neraca, index=False, header=True):
        for c_idx, val in enumerate(r, start=1):
            cell = ws3.cell(row=row, column=c_idx, value=val)
            if row == 3:
                cell.font = Font(bold=True)
            elif c_idx in [2, 3, 4] and isinstance(val, (int, float)):
                cell.number_format = '"Rp"#,##0'
        row += 1

    wb.save(output)
    output.seek(0)
    return output.getvalue()

    if len(st.session_state.transaksi) == 0:
        st.info("Belum ada transaksi untuk diekspor.")
    else:
        df = pd.DataFrame(st.session_state.transaksi)
        excel_file = export_excel_multi(df)
        st.download_button(
            label="📥 Export ke Excel (Lengkap)",
            data=excel_file,
            file_name="laporan_akuntansi_lengkap.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

