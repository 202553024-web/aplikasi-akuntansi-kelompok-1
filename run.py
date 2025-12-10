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
# 6. EXPORT EXCEL
# ============================
elif menu == "Export Excel":
    st.markdown("<div class='subtitle'>📤 Export Excel</div>", unsafe_allow_html=True)

    def export_excel_multi(df):
        output = io.BytesIO()
        wb = Workbook()

        # --- SHEET 1: LAPORAN TRANSAKSI DENGAN PEMBATAS BULAN
        ws1 = wb.active
        ws1.title = "Laporan Transaksi"
        df["Tanggal"] = pd.to_datetime(df["Tanggal"])
        df["Bulan"] = df["Tanggal"].dt.month
        df["Tahun"] = df["Tanggal"].dt.year
        df_sorted = df.sort_values("Tanggal")

        current_row = 1
        for (tahun, bulan), grup in df_sorted.groupby(["Tahun", "Bulan"]):
            nama_bulan = calendar.month_name[bulan].upper()
            ws1.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=5)
            cell = ws1.cell(row=current_row, column=1, value=f"=== LAPORAN BULAN {nama_bulan} {tahun} ===")
            cell.font = Font(bold=True, size=12)
            current_row += 1

            headers = list(grup.columns.drop(["Bulan", "Tahun"]))
            for col_num, header in enumerate(headers, start=1):
                ws1.cell(row=current_row, column=col_num, value=header).font = Font(bold=True)
            current_row += 1

            for r in dataframe_to_rows(grup.drop(["Bulan", "Tahun"], axis=1), index=False, header=False):
                for c_idx, val in enumerate(r, start=1):
                    cell = ws1.cell(row=current_row, column=c_idx, value=val)
                    if c_idx in [4, 5]:
                        if isinstance(val, (int, float)):
                            cell.number_format = '"Rp"#,##0'
                    cell.alignment = Alignment(horizontal="left")
                current_row += 1
            current_row += 2

        # --- SHEET 2: BUKU BESAR
        for akun, data in buku_besar(df).items():
            ws_buku = wb.create_sheet(f"Buku Besar - {akun}")
            for r_idx, row in enumerate(dataframe_to_rows(data, index=False, header=True), start=1):
                for c_idx, val in enumerate(row, start=1):
                    cell = ws_buku.cell(row=r_idx, column=c_idx, value=val)
                    if r_idx == 1:
                        cell.font = Font(bold=True)
                    elif c_idx in [4, 5, 6] and isinstance(val, (int, float)):
                        cell.number_format = '"Rp"#,##0'

        # --- SHEET 3: NERACA SALDO
        ws3 = wb.create_sheet("Neraca Saldo")
        neraca = neraca_saldo(df).reset_index()
        for r_idx, row in enumerate(dataframe_to_rows(neraca, index=False, header=True), start=1):
            for c_idx, val in enumerate(row, start=1):
                cell = ws3.cell(row=r_idx, column=c_idx, value=val)
                if r_idx == 1:
                    cell.font = Font(bold=True)
                elif c_idx in [2, 3, 4] and isinstance(val, (int, float)):
                    cell.number_format = '"Rp"#,##0'

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
