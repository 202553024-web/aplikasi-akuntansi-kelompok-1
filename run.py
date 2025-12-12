import streamlit as st
import pandas as pd
import altair as alt
from datetime import datetime
import io
import calendar
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, PatternFill

# ============================
# CONFIG TAMPAK APLIKASI
# ============================
st.set_page_config(
    page_title="Aplikasi Akuntansi",
    page_icon="💰",
    layout="wide"
)

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
# FUNGSI AKUNTANSI
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
# FUNGSI EXPORT EXCEL (REVISI SESUAI GAMBAR)
# ============================
def export_excel_multi(df):
    import io, calendar
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill
    from openpyxl.utils import get_column_letter
    from openpyxl.utils.dataframe import dataframe_to_rows

    output = io.BytesIO()
    wb = Workbook()
    ws_main = wb.active
    ws_main.title = "Laporan Keuangan"

    # Persiapan Data
    df["Tanggal"] = pd.to_datetime(df["Tanggal"])
    df["Bulan"] = df["Tanggal"].dt.month
    df["Tahun"] = df["Tanggal"].dt.year
    df_sorted = df.sort_values("Tanggal")

    # Warna untuk header
    green_fill = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    current_row = 1
    tahun_sekarang = None

    # ============================
    # HALAMAN LAPORAN KEUANGAN
    # ============================
    for (tahun, bulan), grup in df_sorted.groupby(["Tahun", "Bulan"]):
        # Header Tahun (Hijau)
        if tahun != tahun_sekarang:
            ws_main.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=5)
            cell = ws_main.cell(row=current_row, column=1, value=f"Laporan Keuangan Tahun {tahun}")
            cell.font = Font(bold=True, size=14)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.fill = green_fill
            current_row += 1
            tahun_sekarang = tahun

        # Header Bulan (Kuning)
        nama_bulan = calendar.month_name[bulan].capitalize()
        ws_main.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=5)
        cell = ws_main.cell(row=current_row, column=1, value=f"Bulan {nama_bulan}")
        cell.font = Font(bold=True, size=12)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = yellow_fill
        current_row += 1

        # Header Kolom (Kuning)
        headers = ["Tanggal", "Akun", "Keterangan", "Debit", "Kredit"]
        for col_num, header in enumerate(headers, start=1):
            cell = ws_main.cell(row=current_row, column=col_num, value=header)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.fill = yellow_fill
        current_row += 1
        
        # Data Transaksi
        for r in dataframe_to_rows(grup[["Tanggal", "Akun", "Keterangan", "Debit", "Kredit"]], index=False, header=False):
            for c_idx, val in enumerate(r, start=1):
                cell = ws_main.cell(row=current_row, column=c_idx)
                if c_idx == 1:  # Tanggal
                    if isinstance(val, datetime):
                        cell.value = val.strftime("%Y-%m-%d %H:%M:%S")
                    else:
                        cell.value = val
                    cell.alignment = Alignment(horizontal="left")
                elif c_idx in [4, 5]:  # Debit dan Kredit
                    val = int(val) if pd.notna(val) and val != 0 else 0
                    if val == 0:
                        cell.value = "Rp                    -"
                        cell.alignment = Alignment(horizontal="right")
                    else:
                        cell.value = val
                        cell.alignment = Alignment(horizontal="right")
                        cell.number_format = '"Rp"#,##0'
                else:
                    cell.value = val
                    cell.alignment = Alignment(horizontal="left")
            current_row += 1

    # Set Lebar Kolom
    ws_main.column_dimensions['A'].width = 20
    ws_main.column_dimensions['B'].width = 15
    ws_main.column_dimensions['C'].width = 20
    ws_main.column_dimensions['D'].width = 18
    ws_main.column_dimensions['E'].width = 18


    # ============================
    # SHEET JURNAL UMUM
    # ============================
    ws_jurnal = wb.create_sheet("Jurnal Umum")
    headers = ["Tanggal", "Akun", "Keterangan", "Debit", "Kredit"]
    for col_num, header in enumerate(headers, start=1):
        ws_jurnal.cell(row=1, column=col_num, value=header).font = Font(bold=True)
    for i, r in enumerate(dataframe_to_rows(df[headers], index=False, header=False), start=2):
        for c_idx, val in enumerate(r, start=1):
            cell = ws_jurnal.cell(row=i, column=c_idx)
            if c_idx in [4, 5]:
                val = int(val) if pd.notna(val) else 0
                cell.value = val
                cell.alignment = Alignment(horizontal="right")
                cell.number_format = '"Rp"#,##0'
            else:
                cell.value = val


    # ============================
    # SHEET BUKU BESAR
    # ============================
    ws_bb = wb.create_sheet("Buku Besar")
    akun_list = df["Akun"].unique()
    row_bb = 1

    for akun in akun_list:
        ws_bb.cell(row=row_bb, column=1, value=f"Akun: {akun}").font = Font(bold=True, size=12)
        row_bb += 1

        df_akun = df[df["Akun"] == akun].copy()
        df_akun["Saldo"] = df_akun["Debit"].cumsum() - df_akun["Kredit"].cumsum()
        headers = ["Tanggal", "Keterangan", "Debit", "Kredit", "Saldo"]

        for col_num, header in enumerate(headers, start=1):
            ws_bb.cell(row=row_bb, column=col_num, value=header).font = Font(bold=True)
        row_bb += 1

        for r in dataframe_to_rows(df_akun[headers], index=False, header=False):
            for c_idx, val in enumerate(r, start=1):
                cell = ws_bb.cell(row=row_bb, column=c_idx)
                if c_idx >= 3:
                    val = int(val) if pd.notna(val) else 0
                    cell.value = val
                    cell.alignment = Alignment(horizontal="right")
                    cell.number_format = '"Rp"#,##0'
                else:
                    cell.value = val
            row_bb += 1
        row_bb += 2


    # ============================
    # SHEET NERACA SALDO
    # ============================
    ws_ns = wb.create_sheet("Neraca Saldo")
    neraca = df.groupby("Akun")[["Debit", "Kredit"]].sum().reset_index()
    neraca["Saldo"] = neraca["Debit"] - neraca["Kredit"]

    headers = ["Akun", "Debit", "Kredit", "Saldo"]
    for col_num, header in enumerate(headers, start=1):
        ws_ns.cell(row=1, column=col_num, value=header).font = Font(bold=True)
    for i, r in enumerate(dataframe_to_rows(neraca[headers], index=False, header=False), start=2):
        for c_idx, val in enumerate(r, start=1):
            cell = ws_ns.cell(row=i, column=c_idx)
            if c_idx >= 2:
                val = int(val) if pd.notna(val) else 0
                cell.value = val
                cell.alignment = Alignment(horizontal="right")
                cell.number_format = '"Rp"#,##0'
            else:
                cell.value = val

    wb.save(output)
    output.seek(0)
    return output.getvalue()

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
