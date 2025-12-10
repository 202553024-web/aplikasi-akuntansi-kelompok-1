import streamlit as st
import pandas as pd
import altair as alt
from datetime import datetime
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ============================
# CONFIG TAMPAK APLIKASI
# ============================
st.set_page_config(page_title="Aplikasi Akuntansi", page_icon="💰", layout="wide")

# CSS UI Modern
st.markdown("""
<style>
    .title { font-size: 40px; font-weight: 900; color: #1a237e; text-align:center; }
    .subtitle { font-size: 24px; font-weight: 700; color:#1a237e; margin-top: 20px; }
    div[data-testid="stForm"] {
        background: #f4f6ff;
        padding: 20px;
        border-radius: 12px;
        border: 1px solid #d0d4f0;
    }
    .stButton>button {
        background-color: #1a237e !important;
        color: white !important;
        padding: 10px 22px;
        border-radius: 10px;
        font-size: 17px;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("<div class='title'>📊 Aplikasi Akuntansi Modern</div>", unsafe_allow_html=True)
st.write("")

# ============================
# SESSION
# ============================
if "transaksi" not in st.session_state:
    st.session_state.transaksi = []

# Format Rupiah
def to_rp(n):
    try:
        return "Rp {:,}".format(int(n)).replace(",", ".")
    except:
        return "Rp 0"

# ============================
# FUNGSI
# ============================
def tambah_transaksi(tgl, akun, ket, debit, kredit):
    st.session_state.transaksi.append({
        "Tanggal": tgl,
        "Akun": akun,
        "Keterangan": ket,
        "Debit": int(debit),
        "Kredit": int(kredit),
        "Bulan": pd.to_datetime(tgl).strftime("%B %Y")
    })

def buku_besar(df):
    akun_list = df["Akun"].unique()
    buku_besar_data = {}
    for akun in akun_list:
        df_akun = df[df["Akun"] == akun].copy()
        df_akun["Saldo"] = df_akun["Debit"].cumsum() - df_akun["Kredit"].cumsum()
        buku_besar_data[akun] = df_akun
    return buku_besar_data

def neraca_saldo(df):
    grouped = df.groupby(["Akun", "Bulan"])[["Debit", "Kredit"]].sum().reset_index()
    grouped["Saldo"] = grouped["Debit"] - grouped["Kredit"]
    return grouped

# ============================
# MENU SIDEBAR
# ============================
menu = st.sidebar.radio(
    "📌 PILIH MENU",
    ["Input Transaksi", "Jurnal Umum", "Buku Besar", "Neraca Saldo", "Grafik", "Export Excel"]
)

# ============================
# INPUT TRANSAKSI
# ============================
if menu == "Input Transaksi":
    st.markdown("<div class='subtitle'>📝 Input Transaksi</div>", unsafe_allow_html=True)

    akun_list = ["Kas", "Piutang", "Utang", "Modal", "Pendapatan Jasa",
                 "Beban Gaji", "Beban Listrik", "Beban Sewa"]

    with st.form("input_form"):
        tanggal = st.date_input("Tanggal", datetime.now())
        akun = st.selectbox("Akun", akun_list)
        ket = st.text_input("Keterangan")

        col1, col2 = st.columns(2)
        with col1:
            debit = st.number_input("Debit (Rp)", min_value=0, step=1000, format="%d")
        with col2:
            kredit = st.number_input("Kredit (Rp)", min_value=0, step=1000, format="%d")

        submit = st.form_submit_button("Tambah Transaksi")

    if submit:
        tambah_transaksi(str(tanggal), akun, ket, debit, kredit)
        st.success("Transaksi berhasil ditambahkan!")

    st.write("### 📄 Daftar Transaksi")
    if len(st.session_state.transaksi) > 0:
        df = pd.DataFrame(st.session_state.transaksi)
        df_display = df.copy()
        df_display["Debit"] = df_display["Debit"].apply(to_rp)
        df_display["Kredit"] = df_display["Kredit"].apply(to_rp)
        st.dataframe(df_display, use_container_width=True)

# ============================
# EXPORT EXCEL
# ============================
def auto_fit(ws):
    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                length = len(str(cell.value))
                if length > max_length:
                    max_length = length
            except:
                pass
        ws.column_dimensions[col_letter].width = max_length + 3

def export_excel_multi(df):
    output = io.BytesIO()
    wb = Workbook()

    # -----------------------------
    # SHEET 1 - JURNAL UMUM
    # -----------------------------
    ws = wb.active
    ws.title = "Jurnal Umum"

    months = df["Bulan"].unique()

    row = 1
    for bln in months:
        ws.cell(row=row, column=1, value=bln).font = Font(bold=True, size=13)
        row += 2

        df_bln = df[df["Bulan"] == bln]
        ws.append(list(df_bln.columns))

        for _, r in df_bln.iterrows():
            ws.append(list(r))

        row = ws.max_row + 3

    auto_fit(ws)

    # -----------------------------
    # SHEET 2 - Buku Besar
    # -----------------------------
    ws2 = wb.create_sheet("Buku Besar")

    row = 1
    buku = buku_besar(df)

    for akun, data in buku.items():
        ws2.cell(row=row, column=1, value=akun).font = Font(bold=True, size=13)
        row += 2
        ws2.append(list(data.columns))
        for _, r in data.iterrows():
            ws2.append(list(r))
        row = ws2.max_row + 3

    auto_fit(ws2)

    # -----------------------------
    # SHEET 3 - Neraca Saldo
    # -----------------------------
    ws3 = wb.create_sheet("Neraca Saldo")

    ner = neraca_saldo(df)
    ws3.append(list(ner.columns))
    for _, r in ner.iterrows():
        ws3.append(list(r))

    auto_fit(ws3)

    wb.save(output)
    return output.getvalue()

# ============================
# TOMBOL EXPORT
# ============================
elif menu == "Export Excel":
    st.markdown("<div class='subtitle'>📤 Export Excel</div>", unsafe_allow_html=True)

    if len(st.session_state.transaksi) == 0:
        st.info("Belum ada transaksi.")
    else:
        df = pd.DataFrame(st.session_state.transaksi)
        excel_file = export_excel_multi(df)

        st.download_button(
            label="📥 Export ke Excel (Dipisah Bulan dan Rapi)",
            data=excel_file,
            file_name="laporan_akuntansi.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
