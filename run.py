import streamlit as st
import pandas as pd
import altair as alt
from datetime import datetime
import io
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# ============================
# CONFIG UI
# ============================
st.set_page_config(
    page_title="Aplikasi Akuntansi",
    page_icon="💰",
    layout="wide"
)

# THEME + CSS
st.markdown("""
<style>
    body { background-color: #f5f7fa; }
    .title {
        font-size: 40px; font-weight: 800; text-align:center;
        color:#1a237e; margin-bottom:15px;
    }
    .subtitle {
        font-size: 24px; font-weight: 700; color:#283593;
        margin-top:10px; margin-bottom:5px;
    }
    .stButton>button {
        background:linear-gradient(90deg,#303f9f,#1a237e) !important;
        color:white !important; padding: 10px 22px; border-radius:10px;
        font-size:17px; border:none;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("<div class='title'>📊 Aplikasi Akuntansi</div>", unsafe_allow_html=True)

# ============================
# SESSION
# ============================
if "transaksi" not in st.session_state:
    st.session_state.transaksi = []

# ============================
# FORMAT RUPIAH
# ============================
def to_rp(n):
    return f"Rp {int(n):,}".replace(",", ".")

# ============================
# FUNGSI
# ============================
def tambah_transaksi(tgl, akun, ket, debit, kredit):
    st.session_state.transaksi.append({
        "Tanggal": tgl,
        "Akun": akun,
        "Keterangan": ket,
        "Debit": debit,
        "Kredit": kredit
    })

def buku_besar(df):
    result = {}
    for akun in df["Akun"].unique():
        d = df[df["Akun"] == akun].copy()
        d["Saldo"] = d["Debit"].cumsum() - d["Kredit"].cumsum()
        result[akun] = d
    return result

def neraca_saldo(df):
    grouped = df.groupby("Akun")[["Debit", "Kredit"]].sum()
    grouped["Saldo"] = grouped["Debit"] - grouped["Kredit"]
    return grouped

# ============================
# EXCEL EXPORT
# ============================
def export_excel_multi(df):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine="openpyxl")

    # ==== SHEET 1 - JURNAL UMUM (Dipisah per bulan) ====
    df_sorted = df.sort_values("Tanggal")

    sheet_name = "Jurnal Umum"
    start = 0

    # Loop per bulan
    for bulan, d in df_sorted.groupby(df_sorted["Tanggal"].str[:7]):  # YYYY-MM
        writer.sheets.setdefault(sheet_name, writer.book.create_sheet(sheet_name))
        ws = writer.sheets[sheet_name]

        ws[f"A{start+1}"] = f"Periode: {bulan}"
        ws[f"A{start+1}"].font = Font(bold=True, size=12)

        d.to_excel(writer, index=False, sheet_name=sheet_name, startrow=start+2)
        start += len(d) + 4

    # ==== SHEET 2 - BUKU BESAR ====
    buku = buku_besar(df)
    start = 0
    for akun, d in buku.items():
        d.to_excel(writer, sheet_name="Buku Besar", startrow=start, index=False)
        start += len(d) + 3

    # ==== SHEET 3 - NERACA SALDO ====
    neraca = neraca_saldo(df)
    neraca.to_excel(writer, sheet_name="Neraca Saldo", index=True)

    writer.close()
    return output.getvalue()

# ============================
# MENU
# ============================
menu = st.sidebar.radio("📌 PILIH MENU", [
    "Input Transaksi", "Jurnal Umum", "Buku Besar", "Neraca Saldo", "Grafik", "Export Excel"
])

# ============================
# INPUT TRANSAKSI
# ============================
if menu == "Input Transaksi":
    st.markdown("<div class='subtitle'>📝 Input Transaksi</div>", unsafe_allow_html=True)

    akun_list = ["Kas","Piutang","Utang","Modal","Pendapatan Jasa","Beban Gaji","Beban Listrik","Beban Sewa"]

    tanggal = st.date_input("Tanggal", datetime.now())
    akun = st.selectbox("Akun", akun_list)
    ket = st.text_input("Keterangan")
    col1, col2 = st.columns(2)

    with col1:
        debit = st.number_input("Debit", min_value=0, step=1000)
    with col2:
        kredit = st.number_input("Kredit", min_value=0, step=1000)

    if st.button("Tambah Transaksi"):
        tambah_transaksi(str(tanggal), akun, ket, debit, kredit)
        st.success("Transaksi berhasil ditambahkan!")

    st.write("### 📄 Daftar Transaksi")

    if st.session_state.transaksi:
        df = pd.DataFrame(st.session_state.transaksi)
        df_display = df.copy()
        df_display["Debit"] = df_display["Debit"].apply(to_rp)
        df_display["Kredit"] = df_display["Kredit"].apply(to_rp)
        st.dataframe(df_display, use_container_width=True)
    else:
        st.info("Belum ada transaksi.")

# ============================
# JURNAL UMUM
# ============================
elif menu == "Jurnal Umum":
    st.markdown("<div class='subtitle'>📘 Jurnal Umum</div>", unsafe_allow_html=True)

    if not st.session_state.transaksi:
        st.info("Belum ada data.")
    else:
        df = pd.DataFrame(st.session_state.transaksi)
        df2 = df.copy()
        df2["Debit"] = df2["Debit"].apply(to_rp)
        df2["Kredit"] = df2["Kredit"].apply(to_rp)
        st.dataframe(df2, use_container_width=True)

# ============================
# BUKU BESAR
# ============================
elif menu == "Buku Besar":
    st.markdown("<div class='subtitle'>📗 Buku Besar</div>", unsafe_allow_html=True)

    if not st.session_state.transaksi:
        st.info("Belum ada data.")
    else:
        df = pd.DataFrame(st.session_state.transaksi)
        buku = buku_besar(df)
        for akun, d in buku.items():
            st.write(f"### ▶ {akun}")
            d2 = d.copy()
            d2["Debit"] = d2["Debit"].apply(to_rp)
            d2["Kredit"] = d2["Kredit"].apply(to_rp)
            d2["Saldo"] = d2["Saldo"].apply(to_rp)
            st.dataframe(d2, use_container_width=True)

# ============================
# NERACA SALDO
# ============================
elif menu == "Neraca Saldo":
    st.markdown("<div class='subtitle'>📙 Neraca Saldo</div>", unsafe_allow_html=True)

    if not st.session_state.transaksi:
        st.info("Belum ada data.")
    else:
        df = pd.DataFrame(st.session_state.transaksi)
        neraca = neraca_saldo(df)
        d2 = neraca.copy()
        d2["Debit"] = d2["Debit"].apply(to_rp)
        d2["Kredit"] = d2["Kredit"].apply(to_rp)
        d2["Saldo"] = d2["Saldo"].apply(to_rp)
        st.dataframe(d2, use_container_width=True)

# ============================
# EXPORT EXCEL
# ============================
elif menu == "Export Excel":
    st.markdown("<div class='subtitle'>📤 Export Excel</div>", unsafe_allow_html=True)

    if not st.session_state.transaksi:
        st.info("Belum ada transaksi untuk diekspor.")
    else:
        df = pd.DataFrame(st.session_state.transaksi)
        excel_file = export_excel_multi(df)

        st.download_button(
            "📥 Download Laporan Excel",
            excel_file,
            "laporan_akuntansi.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
