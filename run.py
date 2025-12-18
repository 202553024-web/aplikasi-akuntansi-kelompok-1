import streamlit as st
import pandas as pd
import altair as alt
from datetime import datetime
import io

from openpyxl.styles import Font, PatternFill, Alignment

# ============================
# CONFIG HALAMAN
# ============================
st.set_page_config(
    page_title="Aplikasi Akuntansi",
    page_icon="💰",
    layout="wide"
)

st.markdown("""
<style>
.title { font-size:38px; font-weight:800; text-align:center; }
.subtitle { font-size:22px; font-weight:600; margin-top:20px; }
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
def to_rp(x):
    return f"Rp {int(x):,}".replace(",", ".")

# ============================
# LOGIKA AKUNTANSI
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
    hasil = {}
    for akun in df["Akun"].unique():
        d = df[df["Akun"] == akun].copy()
        d["Saldo"] = d["Debit"].cumsum() - d["Kredit"].cumsum()
        hasil[akun] = d
    return hasil

def neraca_saldo(df):
    g = df.groupby("Akun")[["Debit", "Kredit"]].sum()
    g["Saldo"] = g["Debit"] - g["Kredit"]
    return g.reset_index()

# ============================
# SIDEBAR
# ============================
menu = st.sidebar.radio(
    "📌 Menu",
    ["Input Transaksi", "Jurnal Umum", "Buku Besar", "Neraca Saldo", "Export Excel"]
)

# ============================
# INPUT TRANSAKSI
# ============================
if menu == "Input Transaksi":
    st.markdown("<div class='subtitle'>📝 Input Transaksi</div>", unsafe_allow_html=True)

    akun_list = [
        "Kas", "Piutang", "Utang", "Modal",
        "Pendapatan Jasa", "Beban Gaji",
        "Beban Listrik", "Beban Sewa"
    ]

    tgl = st.date_input("Tanggal", datetime.now())
    akun = st.selectbox("Akun", akun_list)
    ket = st.text_input("Keterangan")

    col1, col2 = st.columns(2)
    debit = col1.number_input("Debit", 0, step=1000)
    kredit = col2.number_input("Kredit", 0, step=1000)

    if st.button("Tambah"):
        tambah_transaksi(str(tgl), akun, ket, debit, kredit)
        st.success("Transaksi ditambahkan")

# ============================
# JURNAL UMUM
# ============================
elif menu == "Jurnal Umum":
    st.markdown("<div class='subtitle'>📘 Jurnal Umum</div>", unsafe_allow_html=True)
    if st.session_state.transaksi:
        df = pd.DataFrame(st.session_state.transaksi)
        df["Debit"] = df["Debit"].apply(to_rp)
        df["Kredit"] = df["Kredit"].apply(to_rp)
        st.dataframe(df, use_container_width=True)
    else:
        st.info("Belum ada data")

# ============================
# BUKU BESAR
# ============================
elif menu == "Buku Besar":
    st.markdown("<div class='subtitle'>📗 Buku Besar</div>", unsafe_allow_html=True)
    if st.session_state.transaksi:
        df = pd.DataFrame(st.session_state.transaksi)
        buku = buku_besar(df)
        for akun, d in buku.items():
            st.write(f"### {akun}")
            d2 = d.copy()
            d2["Debit"] = d2["Debit"].apply(to_rp)
            d2["Kredit"] = d2["Kredit"].apply(to_rp)
            d2["Saldo"] = d2["Saldo"].apply(to_rp)
            st.dataframe(d2, use_container_width=True)
    else:
        st.info("Belum ada data")

# ============================
# NERACA SALDO
# ============================
elif menu == "Neraca Saldo":
    st.markdown("<div class='subtitle'>📙 Neraca Saldo</div>", unsafe_allow_html=True)
    if st.session_state.transaksi:
        df = pd.DataFrame(st.session_state.transaksi)
        ns = neraca_saldo(df)
        ns["Debit"] = ns["Debit"].apply(to_rp)
        ns["Kredit"] = ns["Kredit"].apply(to_rp)
        ns["Saldo"] = ns["Saldo"].apply(to_rp)
        st.dataframe(ns, use_container_width=True)
    else:
        st.info("Belum ada data")

# ============================
# EXPORT EXCEL (TAMPILAN MIRIP CONTOH)
# ============================
def export_excel(df):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine="openpyxl")

    header = PatternFill("solid", fgColor="1F4E78")
    hfont = Font(bold=True, color="FFFFFF")
    tfont = Font(bold=True, size=14)
    center = Alignment(horizontal="center")

    # ===== JURNAL UMUM =====
    df.to_excel(writer, sheet_name="Jurnal Umum", startrow=3, index=False)
    ws = writer.sheets["Jurnal Umum"]
    ws.merge_cells("A1:E1")
    ws["A1"] = "JURNAL UMUM"
    ws["A1"].font = tfont
    ws["A1"].alignment = center

    for c in range(1, 6):
        cell = ws.cell(row=4, column=c)
        cell.fill = header
        cell.font = hfont
        cell.alignment = center

    # ===== BUKU BESAR =====
    ws = writer.book.create_sheet("Buku Besar")
    row = 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
    ws.cell(row=row, column=1, value="BUKU BESAR").font = tfont
    ws.cell(row=row, column=1).alignment = center
    row += 2

    buku = buku_besar(df)
    for akun, d in buku.items():
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
        ws.cell(row=row, column=1, value=f"Nama Akun : {akun}").font = Font(bold=True)
        row += 1

        headers = ["Tanggal", "Keterangan", "Debit", "Kredit", "Saldo"]
        for i, h in enumerate(headers, 1):
            cell = ws.cell(row=row, column=i, value=h)
            cell.fill = header
            cell.font = hfont
            cell.alignment = center
        row += 1

        for _, r in d.iterrows():
            ws.append([r["Tanggal"], r["Keterangan"], r["Debit"], r["Kredit"], r["Saldo"]])
            row += 1
        row += 2

    # ===== NERACA SALDO =====
    ns = neraca_saldo(df)
    ns.to_excel(writer, sheet_name="Neraca Saldo", startrow=3, index=False)
    ws = writer.sheets["Neraca Saldo"]
    ws.merge_cells("A1:D1")
    ws["A1"] = "NERACA SALDO"
    ws["A1"].font = tfont
    ws["A1"].alignment = center

    for c in range(1, 5):
        cell = ws.cell(row=4, column=c)
        cell.fill = header
        cell.font = hfont
        cell.alignment = center

    writer.close()
    return output.getvalue()

# ============================
# TOMBOL EXPORT
# ============================
if menu == "Export Excel":
    st.markdown("<div class='subtitle'>📤 Export Excel</div>", unsafe_allow_html=True)
    if st.session_state.transaksi:
        df = pd.DataFrame(st.session_state.transaksi)
        excel = export_excel(df)
        st.download_button(
            "📥 Download Excel",
            excel,
            "laporan_akuntansi.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("Belum ada data")
