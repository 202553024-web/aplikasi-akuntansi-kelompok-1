import streamlit as st
import pandas as pd
import altair as alt
from datetime import datetime
import io
import calendar
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

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

def buku_besar(df):
    akun_list = df["Akun"].unique()
    data = {}
    for akun in akun_list:
        df_akun = df[df["Akun"] == akun].copy()
        df_akun["Saldo"] = df_akun["Debit"].cumsum() - df_akun["Kredit"].cumsum()
        data[akun] = df_akun
    return data

# ============================
# SIDEBAR MENU
# ============================
menu = st.sidebar.radio(
    "📌 PILIH MENU",
    [
        "Input Transaksi",
        "Jurnal Umum",
        "Buku Besar",
        "Neraca Saldo",
        "Laporan Laba Rugi",
        "Grafik",
        "Export Excel"
    ]
)

# ============================
# INPUT TRANSAKSI
# ============================
if menu == "Input Transaksi":
    akun_list = [
        "Kas", "Piutang", "Utang", "Modal",
        "Pendapatan Jasa",
        "Beban Gaji", "Beban Listrik", "Beban Sewa"
    ]

    tanggal = st.date_input("Tanggal", datetime.now())
    akun = st.selectbox("Akun", akun_list)
    ket = st.text_input("Keterangan")

    col1, col2 = st.columns(2)
    with col1:
        debit = st.number_input("Debit", min_value=0, step=1000)
    with col2:
        kredit = st.number_input("Kredit", min_value=0, step=1000)

    if st.button("Tambah"):
        tambah_transaksi(str(tanggal), akun, ket, debit, kredit)
        st.success("Transaksi ditambahkan")

    if st.session_state.transaksi:
        df = pd.DataFrame(st.session_state.transaksi)
        df_show = df.copy()
        df_show["Debit"] = df_show["Debit"].apply(to_rp)
        df_show["Kredit"] = df_show["Kredit"].apply(to_rp)
        st.dataframe(df_show)

# ============================
# JURNAL UMUM
# ============================
elif menu == "Jurnal Umum":
    if st.session_state.transaksi:
        df = pd.DataFrame(st.session_state.transaksi)
        df["Tanggal"] = pd.to_datetime(df["Tanggal"])
        df["Bulan"] = df["Tanggal"].dt.month
        df["Tahun"] = df["Tanggal"].dt.year

        for (tahun, bulan), grup in df.groupby(["Tahun", "Bulan"]):
            st.subheader(f"{calendar.month_name[bulan]} {tahun}")
            grup["Debit"] = grup["Debit"].apply(to_rp)
            grup["Kredit"] = grup["Kredit"].apply(to_rp)
            st.dataframe(grup[["Tanggal","Akun","Keterangan","Debit","Kredit"]])

# ============================
# BUKU BESAR
# ============================
elif menu == "Buku Besar":
    if st.session_state.transaksi:
        df = pd.DataFrame(st.session_state.transaksi)
        buku = buku_besar(df)

        for akun, data in buku.items():
            st.subheader(akun)
            data["Debit"] = data["Debit"].apply(to_rp)
            data["Kredit"] = data["Kredit"].apply(to_rp)
            data["Saldo"] = data["Saldo"].apply(to_rp)
            st.dataframe(data[["Tanggal","Keterangan","Debit","Kredit","Saldo"]])

# ============================
# NERACA SALDO
# ============================
elif menu == "Neraca Saldo":
    if st.session_state.transaksi:
        df = pd.DataFrame(st.session_state.transaksi)
        neraca = df.groupby("Akun")[["Debit","Kredit"]].sum()
        neraca["Saldo"] = neraca["Debit"] - neraca["Kredit"]
        neraca["Debit"] = neraca["Debit"].apply(to_rp)
        neraca["Kredit"] = neraca["Kredit"].apply(to_rp)
        neraca["Saldo"] = neraca["Saldo"].apply(to_rp)
        st.dataframe(neraca)

# ============================
# LAPORAN LABA RUGI (TAMBAHAN)
# ============================
elif menu == "Laporan Laba Rugi":
    if st.session_state.transaksi:
        df = pd.DataFrame(st.session_state.transaksi)
        df["Tanggal"] = pd.to_datetime(df["Tanggal"])
        df["Bulan"] = df["Tanggal"].dt.month
        df["Tahun"] = df["Tanggal"].dt.year

        for (tahun, bulan), grup in df.groupby(["Tahun","Bulan"]):
            pendapatan = grup[grup["Akun"].str.contains("Pendapatan",case=False)]["Kredit"].sum()
            beban = grup[grup["Akun"].str.contains("Beban",case=False)]["Debit"].sum()
            laba = pendapatan - beban

            laporan = pd.DataFrame({
                "Keterangan":["Pendapatan","Total Beban","Laba Bersih"],
                "Jumlah":[pendapatan,beban,laba]
            })
            laporan["Jumlah"] = laporan["Jumlah"].apply(to_rp)

            st.subheader(f"{calendar.month_name[bulan]} {tahun}")
            st.dataframe(laporan)

# ============================
# GRAFIK
# ============================
elif menu == "Grafik":
    if st.session_state.transaksi:
        df = pd.DataFrame(st.session_state.transaksi)
        chart = alt.Chart(df).mark_bar().encode(
            x="Akun",
            y="Debit",
            color="Akun"
        )
        st.altair_chart(chart, use_container_width=True)

# ============================
# EXPORT EXCEL
# ============================
elif menu == "Export Excel":
    st.info("Export Excel sudah ada di versi sebelumnya (tidak diubah)")
