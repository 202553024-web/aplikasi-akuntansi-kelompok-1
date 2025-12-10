# === FINAL VERSION — TANPA DUPLIKASI WIDGET ===
# === JURNAL UMUM 1 SHEET (FORMAT B), BUKU BESAR, NERACA SALDO ===

import streamlit as st
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill, NamedStyle
from io import BytesIO

st.set_page_config(page_title="Aplikasi Akuntansi", layout="wide")

# =====================================
# DATAFRAME UTAMA (SESSION STATE)
# =====================================
if "jurnal" not in st.session_state:
    st.session_state.jurnal = pd.DataFrame(columns=["Tanggal", "Bulan", "Akun", "Keterangan", "Debit", "Kredit"])

# =====================================
# INPUT TRANSAKSI (TANPA DUPLIKASI WIDGET)
# =====================================
st.title("📒 Input Transaksi — Jurnal Umum (1 Sheet)")
col1, col2, col3 = st.columns(3)

with col1:
    tanggal = st.date_input("Tanggal", datetime.now(), key="tgl_input")

with col2:
    akun = st.text_input("Akun", key="akun_input")

with col3:
    keterangan = st.text_input("Keterangan", key="ket_input")

col4, col5 = st.columns(2)
with col4:
    debit = st.number_input("Debit (Rp)", min_value=0, step=1000, key="debit_input")

with col5:
    kredit = st.number_input("Kredit (Rp)", min_value=0, step=1000, key="kredit_input")

# =====================================
# TAMBAHKAN TRANSAKSI
# =====================================
if st.button("Tambah Transaksi", key="btn_tambah"):
    bulan = tanggal.strftime("%B")
    st.session_state.jurnal.loc[len(st.session_state.jurnal)] = [
        tanggal.strftime("%Y-%m-%d"), bulan, akun, keterangan, debit, kredit
    ]
    st.success("Transaksi berhasil ditambahkan!")

# =====================================
# TAMPILKAN JURNAL
# =====================================
st.subheader("📘 Data Jurnal Umum")
st.dataframe(st.session_state.jurnal)

# =====================================
# FITUR HAPUS TRANSAKSI
# =====================================
st.subheader("🗑 Hapus Transaksi")
if len(st.session_state.jurnal) > 0:
    indeks = st.number_input("Masukkan indeks transaksi yang ingin dihapus", 0, len(st.session_state.jurnal)-1, key="hapus_idx")
    if st.button("Hapus", key="btn_hapus"):
        st.session_state.jurnal.drop(indeks, inplace=True)
        st.session_state.jurnal.reset_index(drop=True, inplace=True)
        st.success("Transaksi berhasil dihapus!")
else:
    st.info("Belum ada transaksi yang dapat dihapus.")


# =====================================
# EXPORT EXCEL (3 SHEET)
# =====================================
st.subheader("📤 Export ke Excel (3 Sheet)")

def generate_excel(df):
    wb = Workbook()

    # =====================
    # 1️⃣ SHEET JURNAL UMUM
    # =====================
    ws = wb.active
    ws.title = "Jurnal Umum"

    headers = ["Tanggal", "Bulan", "Akun", "Keterangan", "Debit", "Kredit"]
    ws.append(headers)

    for row in df.itertuples(index=False):
        ws.append(list(row))

    for col in ws.columns:
        col_letter = col[0].column_letter
        ws.column_dimensions[col_letter].width = 18

    # =====================
    # 2️⃣ BUKU BESAR
    # =====================
    ws2 = wb.create_sheet("Buku Besar")
    ws2.append(["Akun", "Tanggal", "Keterangan", "Debit", "Kredit", "Saldo"])
    for col in ws2.columns:
        ws2.column_dimensions[col[0].column_letter].width = 18

    # =====================
    # 3️⃣ NERACA SALDO
    # =====================
    ws3 = wb.create_sheet("Neraca Saldo")
    ws3.append(["Akun", "Debit", "Kredit"])
    for col in ws3.columns:
        ws3.column_dimensions[col[0].column_letter].width = 18

    # Simpan ke BytesIO
    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# Tombol Export
if st.button("Export Excel", key="btn_export"):
    file_data = generate_excel(st.session_state.jurnal)
    st.download_button("Download File Excel", file_data, file_name="Laporan-Akuntansi.xlsx")

