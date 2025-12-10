import streamlit as st
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import os

# ===== Fungsi Format Uang Rupiah =====
def format_rupiah(x):
    return f"Rp {x:,.0f}".replace(",", ".")

# ===== Fungsi Tambah Data =====
def tambah_data(nama_file, data_baru):
    if not os.path.exists(nama_file):
        # buat file kosong jika belum ada
        df = pd.DataFrame(columns=["Tanggal", "Akun", "Keterangan", "Debit", "Kredit"])
        df.to_excel(nama_file, index=False)

    df = pd.read_excel(nama_file)
    df = pd.concat([df, pd.DataFrame([data_baru])], ignore_index=True)
    df.to_excel(nama_file, index=False)

# ===== Fungsi Tulis ke Excel dengan Format =====
def tulis_laporan_keuangan(nama_file):
    df = pd.read_excel(nama_file)
    df["Tanggal"] = pd.to_datetime(df["Tanggal"])
    df["Bulan"] = df["Tanggal"].dt.strftime("%B %Y")

    # urutkan berdasarkan bulan
    df = df.sort_values("Tanggal")

    wb = load_workbook(nama_file)
    ws = wb.active
    ws.title = "Laporan Keuangan"

    # bersihkan sheet lama
    for row in ws["A1:Z9999"]:
        for cell in row:
            cell.value = None

    row_num = 1

    for bulan, group in df.groupby("Bulan"):
        # Header Bulan
        ws.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=4)
        bulan_cell = ws.cell(row=row_num, column=1)
        bulan_cell.value = f"=== {bulan.upper()} ==="
        bulan_cell.font = Font(bold=True, size=12, color="0000FF")
        bulan_cell.alignment = Alignment(horizontal="center")
        row_num += 1

        # Header tabel
        headers = ["Akun", "Debit", "Kredit", "Saldo"]
        for col_num, header in enumerate(headers, start=1):
            cell = ws.cell(row=row_num, column=col_num)
            cell.value = header
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
            cell.alignment = Alignment(horizontal="center")
        row_num += 1

        # Isi data
        akun_unik = group["Akun"].unique()
        for akun in akun_unik:
            df_akun = group[group["Akun"] == akun]
            debit_total = df_akun["Debit"].sum()
            kredit_total = df_akun["Kredit"].sum()
            saldo = debit_total - kredit_total

            ws.cell(row=row_num, column=1, value=akun)
            ws.cell(row=row_num, column=2, value=debit_total)
            ws.cell(row=row_num, column=3, value=kredit_total)
            ws.cell(row=row_num, column=4, value=saldo)
            row_num += 1

        # Total per bulan
        debit_total_bulan = group["Debit"].sum()
        kredit_total_bulan = group["Kredit"].sum()
        saldo_total_bulan = debit_total_bulan - kredit_total_bulan

        ws.cell(row=row_num, column=1, value="TOTAL").font = Font(bold=True)
        ws.cell(row=row_num, column=2, value=debit_total_bulan)
        ws.cell(row=row_num, column=3, value=kredit_total_bulan)
        ws.cell(row=row_num, column=4, value=saldo_total_bulan)
        row_num += 3  # jarak antar bulan

    # Format angka Rupiah dan border
    thin_border = Border(left=Side(style="thin"), right=Side(style="thin"),
                         top=Side(style="thin"), bottom=Side(style="thin"))

    for row in ws.iter_rows(min_row=1, max_row=row_num - 1, min_col=2, max_col=4):
        for cell in row:
            if isinstance(cell.value, (int, float)):
                cell.number_format = '"Rp"#,##0'
            cell.border = thin_border

    wb.save(nama_file)

# ===== Tampilan Streamlit =====
st.title("📘 Aplikasi Laporan Keuangan Bulanan")
st.write("Input transaksi dan ekspor ke Excel dengan format rapi per bulan.")

nama_file = "Laporan_Keuangan.xlsx"

with st.form("form_input"):
    tanggal = st.date_input("Tanggal Transaksi", datetime.now())
    akun = st.text_input("Nama Akun")
    keterangan = st.text_input("Keterangan")
    debit = st.number_input("Debit (Rp)", min_value=0)
    kredit = st.number_input("Kredit (Rp)", min_value=0)
    submitted = st.form_submit_button("Tambah Transaksi")

if submitted:
    tambah_data(nama_file, {
        "Tanggal": tanggal,
        "Akun": akun,
        "Keterangan": keterangan,
        "Debit": debit,
        "Kredit": kredit
    })
    tulis_laporan_keuangan(nama_file)
    st.success("✅ Transaksi berhasil disimpan dan laporan diperbarui!")

if os.path.exists(nama_file):
    df_tampil = pd.read_excel(nama_file)
    df_tampil["Debit"] = df_tampil["Debit"].apply(lambda x: format_rupiah(x))
    df_tampil["Kredit"] = df_tampil["Kredit"].apply(lambda x: format_rupiah(x))
    st.dataframe(df_tampil)

    with open(nama_file, "rb") as f:
        st.download_button("📥 Unduh Laporan Excel", f, file_name="Laporan_Keuangan.xlsx")
