import pandas as pd
import io
from datetime import datetime
from openpyxl import Workbook

# ============================
#  DATA CONTOH
# ============================
data = [
    {"Tanggal": "2025-01-05", "Akun": "Kas", "Keterangan": "Penjualan tunai", "Debit": 2000000, "Kredit": 0},
    {"Tanggal": "2025-01-05", "Akun": "Pendapatan Jasa", "Keterangan": "Penjualan tunai", "Debit": 0, "Kredit": 2000000},
    {"Tanggal": "2025-02-10", "Akun": "Beban Listrik", "Keterangan": "Pembayaran listrik", "Debit": 500000, "Kredit": 0},
    {"Tanggal": "2025-02-10", "Akun": "Kas", "Keterangan": "Pembayaran listrik", "Debit": 0, "Kredit": 500000},
]

df = pd.DataFrame(data)
df["Tanggal"] = pd.to_datetime(df["Tanggal"])

# ============================
#  FUNGSI PEMBANTU
# ============================
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
#  FUNGSI EXPORT EXCEL (DENGAN PEMBATAS PER BULAN)
# ============================
def export_excel_multi(df, file_name="laporan_akuntansi_lengkap.xlsx"):
    df["Tanggal"] = pd.to_datetime(df["Tanggal"])
    kelompok = df.groupby([df["Tanggal"].dt.year, df["Tanggal"].dt.month])

    output = io.BytesIO()
    writer = pd.ExcelWriter(file_name, engine="openpyxl")

    # --- JURNAL UMUM ---
    ws_jurnal = writer.book.create_sheet("Jurnal Umum")
    row = 1
    for (tahun, bulan), group in kelompok:
        nama_bulan = group["Tanggal"].dt.month_name().iloc[0].upper()
        ws_jurnal.cell(row=row, column=1, value=f"=== {nama_bulan} {tahun} ===")
        row += 2
        group_sorted = group.sort_values("Tanggal")
        group_sorted.to_excel(writer, sheet_name="Jurnal Umum", startrow=row, index=False)
        row += len(group_sorted) + 3

    # --- BUKU BESAR ---
    ws_bb = writer.book.create_sheet("Buku Besar")
    row = 1
    for (tahun, bulan), group in kelompok:
        nama_bulan = group["Tanggal"].dt.month_name().iloc[0].upper()
        ws_bb.cell(row=row, column=1, value=f"=== {nama_bulan} {tahun} ===")
        row += 2

        buku = buku_besar(group)
        for akun, data in buku.items():
            ws_bb.cell(row=row, column=1, value=f">> Akun: {akun}")
            row += 1
            data.to_excel(writer, sheet_name="Buku Besar", startrow=row, index=False)
            row += len(data) + 2
        row += 1

    # --- NERACA SALDO ---
    ws_ns = writer.book.create_sheet("Neraca Saldo")
    row = 1
    for (tahun, bulan), group in kelompok:
        nama_bulan = group["Tanggal"].dt.month_name().iloc[0].upper()
        ws_ns.cell(row=row, column=1, value=f"=== {nama_bulan} {tahun} ===")
        row += 2

        neraca = neraca_saldo(group)
        neraca.to_excel(writer, sheet_name="Neraca Saldo", startrow=row)
        row += len(neraca) + 3

    # Hapus sheet default (jika ada)
    if "Sheet" in writer.book.sheetnames:
        writer.book.remove(writer.book["Sheet"])

    writer.close()
    print(f"✅ File Excel berhasil dibuat: {file_name}")

# ============================
#  JALANKAN EXPORT
# ============================
export_excel_multi(df)
