import streamlit as st
import pandas as pd
from datetime import datetime
import io

from openpyxl.styles import Font, PatternFill, Alignment

# ================= CONFIG =================
st.set_page_config("Aplikasi Akuntansi", "💰", layout="wide")

if "transaksi" not in st.session_state:
    st.session_state.transaksi = []

# ================= FUNGSI =================
def tambah_transaksi(tgl, akun, ket, debit, kredit):
    st.session_state.transaksi.append({
        "Tanggal": pd.to_datetime(tgl),
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

def laba_rugi(df):
    pendapatan = df[df["Akun"].str.contains("Pendapatan", case=False)]
    beban = df[df["Akun"].str.contains("Beban", case=False)]

    total_pendapatan = pendapatan["Kredit"].sum()
    total_beban = beban["Debit"].sum()

    laba = total_pendapatan - total_beban

    return total_pendapatan, total_beban, laba

# ================= EXPORT EXCEL =================
def export_excel(df):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine="openpyxl")

    header = PatternFill("solid", fgColor="1F4E78")
    hfont = Font(bold=True, color="FFFFFF")
    tfont = Font(bold=True, size=14)
    center = Alignment(horizontal="center")

    # ================= LAPORAN KEUANGAN =================
    ws = writer.book.create_sheet("Laporan Keuangan")
    row = 1

    for tahun, df_tahun in df.groupby(df["Tanggal"].dt.year):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
        ws.cell(row=row, column=1, value=f"Laporan Keuangan Tahun {tahun}").font = tfont
        ws.cell(row=row, column=1).alignment = center
        row += 1

        for bulan, df_bulan in df_tahun.groupby(df_tahun["Tanggal"].dt.month):
            nama_bulan = df_bulan["Tanggal"].dt.strftime("%B").iloc[0]
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
            ws.cell(row=row, column=1, value=f"Bulan {nama_bulan}").font = Font(bold=True)
            ws.cell(row=row, column=1).alignment = center
            row += 1

            headers = ["Tanggal", "Akun", "Keterangan", "Debit", "Kredit"]
            for i, h in enumerate(headers, 1):
                c = ws.cell(row=row, column=i, value=h)
                c.fill = header
                c.font = hfont
                c.alignment = center
            row += 1

            for _, r in df_bulan.iterrows():
                ws.append([
                    r["Tanggal"],
                    r["Akun"],
                    r["Keterangan"],
                    r["Debit"],
                    r["Kredit"]
                ])
                row += 1
            row += 1

    # ================= LABA RUGI =================
    ws = writer.book.create_sheet("Laba Rugi")
    ws.merge_cells("A1:B1")
    ws["A1"] = "Laporan Laba Rugi"
    ws["A1"].font = tfont
    ws["A1"].alignment = center

    total_pendapatan, total_beban, laba = laba_rugi(df)

    data_lr = [
        ("Pendapatan", total_pendapatan),
        ("Beban", total_beban),
        ("Laba / Rugi Bersih", laba)
    ]

    row = 3
    for nama, nilai in data_lr:
        ws.cell(row=row, column=1, value=nama)
        ws.cell(row=row, column=2, value=nilai)
        ws.cell(row=row, column=2).number_format = '"Rp" #,##0.00'
        row += 1

    # ================= FORMAT UANG =================
    for sheet in writer.book.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '"Rp" #,##0.00'

    writer.close()
    return output.getvalue()

# ================= UI =================
menu = st.sidebar.radio("Menu", ["Input", "Export Excel"])

if menu == "Input":
    tgl = st.date_input("Tanggal", datetime.now())
    akun = st.text_input("Akun")
    ket = st.text_input("Keterangan")
    debit = st.number_input("Debit", 0)
    kredit = st.number_input("Kredit", 0)

    if st.button("Tambah"):
        tambah_transaksi(tgl, akun, ket, debit, kredit)
        st.success("Transaksi ditambahkan")

    st.dataframe(pd.DataFrame(st.session_state.transaksi))

else:
    if st.session_state.transaksi:
        df = pd.DataFrame(st.session_state.transaksi)
        file = export_excel(df)
        st.download_button(
            "📥 Download Excel",
            file,
            "laporan_keuangan_lengkap.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("Belum ada data")
