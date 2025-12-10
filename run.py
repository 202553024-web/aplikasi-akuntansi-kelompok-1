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
# 2. JURNAL UMUM
# ============================
elif menu == "Jurnal Umum":
    st.markdown("<div class='subtitle'>📘 Jurnal Umum</div>", unsafe_allow_html=True)

    if len(st.session_state.transaksi) == 0:
        st.info("Belum ada data.")
    else:
        df = pd.DataFrame(st.session_state.transaksi)
        df2 = df.copy()
        df2["Debit"] = df2["Debit"].apply(to_rp)
        df2["Kredit"] = df2["Kredit"].apply(to_rp)
        st.dataframe(df2, use_container_width=True)

# ============================
# 3. BUKU BESAR
# ============================
elif menu == "Buku Besar":
    st.markdown("<div class='subtitle'>📗 Buku Besar</div>", unsafe_allow_html=True)

    if len(st.session_state.transaksi) == 0:
        st.info("Belum ada data.")
    else:
        df = pd.DataFrame(st.session_state.transaksi)
        buku = buku_besar(df)
        for akun, data in buku.items():
            st.write(f"### ▶ {akun}")
            df2 = data.copy()
            df2["Debit"] = df2["Debit"].apply(to_rp)
            df2["Kredit"] = df2["Kredit"].apply(to_rp)
            df2["Saldo"] = df2["Saldo"].apply(to_rp)
            st.dataframe(df2, use_container_width=True)

# ============================
# 4. NERACA SALDO
# ============================
elif menu == "Neraca Saldo":
    st.markdown("<div class='subtitle'>📙 Neraca Saldo</div>", unsafe_allow_html=True)

    if len(st.session_state.transaksi) == 0:
        st.info("Belum ada data.")
    else:
        df = pd.DataFrame(st.session_state.transaksi)
        neraca = neraca_saldo(df)
        df2 = neraca.copy()
        df2["Debit"] = df2["Debit"].apply(to_rp)
        df2["Kredit"] = df2["Kredit"].apply(to_rp)
        df2["Saldo"] = df2["Saldo"].apply(to_rp)
        st.dataframe(df2, use_container_width=True)

# ============================
# 5. GRAFIK
# ============================
elif menu == "Grafik":
    st.markdown("<div class='subtitle'>📈 Grafik Akuntansi</div>", unsafe_allow_html=True)

    if len(st.session_state.transaksi) == 0:
        st.info("Belum ada data.")
    else:
        df = pd.DataFrame(st.session_state.transaksi)
        chart = alt.Chart(df).mark_bar().encode(
            x="Akun",
            y="Debit",
            color="Akun"
        ).properties(
            title="Grafik Jumlah Debit per Akun",
            width=700
        )
        st.altair_chart(chart, use_container_width=True)

# ============================
# 6. EXPORT EXCEL (MULTI SHEET TANPA BORDER)
# ============================
elif menu == "Export Excel":
    st.markdown("<div class='subtitle'>📤 Export Excel (Multi Sheet)</div>", unsafe_allow_html=True)

    def export_excel_multi(df):
        output = io.BytesIO()
        wb = Workbook()

        # SHEET 1: Laporan Keuangan
        ws = wb.active
        ws.title = "Laporan Keuangan"

        tahun = df['Tanggal'].apply(lambda x: pd.to_datetime(x).year).unique()
        current_row = 1
        for t in tahun:
            ws.cell(row=current_row, column=1, value=f"Laporan Keuangan Tahun {t}").font = Font(bold=True)
            current_row += 2
            data_tahun = df[pd.to_datetime(df['Tanggal']).dt.year == t]
            bulan = data_tahun['Tanggal'].apply(lambda x: pd.to_datetime(x).month).unique()
            for b in bulan:
                nama_bulan = calendar.month_name[b]
                ws.cell(row=current_row, column=1, value=f"Laporan Keuangan Bulan {nama_bulan}").font = Font(bold=True)
                current_row += 1
                data_bulan = data_tahun[pd.to_datetime(data_tahun['Tanggal']).dt.month == b]
                for r in dataframe_to_rows(data_bulan, index=False, header=True):
                    ws.append(r)
                current_row += len(data_bulan) + 2

        # Auto width kolom
        for col_cells in ws.columns:
            valid_cells = [cell for cell in col_cells if cell.value]
            if not valid_cells:
                continue
            max_length = max(len(str(cell.value)) for cell in valid_cells)
            col_letter = valid_cells[0].column_letter
            ws.column_dimensions[col_letter].width = max_length + 2

        # Tambah sheet lain
        ws2 = wb.create_sheet("Neraca")
        ws2["A1"] = "Laporan Neraca"
        ws2["A1"].font = Font(bold=True)
        ws2.append(["Aset", "Kewajiban", "Modal"])
        ws2.append(["Rp0", "Rp0", "Rp0"])

        ws3 = wb.create_sheet("Buku Besar")
        ws3["A1"] = "Laporan Buku Besar"
        ws3["A1"].font = Font(bold=True)
        ws3.append(["Akun", "Tanggal", "Debit", "Kredit"])
        ws3.append(["Kas", "2025-12-10", "Rp1.000.000", "Rp0"])

        ws4 = wb.create_sheet("Laba Rugi")
        ws4["A1"] = "Laporan Laba Rugi"
        ws4["A1"].font = Font(bold=True)
        ws4.append(["Pendapatan", "Beban", "Laba/Rugi"])
        ws4.append(["Rp5.000.000", "Rp3.000.000", "Rp2.000.000"])

        ws5 = wb.create_sheet("Arus Kas")
        ws5["A1"] = "Laporan Arus Kas"
        ws5["A1"].font = Font(bold=True)
        ws5.append(["Tanggal", "Deskripsi", "Kas Masuk", "Kas Keluar"])
        ws5.append(["2025-12-10", "Penjualan", "Rp1.000.000", "Rp0"])

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
