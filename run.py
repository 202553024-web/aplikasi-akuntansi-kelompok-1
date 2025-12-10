import streamlit as st
import pandas as pd
import altair as alt
from datetime import datetime
import io
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, Border, Side

# ============================
# CONFIG APP
# ============================
st.set_page_config(page_title="Aplikasi Akuntansi", page_icon="💰", layout="wide")

st.markdown("""
<style>
    .title { 
        font-size: 42px; 
        font-weight: 900; 
        color: #0d47a1; 
        text-align:center;
        margin-bottom: 5px;
    }
    .subtitle { 
        font-size: 24px; 
        font-weight: 650; 
        color:#0d47a1; 
        margin-top: 20px;
        margin-bottom: 10px;
    }
    .card {
        background: #f5f7fa;
        padding:18px;
        border-radius:12px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        margin-bottom: 20px;
    }
    .stButton>button {
        background-color: #0d47a1 !important;
        color: white !important;
        padding: 10px 20px;
        border-radius: 10px;
        font-size: 17px;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("<div class='title'>💰 Aplikasi Akuntansi Modern</div>", unsafe_allow_html=True)
st.write("")

# ============================
# SESSION DATA
# ============================
if "transaksi" not in st.session_state:
    st.session_state.transaksi = []

# ============================
# FORMAT RUPIAH
# ============================
def to_rp(n):
    return "Rp {:,}".format(int(n)).replace(",", ".") if n else "Rp 0"

# ============================
# FUNGSI AKUNTANSI
# ============================
def tambah_transaksi(tgl, akun, ket, debit, kredit):
    st.session_state.transaksi.append({
        "Tanggal": tgl,
        "Akun": akun,
        "Keterangan": ket,
        "Debit": int(debit),
        "Kredit": int(kredit),
        "Bulan": str(pd.to_datetime(tgl).month)
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
    grouped = df.groupby("Akun")[["Debit", "Kredit"]].sum()
    grouped["Saldo"] = grouped["Debit"] - grouped["Kredit"]
    return grouped

# ============================
# EXPORT EXCEL — PER BULAN
# ============================
def format_excel(ws):
    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for row in ws.iter_rows():
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(horizontal="center")
            if cell.row == 1:
                cell.font = Font(bold=True)

    # Auto width
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                max_len = max(max_len, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[col_letter].width = max_len + 2


def export_excel_multi(df):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine="openpyxl")

    df["Tanggal"] = pd.to_datetime(df["Tanggal"])
    df["Bulan"] = df["Tanggal"].dt.month

    # Loop per bulan
    for bulan, df_bln in df.groupby("Bulan"):

        # 1. Jurnal Umum
        df_bln.to_excel(writer, sheet_name=f"Jurnal {bulan}", index=False)

        # 2. Buku Besar
        buku = buku_besar(df_bln)
        start_row = 0
        ws_name = f"Buku Besar {bulan}"

        for akun, data in buku.items():
            data.to_excel(writer, sheet_name=ws_name, startrow=start_row, index=False)
            start_row += len(data) + 3

        # 3. Neraca Saldo
        ner = neraca_saldo(df_bln)
        ner.to_excel(writer, sheet_name=f"Neraca {bulan}")

    writer.close()

    # Format semua sheet
    from openpyxl import load_workbook
    book = load_workbook(output)
    for sheet in book.sheetnames:
        ws = book[sheet]
        format_excel(ws)

    # Simpan kembali
    new_output = io.BytesIO()
    book.save(new_output)
    return new_output.getvalue()

# ============================
# SIDEBAR MENU
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
    st.markdown("<div class='card'>", unsafe_allow_html=True)

    akun_list = ["Kas","Piutang","Utang","Modal","Pendapatan Jasa","Beban Gaji","Beban Listrik","Beban Sewa"]

    tanggal = st.date_input("Tanggal", datetime.now())
    akun = st.selectbox("Akun", akun_list)
    ket = st.text_input("Keterangan")

    col1, col2 = st.columns(2)
    debit = col1.number_input("Debit (Rp)", min_value=0, step=1000, format="%d")
    kredit = col2.number_input("Kredit (Rp)", min_value=0, step=1000, format="%d")

    if st.button("Tambah Transaksi"):
        tambah_transaksi(str(tanggal), akun, ket, debit, kredit)
        st.success("Transaksi berhasil ditambahkan!")

    st.markdown("</div>", unsafe_allow_html=True)

    st.write("### 📄 Daftar Transaksi")
    if st.session_state.transaksi:
        df = pd.DataFrame(st.session_state.transaksi)
        df2 = df.copy()
        df2["Debit"] = df2["Debit"].apply(to_rp)
        df2["Kredit"] = df2["Kredit"].apply(to_rp)
        st.dataframe(df2, use_container_width=True)

# ============================
# JURNAL UMUM
# ============================
elif menu == "Jurnal Umum":
    st.markdown("<div class='subtitle'>📘 Jurnal Umum</div>", unsafe_allow_html=True)
    if st.session_state.transaksi:
        df = pd.DataFrame(st.session_state.transaksi)
        df2 = df.copy()
        df2["Debit"] = df2["Debit"].apply(to_rp)
        df2["Kredit"] = df2["Kredit"].apply(to_rp)
        st.dataframe(df2, use_container_width=True)
    else:
        st.info("Belum ada data.")

# ============================
# BUKU BESAR
# ============================
elif menu == "Buku Besar":
    st.markdown("<div class='subtitle'>📗 Buku Besar</div>", unsafe_allow_html=True)

    if st.session_state.transaksi:
        df = pd.DataFrame(st.session_state.transaksi)
        buku = buku_besar(df)

        for akun, data in buku.items():
            st.write(f"### ▶ {akun}")
            df2 = data.copy()
            df2["Debit"] = df2["Debit"].apply(to_rp)
            df2["Kredit"] = df2["Kredit"].apply(to_rp)
            df2["Saldo"] = df2["Saldo"].apply(to_rp)
            st.dataframe(df2, use_container_width=True)
    else:
        st.info("Belum ada data.")

# ============================
# NERACA SALDO
# ============================
elif menu == "Neraca Saldo":
    st.markdown("<div class='subtitle'>📙 Neraca Saldo</div>", unsafe_allow_html=True)
    if st.session_state.transaksi:
        df = pd.DataFrame(st.session_state.transaksi)
        ner = neraca_saldo(df)
        df2 = ner.copy()
        df2["Debit"] = df2["Debit"].apply(to_rp)
        df2["Kredit"] = df2["Kredit"].apply(to_rp)
        df2["Saldo"] = df2["Saldo"].apply(to_rp)
        st.dataframe(df2, use_container_width=True)
    else:
        st.info("Belum ada data.")

# ============================
# GRAFIK
# ============================
elif menu == "Grafik":
    st.markdown("<div class='subtitle'>📈 Grafik Akuntansi</div>", unsafe_allow_html=True)

    if st.session_state.transaksi:
        df = pd.DataFrame(st.session_state.transaksi)

        chart = alt.Chart(df).mark_bar().encode(
            x="Akun",
            y="Debit",
            color="Akun"
        ).properties(title="Grafik Jumlah Debit per Akun", width=700)

        st.altair_chart(chart, use_container_width=True)
    else:
        st.info("Belum ada data.")

# ============================
# EXPORT EXCEL
# ============================
elif menu == "Export Excel":
    st.markdown("<div class='subtitle'>📤 Export Excel</div>", unsafe_allow_html=True)

    if not st.session_state.transaksi:
        st.info("Belum ada transaksi.")
    else:
        df = pd.DataFrame(st.session_state.transaksi)
        excel_file = export_excel_multi(df)

        st.download_button(
            label="📥 Export ke Excel (Per Bulan + Rapi)",
            data=excel_file,
            file_name="laporan_akuntansi_per_bulan.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
