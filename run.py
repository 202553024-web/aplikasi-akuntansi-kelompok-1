# Streamlit Accounting App (UI preserved, backend improved)
import streamlit as st
import pandas as pd
import altair as alt
from datetime import datetime
import io

# ============================
# CONFIG UI (Tidak diubah)
# ============================
st.set_page_config(page_title="Aplikasi Akuntansi", page_icon="💰", layout="wide")
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
# FORMAT RUPIAH (Backend Only)
# ============================
def to_rp(n):
    try:
        return "Rp {:,}".format(int(n)).replace(",", ".")
    except:
        return "Rp 0"

# ============================
# BACKEND FUNCTIONS
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
    if 0 <= idx < len(st.session_state.transaksi):
        st.session_state.transaksi.pop(idx)

# Buku Besar otomatis

def buku_besar(df):
    akun_list = df["Akun"].unique()
    buku_besar_data = {}
    for akun in akun_list:
        df_akun = df[df["Akun"] == akun].copy()
        df_akun["Saldo"] = df_akun["Debit"].cumsum() - df_akun["Kredit"].cumsum()
        buku_besar_data[akun] = df_akun
    return buku_besar_data

# Neraca Saldo otomatis
def neraca_saldo(df):
    grouped = df.groupby("Akun")["Debit","Kredit"].sum()
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
# 1. INPUT TRANSAKSI (UI tetap)
# ============================
if menu == "Input Transaksi":
    st.markdown("<div class='subtitle'>📝 Input Transaksi</div>", unsafe_allow_html=True)

    akun_list = [
        "Kas", "Piutang", "Utang", "Modal",
        "Pendapatan Jasa", "Beban Gaji", "Beban Listrik", "Beban Sewa"
    ]

    col1, col2 = st.columns(2)
    with col1:
        tanggal = st.date_input("Tanggal", datetime.now(), key="tgl_input")
        akun = st.selectbox("Akun", akun_list, key="akun_input")
    with col2:
        ket = st.text_input("Keterangan", key="ket_input")

    col3, col4 = st.columns(2)
    with col3:
        debit = st.number_input("Debit (Rp)", min_value=0, step=1000, key="debit_input")
    with col4:
        kredit = st.number_input("Kredit (Rp)", min_value=0, step=1000, key="kredit_input")

    if st.button("Tambah Transaksi"):
        tambah_transaksi(str(tanggal), akun, ket, debit, kredit)
        st.success("Transaksi berhasil ditambahkan!")

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
# 2. JURNAL UMUM (UI tetap)
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
# 3. BUKU BESAR (UI tetap)
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
# 5. GRAFIK (UI tetap)
# ============================
elif menu == "Grafik":
    st.markdown("<div class='subtitle'>📈 Grafik Akuntansi</div>", unsafe_allow_html=True)

    if len(st.session_state.transaksi) == 0:
        st.info("Belum ada data.")
    else:
        df = pd.DataFrame(st.session_state.transaksi)
        chart = alt.Chart(df).mark_bar().encode(
            x="Akun", y="Debit", color="Akun"
        ).properties(title="Grafik Jumlah Debit per Akun", width=700)
        st.altair_chart(chart, use_container_width=True)

# ============================
# EXPORT EXCEL (Backend only improved)
# ============================
def export_excel_multi(df):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine="openpyxl")

    # --- SHEET 1 : Jurnal Umum ---
    df.to_excel(writer, index=False, sheet_name="Jurnal Umum")

    # --- SHEET 2 : Buku Besar ---
    buku = buku_besar(df)
    start = 0
    for akun, data in buku.items():
        data.to_excel(writer, sheet_name="Buku Besar", startrow=start, index=False)
        start += len(data) + 3

    # --- SHEET 3 : Neraca Saldo ---
    ner = neraca_saldo(df)
    ner.to_excel(writer, sheet_name="Neraca Saldo")

    writer.close()
    return output.getvalue()

# ============================
# TOMBOL EXPORT (UI tetap)
# ============================
elif menu == "Export Excel":
    st.markdown("<div class='subtitle'>📤 Export Excel</div>", unsafe_allow_html=True)

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
