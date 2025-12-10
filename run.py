import streamlit as st
import pandas as pd
import altair as alt
from datetime import datetime
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ============================
# CONFIG TAMPAK APLIKASI
# ============================
st.set_page_config(
    page_title="Aplikasi Akuntansi",
    page_icon="💰",
    layout="wide"
)

# CSS untuk UI modern
st.markdown("""
<style>
    .title { font-size: 40px; font-weight: 900; color: #0d47a1; text-align:center; margin-bottom:6px; }
    .subtitle { font-size: 20px; font-weight: 700; color:#0d47a1; margin-top: 8px; margin-bottom:6px; }
    .card { padding: 18px; border-radius: 12px; background: linear-gradient(180deg,#ffffff 0%, #f6f9ff 100%); box-shadow: 0 2px 8px rgba(13,71,161,0.06); border:1px solid rgba(13,71,161,0.06); }
    .stButton>button { background-color: #0d47a1 !important; color: white !important; padding: 10px 18px; border-radius: 10px; font-size: 15px; font-weight:600; }
    .small-muted { font-size:12px; color:#6b7280; }
</style>
""", unsafe_allow_html=True)

st.markdown("<div class='title'>📊 Aplikasi Akuntansi</div>", unsafe_allow_html=True)
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
    try:
        return "Rp {:,}".format(int(n)).replace(",", ".")
    except:
        return "Rp 0"

# ============================
# FUNGSI AKUNTANSI
# ============================
def tambah_transaksi(tgl, akun, ket, debit, kredit):
    st.session_state.transaksi.append({
        "Tanggal": str(tgl),
        "Akun": akun,
        "Keterangan": ket,
        "Debit": int(debit),
        "Kredit": int(kredit)
    })

def hapus_transaksi(idx):
    try:
        st.session_state.transaksi.pop(int(idx))
    except Exception:
        pass

def buku_besar(df):
    akun_list = df["Akun"].unique()
    buku_besar_data = {}

    for akun in akun_list:
        df_akun = df[df["Akun"] == akun].copy()
        df_akun["Saldo"] = df_akun["Debit"].cumsum() - df_akun["Kredit"].cumsum()
        buku_besar_data[akun] = df_akun.reset_index(drop=True)

    return buku_besar_data

def neraca_saldo(df):
    grouped = df.groupby("Akun")[["Debit", "Kredit"]].sum()
    grouped["Saldo"] = grouped["Debit"] - grouped["Kredit"]
    return grouped

# ============================
# EXPORT EXCEL (Transaksi: 1 sheet, per-bulan di kolom dengan judul merged)
# ============================
def _write_block_transactions(ws, sub_df, col_start, headers, month_title):
    """Tulis satu blok bulan ke worksheet ws mulai dari kolom col_start.
       Mengembalikan next col_start setelah blok (ditambah 2 kolom gap)."""
    # merged title
    col_end = col_start + len(headers) - 1
    ws.merge_cells(start_row=1, start_column=col_start, end_row=1, end_column=col_end)
    cell = ws.cell(row=1, column=col_start)
    cell.value = month_title
    cell.font = Font(bold=True, size=13)
    cell.alignment = Alignment(horizontal='center', vertical='center')

    # header row (row 2)
    for i, h in enumerate(headers):
        c = ws.cell(row=2, column=col_start + i, value=h)
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal='center')

    # data rows start row 3
    for r_idx, row in enumerate(sub_df[headers].values, start=3):
        for c_idx, val in enumerate(row):
            cell = ws.cell(row=r_idx, column=col_start + c_idx, value=val)
            # align tanggal center for clarity
            if headers[c_idx].lower().startswith('tanggal'):
                cell.alignment = Alignment(horizontal='center')
            else:
                cell.alignment = Alignment(horizontal='left')

    return col_end + 2  # next start (gap 2 cols)

def export_excel_multi(df):
    """Menerima df transaksi (Tanggal,Akun,Keterangan,Debit,Kredit) dan mengembalikan bytes xlsx."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Transaksi"

    # pastikan kolom tanggal bertipe datetime
    df2 = df.copy()
    df2['Tanggal'] = pd.to_datetime(df2['Tanggal'])
    # month title format: "Mmm YYYY" atau "FullMonth YYYY"
    df2['BulanTitle'] = df2['Tanggal'].dt.strftime('%B %Y')

    headers = ['Tanggal', 'Akun', 'Keterangan', 'Debit', 'Kredit']
    bulan_order = df2['BulanTitle'].drop_duplicates().tolist()

    col_start = 1
    max_rows = 2
    for bln in bulan_order:
        sub = df2[df2['BulanTitle'] == bln].copy()
        if sub.empty:
            col_start += len(headers) + 2
            continue
        sub = sub.sort_values('Tanggal')
        sub_display = sub.copy()
        sub_display['Tanggal'] = sub_display['Tanggal'].dt.strftime('%Y-%m-%d')
        # write block
        col_start = _write_block_transactions(ws, sub_display, col_start, headers, bln)
        max_rows = max(max_rows, 2 + len(sub_display))

    # add thin border to used range
    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    last_col = col_start
    for row in ws.iter_rows(min_row=1, max_row=max_rows, min_col=1, max_col=last_col):
        for cell in row:
            if cell.value is not None:
                cell.border = border

    # Sheet Buku Besar
    ws2 = wb.create_sheet("Buku Besar")
    buku = buku_besar(df)
    r = 1
    for akun, data in buku.items():
        ws2.cell(row=r, column=1, value=akun).font = Font(bold=True, size=13)
        r += 1
        # header
        for ci, h in enumerate(list(data.columns), start=1):
            c = ws2.cell(row=r, column=ci, value=h)
            c.font = Font(bold=True)
            c.alignment = Alignment(horizontal='center')
        r += 1
        # data
        for _, rowdata in data.iterrows():
            for ci, val in enumerate(rowdata, start=1):
                ws2.cell(row=r, column=ci, value=val)
            r += 1
        r += 2

    # Sheet Neraca Saldo
    ws3 = wb.create_sheet("Neraca Saldo")
    ner = neraca_saldo(df)
    # header
    for ci, h in enumerate(list(ner.columns), start=1):
        c = ws3.cell(row=1, column=ci, value=h)
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal='center')
    # data
    rr = 2
    for idx in ner.index:
        for ci, val in enumerate(ner.loc[idx], start=1):
            ws3.cell(row=rr, column=ci, value=val)
        rr += 1

    # Auto width for all sheets
    for sheet in wb.worksheets:
        for col in sheet.columns:
            try:
                max_length = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                sheet.column_dimensions[col_letter].width = max_length + 3
            except Exception:
                pass

    # save to bytes
    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream.getvalue()

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
    with st.container():
        st.markdown("<div class='card'>", unsafe_allow_html=True)
        akun_list = [
            "Kas","Piutang","Utang","Modal","Pendapatan Jasa",
            "Beban Gaji","Beban Listrik","Beban Sewa"
        ]

        tanggal = st.date_input("Tanggal", datetime.now())
        akun = st.selectbox("Akun", akun_list)
        ket = st.text_input("Keterangan")

        col1, col2 = st.columns([1,1])
        with col1:
            debit = st.number_input("Debit (Rp)", min_value=0, step=1000, format="%d")
        with col2:
            kredit = st.number_input("Kredit (Rp)", min_value=0, step=1000, format="%d")

        col_a, col_b = st.columns([1,1])
        with col_a:
            if st.button("Tambah Transaksi"):
                tambah_transaksi(tanggal, akun, ket, debit, kredit)
                st.success("Transaksi berhasil ditambahkan!")
        with col_b:
            if st.button("Bersihkan Semua"):
                st.session_state.transaksi = []
                st.warning("Semua transaksi dihapus!")

        st.markdown("</div>", unsafe_allow_html=True)

    st.write("### 📄 Daftar Transaksi")
    if len(st.session_state.transaksi) > 0:
        df = pd.DataFrame(st.session_state.transaksi)
        df_display = df.copy()
        df_display["Debit"] = df_display["Debit"].apply(to_rp)
        df_display["Kredit"] = df_display["Kredit"].apply(to_rp)
        st.dataframe(df_display, use_container_width=True)

        cols = st.columns([1,1,1])
        with cols[0]:
            idx = st.number_input("Index hapus (0..n-1)", min_value=0, max_value=len(df)-1, value=0, step=1)
        with cols[1]:
            if st.button("Hapus Index"):
                hapus_transaksi(idx)
                st.success(f"Transaksi index {idx} dihapus.")
        with cols[2]:
            if st.button("Export Preview (CSV)"):
                csv = df.to_csv(index=False).encode('utf-8')
                st.download_button("Download CSV Preview", csv, "transaksi_preview.csv", "text/csv")
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
# 5. Grafik
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
        ).properties(title="Grafik Jumlah Debit per Akun", width=700)
        st.altair_chart(chart, use_container_width=True)

# ============================
# EXPORT EXCEL
# ============================
elif menu == "Export Excel":
    st.markdown("<div class='subtitle'>📤 Export Excel</div>", unsafe_allow_html=True)
    if len(st.session_state.transaksi) == 0:
        st.info("Belum ada transaksi untuk diekspor.")
    else:
        df = pd.DataFrame(st.session_state.transaksi)
        excel_file = export_excel_multi(df)

        st.download_button(
            label="📥 Export ke Excel (Transaksi 1 sheet - Bulan per kolom)",
            data=excel_file,
            file_name="laporan_akuntansi_lengkap.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
