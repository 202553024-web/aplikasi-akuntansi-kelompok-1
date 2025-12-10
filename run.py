import streamlit as st
import pandas as pd
import altair as alt
from datetime import datetime
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Aplikasi Akuntansi Modern", page_icon="💰", layout="wide")

# ================================
# CSS MODERN
# ================================
st.markdown("""
<style>
    .title { font-size: 40px; font-weight: 900; color: #0d47a1; text-align:center; }
    .subtitle { font-size: 23px; font-weight: 700; color:#0d47a1; margin-top: 10px; }
    .card {
        padding: 20px;
        border-radius: 14px;
        background: linear-gradient(180deg,#ffffff 0%, #f5f7fa 100%);
        box-shadow: 1px 3px 8px rgba(0,0,0,0.06);
        border: 1px solid rgba(13,71,161,0.06);
    }
    .stButton>button {
        background-color: #0d47a1 !important;
        color: white !important;
        padding: 10px 22px;
        border-radius: 12px;
        font-size: 18px;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("<div class='title'>📊 Aplikasi Akuntansi Modern</div>", unsafe_allow_html=True)
st.write("")

# ================================
# SESSION STATE
# ================================
if "transaksi" not in st.session_state:
    st.session_state.transaksi = []

def to_rp(n):
    try:
        return "Rp {:,}".format(int(n)).replace(",", ".")
    except:
        return "Rp 0"


# =====================================================
# EXPORT EXCEL FUNGSI
# =====================================================
def _write_block_transactions(ws, df_block, col_start, headers):
    col_end = col_start + len(headers) - 1
    ws.merge_cells(start_row=1, start_column=col_start, end_row=1, end_column=col_end)
    month_cell = ws.cell(row=1, column=col_start)
    month_cell.value = df_block['Bulan'].iloc[0]
    month_cell.font = Font(bold=True, size=14)
    month_cell.alignment = Alignment(horizontal='center', vertical='center')

    for i, h in enumerate(headers):
        c = ws.cell(row=2, column=col_start + i, value=h)
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal='center')

    for r_idx, row in enumerate(df_block[headers].values, start=3):
        for c_idx, val in enumerate(row):
            cell = ws.cell(row=r_idx, column=col_start + c_idx, value=val)
            if headers[c_idx].lower().startswith('tanggal'):
                cell.alignment = Alignment(horizontal='center')
            else:
                cell.alignment = Alignment(horizontal='left')

    return col_end + 2


def export_excel_multi(df, buku, neraca):
    wb = Workbook()

    # ================= SHEET 1 =================
    ws = wb.active
    ws.title = "Transaksi"

    df['Tanggal'] = pd.to_datetime(df['Tanggal'])
    df['Bulan'] = df['Tanggal'].dt.strftime('%B %Y')
    bulan_order = list(df['Bulan'].drop_duplicates().values)

    headers = ['Tanggal', 'Akun', 'Keterangan', 'Debit', 'Kredit']

    col_start = 1
    max_rows = 0

    for bln in bulan_order:
        sub = df[df['Bulan'] == bln].copy()
        if sub.empty:
            col_start += len(headers) + 2
            continue

        sub['Tanggal'] = sub['Tanggal'].dt.strftime('%Y-%m-%d')
        sub.insert(0, 'Bulan', bln)

        col_start = _write_block_transactions(ws, sub, col_start, headers)
        max_rows = max(max_rows, 2 + len(sub))

    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for row in ws.iter_rows(min_row=1, max_row=max_rows, min_col=1, max_col=col_start):
        for cell in row:
            cell.border = border

    # ================= SHEET 2 =================
    ws2 = wb.create_sheet("Buku Besar")
    r = 1
    for akun, data in buku.items():
        ws2.cell(row=r, column=1, value=akun).font = Font(bold=True, size=13)
        r += 1
        for ci, h in enumerate(list(data.columns), start=1):
            c = ws2.cell(row=r, column=ci, value=h)
            c.font = Font(bold=True)
            c.alignment = Alignment(horizontal='center')
        r += 1
        for _, row_data in data.iterrows():
            for ci, val in enumerate(row_data, start=1):
                ws2.cell(row=r, column=ci, value=val)
            r += 1
        r += 2

    # ================= SHEET 3 =================
    ws3 = wb.create_sheet("Neraca Saldo")
    ner_headers = list(neraca.columns)
    for ci, h in enumerate(ner_headers, start=1):
        c = ws3.cell(row=1, column=ci, value=h)
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal='center')

    r = 2
    for idx in neraca.index:
        for ci, val in enumerate(neraca.loc[idx], start=1):
            ws3.cell(row=r, column=ci, value=val)
        r += 1

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

    file_data = io.BytesIO()
    wb.save(file_data)
    return file_data.getvalue()


# =====================================================
# MENU
# =====================================================
menu = st.sidebar.radio("📌 PILIH MENU", ["Input Transaksi", "Jurnal", "Buku Besar", "Neraca Saldo", "Grafik", "Export Excel"])


# =====================================================
# INPUT TRANSAKSI
# =====================================================
if menu == "Input Transaksi":
    st.markdown("<div class='subtitle'>📝 Input Transaksi</div>", unsafe_allow_html=True)

    akun_list = ["Kas", "Piutang", "Utang", "Modal", "Pendapatan Jasa", "Beban Gaji", "Beban Listrik", "Beban Sewa"]

    with st.container():
        st.markdown("<div class='card'>", unsafe_allow_html=True)
        tanggal = st.date_input("Tanggal", datetime.now())
        akun = st.selectbox("Akun", akun_list)
        ket = st.text_input("Keterangan")
        col1, col2 = st.columns(2)
        debit = col1.number_input("Debit", min_value=0, step=1000)
        kredit = col2.number_input("Kredit", min_value=0, step=1000)

        if st.button("Tambah"):
            st.session_state.transaksi.append({
                "Tanggal": str(tanggal),
                "Akun": akun,
                "Keterangan": ket,
                "Debit": int(debit),
                "Kredit": int(kredit)
            })
            st.success("Transaksi ditambahkan!")
        st.markdown("</div>", unsafe_allow_html=True)

    st.write("### 📄 Daftar Transaksi")

    if len(st.session_state.transaksi) > 0:
        df = pd.DataFrame(st.session_state.transaksi)
        df_disp = df.copy()
        df_disp['Debit'] = df_disp['Debit'].apply(lambda x: to_rp(x))
        df_disp['Kredit'] = df_disp['Kredit'].apply(lambda x: to_rp(x))
        st.dataframe(df_disp, use_container_width=True)


# =====================================================
# JURNAL
# =====================================================
elif menu == "Jurnal":
    st.markdown("<div class='subtitle'>📘 Jurnal</div>", unsafe_allow_html=True)
    if len(st.session_state.transaksi) == 0:
        st.info("Belum ada data.")
    else:
        df = pd.DataFrame(st.session_state.transaksi)
        df2 = df.copy()
        df2['Debit'] = df2['Debit'].apply(lambda x: to_rp(x))
        df2['Kredit'] = df2['Kredit'].apply(lambda x: to_rp(x))
        st.dataframe(df2, use_container_width=True)


# =====================================================
# BUKU BESAR
# =====================================================
elif menu == "Buku Besar":
    st.markdown("<div class='subtitle'>📗 Buku Besar</div>", unsafe_allow_html=True)

    if len(st.session_state.transaksi) == 0:
        st.info("Belum ada data.")
    else:
        df = pd.DataFrame(st.session_state.transaksi)
        buku = {}

        for akun in df["Akun"].unique():
            sub = df[df["Akun"] == akun].copy()
            sub["Saldo"] = sub["Debit"].cumsum() - sub["Kredit"].cumsum()
            buku[akun] = sub

            st.write(f"### ▶ {akun}")
            df_disp = sub.copy()
            df_disp["Debit"] = df_disp["Debit"].apply(lambda x: to_rp(x))
            df_disp["Kredit"] = df_disp["Kredit"].apply(lambda x: to_rp(x))
            df_disp["Saldo"] = df_disp["Saldo"].apply(lambda x: to_rp(x))
            st.dataframe(df_disp)


# =====================================================
# NERACA SALDO
# =====================================================
elif menu == "Neraca Saldo":
    st.markdown("<div class='subtitle'>📙 Neraca Saldo</div>", unsafe_allow_html=True)

    if len(st.session_state.transaksi) == 0:
        st.info("Belum ada data.")
    else:
        df = pd.DataFrame(st.session_state.transaksi)
        neraca = df.groupby("Akun")[['Debit', 'Kredit']].sum()
        neraca['Saldo'] = neraca['Debit'] - neraca['Kredit']

        ner_disp = neraca.copy()
        ner_disp['Debit'] = ner_disp['Debit'].apply(lambda x: to_rp(x))
        ner_disp['Kredit'] = ner_disp['Kredit'].apply(lambda x: to_rp(x))
        ner_disp['Saldo'] = ner_disp['Saldo'].apply(lambda x: to_rp(x))

        st.dataframe(ner_disp)


# =====================================================
# GRAFIK
# =====================================================
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


# =====================================================
# EXPORT EXCEL
# =====================================================
elif menu == "Export Excel":
    st.markdown("<div class='subtitle'>📤 Export Excel (Rapi)</div>", unsafe_allow_html=True)

    if len(st.session_state.transaksi) == 0:
        st.info("Belum ada transaksi.")
    else:
        df = pd.DataFrame(st.session_state.transaksi)

        buku = {}
        for akun in df["Akun"].unique():
            sub = df[df["Akun"] == akun].copy()
            sub["Saldo"] = sub["Debit"].cumsum() - sub["Kredit"].cumsum()
            buku[akun] = sub

        neraca = df.groupby("Akun")[['Debit', 'Kredit']].sum()
        neraca['Saldo'] = neraca['Debit'] - neraca['Kredit']

        file_xlsx = export_excel_multi(df, buku, neraca)

        st.download_button(
            "📥 Download Excel",
            file_xlsx,
            file_name="Laporan_Akuntansi_Rapi.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
