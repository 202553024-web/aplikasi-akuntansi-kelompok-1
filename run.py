import streamlit as st
import pandas as pd
import altair as alt
from datetime import datetime
import io
import calendar
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

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
# FUNGSI EXPORT EXCEL (DIPERBAIKI)
# ============================
def export_excel_multi(df):
    import io, calendar
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.utils.dataframe import dataframe_to_rows

    output = io.BytesIO()
    wb = Workbook()
    
    # Definisi Border
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Definisi Warna
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    title_fill = PatternFill(start_color="B4C7E7", end_color="B4C7E7", fill_type="solid")
    year_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    
    # ============================
    # SHEET 1: LAPORAN KEUANGAN
    # ============================
    ws_main = wb.active
    ws_main.title = "Laporan Keuangan"

    df["Tanggal"] = pd.to_datetime(df["Tanggal"])
    df["Bulan"] = df["Tanggal"].dt.month
    df["Tahun"] = df["Tanggal"].dt.year
    df_sorted = df.sort_values("Tanggal")

    current_row = 1
    tahun_sekarang = None

    for (tahun, bulan), grup in df_sorted.groupby(["Tahun", "Bulan"]):
        # Header Tahun
        if tahun != tahun_sekarang:
            ws_main.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=5)
            cell = ws_main.cell(row=current_row, column=1, value=f"Laporan Keuangan Tahun {tahun}")
            cell.font = Font(bold=True, size=14)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.fill = year_fill
            for col in range(1, 6):
                ws_main.cell(row=current_row, column=col).border = thin_border
            current_row += 1
            tahun_sekarang = tahun

        # Header Bulan
        nama_bulan = calendar.month_name[bulan].capitalize()
        ws_main.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=5)
        cell = ws_main.cell(row=current_row, column=1, value=f"Bulan {nama_bulan}")
        cell.font = Font(bold=True, size=11)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = title_fill
        for col in range(1, 6):
            ws_main.cell(row=current_row, column=col).border = thin_border
        current_row += 1

        # Header Kolom
        headers = ["Tanggal", "Akun", "Keterangan", "Debit", "Kredit"]
        for col_num, header in enumerate(headers, start=1):
            cell = ws_main.cell(row=current_row, column=col_num, value=header)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.fill = header_fill
            cell.border = thin_border
        current_row += 1
        
        # Data Transaksi
        for r in dataframe_to_rows(grup[["Tanggal", "Akun", "Keterangan", "Debit", "Kredit"]], index=False, header=False):
            for c_idx, val in enumerate(r, start=1):
                cell = ws_main.cell(row=current_row, column=c_idx)
                cell.border = thin_border
                
                if c_idx in [4, 5]:  # Debit/Kredit
                    val = int(val) if pd.notna(val) and val != 0 else 0
                    if val == 0:
                        cell.value = "Rp                    -"
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    else:
                        cell.value = val
                        cell.number_format = '"Rp"#,##0.00'
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.value = val
                    cell.alignment = Alignment(horizontal="left", vertical="center")
            current_row += 1
        current_row += 1

    # Set Lebar Kolom
    ws_main.column_dimensions['A'].width = 20
    ws_main.column_dimensions['B'].width = 18
    ws_main.column_dimensions['C'].width = 20
    ws_main.column_dimensions['D'].width = 20
    ws_main.column_dimensions['E'].width = 20

    # ============================
    # SHEET 2: JURNAL UMUM
    # ============================
    ws_jurnal = wb.create_sheet("Jurnal Umum")
    
    # Title
    ws_jurnal.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5)
    title_cell = ws_jurnal.cell(row=1, column=1, value="Jurnal Umum")
    title_cell.font = Font(bold=True, size=14)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    title_cell.fill = year_fill
    for col in range(1, 6):
        ws_jurnal.cell(row=1, column=col).border = thin_border
    
    # Header
    headers = ["Tanggal", "Akun", "Keterangan", "Debit", "Kredit"]
    for col_num, header in enumerate(headers, start=1):
        cell = ws_jurnal.cell(row=3, column=col_num, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = header_fill
        cell.border = thin_border
    
    # Data
    for i, r in enumerate(dataframe_to_rows(df[headers], index=False, header=False), start=4):
        for c_idx, val in enumerate(r, start=1):
            cell = ws_jurnal.cell(row=i, column=c_idx)
            cell.border = thin_border
            
            if c_idx in [4, 5]:
                val = int(val) if pd.notna(val) and val != 0 else 0
                if val == 0:
                    cell.value = "Rp                    -"
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.value = val
                    cell.number_format = '"Rp"#,##0.00'
                    cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.value = val
                cell.alignment = Alignment(horizontal="left", vertical="center")
    
    ws_jurnal.column_dimensions['A'].width = 20
    ws_jurnal.column_dimensions['B'].width = 18
    ws_jurnal.column_dimensions['C'].width = 20
    ws_jurnal.column_dimensions['D'].width = 20
    ws_jurnal.column_dimensions['E'].width = 20

    # ============================
    # SHEET 3: BUKU BESAR
    # ============================
    ws_bb = wb.create_sheet("Buku Besar")
    
    # Title
    ws_bb.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5)
    title_cell = ws_bb.cell(row=1, column=1, value="Buku Besar")
    title_cell.font = Font(bold=True, size=14)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    title_cell.fill = year_fill
    for col in range(1, 6):
        ws_bb.cell(row=1, column=col).border = thin_border
    
    akun_list = df["Akun"].unique()
    row_bb = 3

    for akun in akun_list:
        # Nama Akun
        ws_bb.merge_cells(start_row=row_bb, start_column=1, end_row=row_bb, end_column=2)
        cell = ws_bb.cell(row=row_bb, column=1, value=f"Nama Akun :")
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="left", vertical="center")
        cell.fill = title_fill
        cell.border = thin_border
        ws_bb.cell(row=row_bb, column=2).border = thin_border
        
        cell_akun = ws_bb.cell(row=row_bb, column=3, value=akun)
        cell_akun.font = Font(bold=False)
        cell_akun.alignment = Alignment(horizontal="left", vertical="center")
        cell_akun.fill = title_fill
        
        ws_bb.merge_cells(start_row=row_bb, start_column=3, end_row=row_bb, end_column=5)
        for col in range(3, 6):
            ws_bb.cell(row=row_bb, column=col).border = thin_border
            ws_bb.cell(row=row_bb, column=col).fill = title_fill
        row_bb += 1

        # Header
        df_akun = df[df["Akun"] == akun].copy()
        df_akun["Saldo"] = df_akun["Debit"].cumsum() - df_akun["Kredit"].cumsum()
        headers = ["Tanggal", "Keterangan", "Debit", "Kredit", "Saldo"]

        for col_num, header in enumerate(headers, start=1):
            cell = ws_bb.cell(row=row_bb, column=col_num, value=header)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.fill = header_fill
            cell.border = thin_border
        row_bb += 1

        # Data
        for r in dataframe_to_rows(df_akun[headers], index=False, header=False):
            for c_idx, val in enumerate(r, start=1):
                cell = ws_bb.cell(row=row_bb, column=c_idx)
                cell.border = thin_border
                
                if c_idx >= 3:
                    val = int(val) if pd.notna(val) and val != 0 else 0
                    if val == 0 and c_idx in [3, 4]:
                        cell.value = "Rp                    -"
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    else:
                        cell.value = val
                        cell.number_format = '"Rp"#,##0.00'
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.value = val
                    cell.alignment = Alignment(horizontal="left", vertical="center")
            row_bb += 1
        row_bb += 1

    ws_bb.column_dimensions['A'].width = 20
    ws_bb.column_dimensions['B'].width = 20
    ws_bb.column_dimensions['C'].width = 18
    ws_bb.column_dimensions['D'].width = 18
    ws_bb.column_dimensions['E'].width = 18

    # ============================
    # SHEET 4: NERACA SALDO
    # ============================
    ws_ns = wb.create_sheet("Neraca Saldo")
    
    # Title
    ws_ns.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
    title_cell = ws_ns.cell(row=1, column=1, value="Neraca Saldo")
    title_cell.font = Font(bold=True, size=14)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    title_cell.fill = year_fill
    for col in range(1, 5):
        ws_ns.cell(row=1, column=col).border = thin_border
    
    # Header
    neraca = df.groupby("Akun")[["Debit", "Kredit"]].sum().reset_index()
    neraca["Saldo"] = neraca["Debit"] - neraca["Kredit"]

    headers = ["Akun", "Debit", "Kredit", "Saldo"]
    for col_num, header in enumerate(headers, start=1):
        cell = ws_ns.cell(row=3, column=col_num, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = header_fill
        cell.border = thin_border
    
    # Data
    for i, r in enumerate(dataframe_to_rows(neraca[headers], index=False, header=False), start=4):
        for c_idx, val in enumerate(r, start=1):
            cell = ws_ns.cell(row=i, column=c_idx)
            cell.border = thin_border
            
            if c_idx >= 2:
                val = int(val) if pd.notna(val) and val != 0 else 0
                if val == 0 and c_idx in [2, 3]:
                    cell.value = "Rp                    -"
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.value = val
                    cell.number_format = '"Rp"#,##0.00'
                    cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.value = val
                cell.alignment = Alignment(horizontal="left", vertical="center")

    ws_ns.column_dimensions['A'].width = 20
    ws_ns.column_dimensions['B'].width = 20
    ws_ns.column_dimensions['C'].width = 20
    ws_ns.column_dimensions['D'].width = 20

    wb.save(output)
    output.seek(0)
    return output.getvalue()

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
        df["Tanggal"] = pd.to_datetime(df["Tanggal"])
        df["Bulan"] = df["Tanggal"].dt.month
        df["Tahun"] = df["Tanggal"].dt.year

        tahun_sekarang = None
        for (tahun, bulan), grup in df.groupby(["Tahun", "Bulan"]):
            # Header Tahun
            if tahun != tahun_sekarang:
                st.markdown(f"### 📅 Tahun {tahun}")
                tahun_sekarang = tahun

            # Header Bulan
            nama_bulan = calendar.month_name[bulan].capitalize()
            st.markdown(f"#### 📌 Bulan {nama_bulan}")

            df_show = grup.copy()
            df_show["Debit"] = df_show["Debit"].apply(to_rp)
            df_show["Kredit"] = df_show["Kredit"].apply(to_rp)

            st.dataframe(df_show[["Tanggal", "Akun", "Keterangan", "Debit", "Kredit"]], use_container_width=True)

# ============================
# 3. BUKU BESAR
# ============================
elif menu == "Buku Besar":
    st.markdown("<div class='subtitle'>📗 Buku Besar</div>", unsafe_allow_html=True)

    if len(st.session_state.transaksi) == 0:
        st.info("Belum ada data.")
    else:
        df = pd.DataFrame(st.session_state.transaksi)
        df["Tanggal"] = pd.to_datetime(df["Tanggal"])
        df["Bulan"] = df["Tanggal"].dt.month
        df["Tahun"] = df["Tanggal"].dt.year

        tahun_sekarang = None
        for (tahun, bulan), grup in df.groupby(["Tahun", "Bulan"]):

            # Header Tahun
            if tahun != tahun_sekarang:
                st.markdown(f"### 📅 Tahun {tahun}")
                tahun_sekarang = tahun

            nama_bulan = calendar.month_name[bulan].capitalize()
            st.markdown(f"#### 📌 Bulan {nama_bulan}")

            # Akun per bulan
            buku = buku_besar(grup)
            for akun, data in buku.items():
                st.markdown(f"##### ▶ {akun}")

                df_show = data.copy()
                df_show["Debit"] = df_show["Debit"].apply(to_rp)
                df_show["Kredit"] = df_show["Kredit"].apply(to_rp)
                df_show["Saldo"] = df_show["Saldo"].apply(to_rp)

                st.dataframe(df_show[["Tanggal", "Keterangan", "Debit", "Kredit", "Saldo"]], use_container_width=True)
                st.write("---")

# ============================
# 4. NERACA SALDO
# ============================
elif menu == "Neraca Saldo":
    st.markdown("<div class='subtitle'>📙 Neraca Saldo</div>", unsafe_allow_html=True)

    if len(st.session_state.transaksi) == 0:
        st.info("Belum ada data.")
    else:
        df = pd.DataFrame(st.session_state.transaksi)
        df["Tanggal"] = pd.to_datetime(df["Tanggal"])
        df["Bulan"] = df["Tanggal"].dt.month
        df["Tahun"] = df["Tanggal"].dt.year

        tahun_sekarang = None
        for (tahun, bulan), grup in df.groupby(["Tahun", "Bulan"]):
            
            # Header Tahun
            if tahun != tahun_sekarang:
                st.markdown(f"### 📅 Tahun {tahun}")
                tahun_sekarang = tahun

            nama_bulan = calendar.month_name[bulan].capitalize()
            st.markdown(f"#### 📌 Bulan {nama_bulan}")

            neraca = grup.groupby("Akun")[["Debit", "Kredit"]].sum()
            neraca["Saldo"] = neraca["Debit"] - neraca["Kredit"]

            df_show = neraca.copy()
            df_show["Debit"] = df_show["Debit"].apply(to_rp)
            df_show["Kredit"] = df_show["Kredit"].apply(to_rp)
            df_show["Saldo"] = df_show["Saldo"].apply(to_rp)

            st.dataframe(df_show, use_container_width=True)
            st.write("---")

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
# 6. EXPORT EXCEL (MULTI SHEET)
# ============================
elif menu == "Export Excel":
    st.markdown("<div class='subtitle'>📤 Export Excel (Multi Sheet)</div>", unsafe_allow_html=True)

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
