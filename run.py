import streamlit as st
import pandas as pd
import altair as alt
from datetime import datetime
import io
import calendar
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

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
# FUNGSI EXPORT EXCEL (REVISI FINAL)
# ============================
def export_excel_multi(df):
    import io, calendar
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.utils.dataframe import dataframe_to_rows

    output = io.BytesIO()
    wb = Workbook()
    
    # Warna dan Style
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    subheader_fill = PatternFill(start_color="D6DCE4", end_color="D6DCE4", fill_type="solid")
    subheader_blue_fill = PatternFill(start_color="B4C7E7", end_color="B4C7E7", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True, size=11)
    black_font = Font(color="000000", bold=True, size=11)
    
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # Persiapan Data
    df["Tanggal"] = pd.to_datetime(df["Tanggal"])
    df["Bulan"] = df["Tanggal"].dt.month
    df["Tahun"] = df["Tanggal"].dt.year
    df_sorted = df.sort_values("Tanggal")

    # ============================
    # SHEET 1: LAPORAN KEUANGAN
    # ============================
    ws_main = wb.active
    ws_main.title = "Laporan Keuangan"
    current_row = 1
    tahun_sekarang = None

    for (tahun, bulan), grup in df_sorted.groupby(["Tahun", "Bulan"]):
        # Header Tahun
        if tahun != tahun_sekarang:
            ws_main.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=5)
            cell = ws_main.cell(row=current_row, column=1, value=f"Laporan Keuangan Tahun {tahun}")
            cell.font = Font(bold=True, size=12)
            cell.alignment = Alignment(horizontal="center")
            cell.fill = subheader_fill
            for col in range(1, 6):
                ws_main.cell(row=current_row, column=col).border = thin_border
            current_row += 1
            tahun_sekarang = tahun

        # Header Bulan
        nama_bulan = calendar.month_name[bulan].capitalize()
        ws_main.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=5)
        cell = ws_main.cell(row=current_row, column=1, value=f"Bulan {nama_bulan}")
        cell.font = black_font
        cell.alignment = Alignment(horizontal="center")
        cell.fill = subheader_fill
        for col in range(1, 6):
            ws_main.cell(row=current_row, column=col).border = thin_border
        current_row += 1

        # Header Kolom
        headers = ["Tanggal", "Akun", "Keterangan", "Debit", "Kredit"]
        for col_num, header in enumerate(headers, start=1):
            cell = ws_main.cell(row=current_row, column=col_num, value=header)
            cell.font = white_font
            cell.alignment = Alignment(horizontal="center")
            cell.fill = header_fill
            cell.border = thin_border
        current_row += 1
        
        # Data Transaksi
        for r in dataframe_to_rows(grup[headers], index=False, header=False):
            for c_idx, val in enumerate(r, start=1):
                cell = ws_main.cell(row=current_row, column=c_idx)
                cell.border = thin_border
                if c_idx in [4, 5]:
                    val = int(val) if pd.notna(val) and val != 0 else None
                    if val is None:
                        cell.value = "-"
                        cell.alignment = Alignment(horizontal="center")
                    else:
                        cell.value = val
                        cell.alignment = Alignment(horizontal="right")
                        cell.number_format = '#,##0'
                else:
                    cell.value = val
                    if c_idx == 1:
                        cell.alignment = Alignment(horizontal="center")
            current_row += 1
        current_row += 1

    # Auto-fit kolom
    for col in ws_main.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws_main.column_dimensions[column].width = adjusted_width

    # ============================
    # SHEET 2: JURNAL UMUM
    # ============================
    ws_jurnal = wb.create_sheet("Jurnal Umum")
    
    # Header Utama
    ws_jurnal.merge_cells('A1:E1')
    cell = ws_jurnal['A1']
    cell.value = "Jurnal Umum"
    cell.font = Font(bold=True, size=12)
    cell.alignment = Alignment(horizontal="center")
    
    # Header Kolom
    headers = ["Tanggal", "Akun", "Keterangan", "Debit", "Kredit"]
    for col_num, header in enumerate(headers, start=1):
        cell = ws_jurnal.cell(row=3, column=col_num, value=header)
        cell.font = white_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
        cell.border = thin_border
        
    # Data
    for i, r in enumerate(dataframe_to_rows(df[headers], index=False, header=False), start=4):
        for c_idx, val in enumerate(r, start=1):
            cell = ws_jurnal.cell(row=i, column=c_idx)
            cell.border = thin_border
            if c_idx in [4, 5]:
                val = int(val) if pd.notna(val) and val != 0 else None
                if val is None:
                    cell.value = "-"
                    cell.alignment = Alignment(horizontal="center")
                else:
                    cell.value = val
                    cell.alignment = Alignment(horizontal="right")
                    cell.number_format = '#,##0'
            else:
                cell.value = val
                if c_idx == 1:
                    cell.alignment = Alignment(horizontal="center")
    
    # Auto-fit kolom
    for col in ws_jurnal.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws_jurnal.column_dimensions[column].width = adjusted_width

    # ============================
    # SHEET 3: BUKU BESAR
    # ============================
    ws_bb = wb.create_sheet("Buku Besar")
    
    # Header Utama
    ws_bb.merge_cells('A1:F1')
    cell = ws_bb['A1']
    cell.value = "Buku Besar"
    cell.font = Font(bold=True, size=12)
    cell.alignment = Alignment(horizontal="center")
    
    akun_list = df["Akun"].unique()
    row_bb = 3

    for akun in akun_list:
        # Nama Akun
        ws_bb.merge_cells(start_row=row_bb, start_column=1, end_row=row_bb, end_column=2)
        cell = ws_bb.cell(row=row_bb, column=1, value=f"Nama Akun :")
        cell.font = black_font
        cell.fill = subheader_blue_fill
        cell.border = thin_border
        cell = ws_bb.cell(row=row_bb, column=2)
        cell.fill = subheader_blue_fill
        cell.border = thin_border
        
        cell = ws_bb.cell(row=row_bb, column=3, value=akun)
        cell.font = black_font
        ws_bb.merge_cells(start_row=row_bb, start_column=3, end_row=row_bb, end_column=5)
        for col in range(3, 6):
            ws_bb.cell(row=row_bb, column=col).border = thin_border
            ws_bb.cell(row=row_bb, column=col).fill = subheader_blue_fill
        row_bb += 1

        # Header Kolom
        headers = ["Tanggal", "Keterangan", "Debit", "Kredit", "Saldo"]
        for col_num, header in enumerate(headers, start=1):
            cell = ws_bb.cell(row=row_bb, column=col_num, value=header)
            cell.font = white_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center")
            cell.border = thin_border
        row_bb += 1

        # Data
        df_akun = df[df["Akun"] == akun].copy()
        df_akun["Saldo"] = df_akun["Debit"].cumsum() - df_akun["Kredit"].cumsum()
        
        for r in dataframe_to_rows(df_akun[headers], index=False, header=False):
            for c_idx, val in enumerate(r, start=1):
                cell = ws_bb.cell(row=row_bb, column=c_idx)
                cell.border = thin_border
                if c_idx >= 3:
                    val = int(val) if pd.notna(val) and val != 0 else None
                    if val is None or val == 0:
                        if c_idx in [3, 4]:
                            cell.value = "-"
                            cell.alignment = Alignment(horizontal="center")
                        else:
                            cell.value = val if val is not None else 0
                            cell.alignment = Alignment(horizontal="right")
                            cell.number_format = '#,##0'
                    else:
                        cell.value = val
                        cell.alignment = Alignment(horizontal="right")
                        cell.number_format = '#,##0'
                else:
                    cell.value = val
                    if c_idx == 1:
                        cell.alignment = Alignment(horizontal="center")
            row_bb += 1
        row_bb += 2

    # Auto-fit kolom
    for col in ws_bb.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws_bb.column_dimensions[column].width = adjusted_width

    # ============================
    # SHEET 4: NERACA SALDO
    # ============================
    ws_ns = wb.create_sheet("Neraca Saldo")
    
    # Header Utama
    ws_ns.merge_cells('A1:D1')
    cell = ws_ns['A1']
    cell.value = "Neraca Saldo"
    cell.font = Font(bold=True, size=12)
    cell.alignment = Alignment(horizontal="center")
    
    # Header Kolom
    headers = ["Akun", "Debit", "Kredit", "Saldo"]
    for col_num, header in enumerate(headers, start=1):
        cell = ws_ns.cell(row=3, column=col_num, value=header)
        cell.font = white_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
        cell.border = thin_border
    
    # Data
    neraca = df.groupby("Akun")[["Debit", "Kredit"]].sum().reset_index()
    neraca["Saldo"] = neraca["Debit"] - neraca["Kredit"]
    
    for i, r in enumerate(dataframe_to_rows(neraca[headers], index=False, header=False), start=4):
        for c_idx, val in enumerate(r, start=1):
            cell = ws_ns.cell(row=i, column=c_idx)
            cell.border = thin_border
            if c_idx >= 2:
                val = int(val) if pd.notna(val) and val != 0 else None
                if val is None:
                    cell.value = "-"
                    cell.alignment = Alignment(horizontal="center")
                else:
                    cell.value = val
                    cell.alignment = Alignment(horizontal="right")
                    cell.number_format = '#,##0'
            else:
                cell.value = val

    # Auto-fit kolom
    for col in ws_ns.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws_ns.column_dimensions[column].width = adjusted_width

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
