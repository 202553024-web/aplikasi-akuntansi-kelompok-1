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
# KLASIFIKASI AKUN
# ============================
pendapatan_akun = ["Pendapatan Jasa"]
beban_akun = ["Beban Gaji", "Beban Listrik", "Beban Sewa"]

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

def laporan_laba_rugi(df):
    total_pendapatan = df[df["Akun"].isin(pendapatan_akun)]["Kredit"].sum() - df[df["Akun"].isin(pendapatan_akun)]["Debit"].sum()
    total_beban = df[df["Akun"].isin(beban_akun)]["Debit"].sum() - df[df["Akun"].isin(beban_akun)]["Kredit"].sum()
    laba_rugi = total_pendapatan - total_beban
    return {
        "Total Pendapatan": total_pendapatan,
        "Total Beban": total_beban,
        "Laba/Rugi": laba_rugi
    }

# ============================
# FUNGSI EXPORT EXCEL (DIPERBAIKI DENGAN LABA RUGI)
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
                
                if c_idx in [4, 5]:  # Debit/Kredit
                    val = int(val) if pd.notna(val) and val != 0 else 0
                    if val == 0:
                        cell.value = "Rp                    -"
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
    # SHEET 2: JURNAL UMUM (DIKELOMPOKKAN PER BULAN)
    # ============================
    ws_jurnal = wb.create_sheet("Jurnal Umum")
    
    current_row_jurnal = 1
    tahun_sekarang_jurnal = None

    for (tahun, bulan), grup in df_sorted.groupby(["Tahun", "Bulan"]):
        # Title Jurnal Umum
        ws_jurnal.merge_cells(start_row=current_row_jurnal, start_column=1, end_row=current_row_jurnal, end_column=5)
        title_cell = ws_jurnal.cell(row=current_row_jurnal, column=1, value="Jurnal Umum")
        title_cell.font = Font(bold=True, size=14)
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        title_cell.fill = year_fill
        for col in range(1, 6):
            ws_jurnal.cell(row=current_row_jurnal, column=col).border = thin_border
        current_row_jurnal += 1

        # Periode Bulan dan Tahun
        nama_bulan = calendar.month_name[bulan].capitalize()
        ws_jurnal.merge_cells(start_row=current_row_jurnal, start_column=1, end_row=current_row_jurnal, end_column=5)
        periode_cell = ws_jurnal.cell(row=current_row_jurnal, column=1, value=f"Periode {nama_bulan} {tahun}")
        periode_cell.font = Font(bold=True, size=12)
        periode_cell.alignment = Alignment(horizontal="center", vertical="center")
        periode_cell.fill = year_fill
        for col in range(1, 6):
            ws_jurnal.cell(row=current_row_jurnal, column=col).border = thin_border
        current_row_jurnal += 2
        
        # Header
        headers = ["Tanggal", "Akun", "Keterangan", "Debit", "Kredit"]
        for col_num, header in enumerate(headers, start=1):
            cell = ws_jurnal.cell(row=current_row_jurnal, column=col_num, value=header)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.fill = header_fill
            cell.border = thin_border
        current_row_jurnal += 1
        
        # Data Transaksi
        total_debit = 0
        total_kredit = 0
        
        for r in dataframe_to_rows(grup[["Tanggal", "Akun", "Keterangan", "Debit", "Kredit"]], index=False, header=False):
            for c_idx, val in enumerate(r, start=1):
                cell = ws_jurnal.cell(row=current_row_jurnal, column=c_idx)
                cell.border = thin_border
                
                if c_idx in [4, 5]:
                    val = int(val) if pd.notna(val) and val != 0 else 0
                    if c_idx == 4:
                        total_debit += val
                    else:
                        total_kredit += val
                        
                    if val == 0:
                        cell.value = "Rp                    -"
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    else:
                        cell.value = val
                        cell.number_format = '"Rp"#,##0.00'
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.value = val
                    cell.alignment = Alignment(horizontal="left", vertical="center")
            current_row_jurnal += 1
        
        # Baris Total
        ws_jurnal.merge_cells(start_row=current_row_jurnal, start_column=1, end_row=current_row_jurnal, end_column=3)
        total_label_cell = ws_jurnal.cell(row=current_row_jurnal, column=1, value="Total")
        total_label_cell.font = Font(bold=True)
        total_label_cell.alignment = Alignment(horizontal="center", vertical="center")
        total_label_cell.fill = title_fill
        for col in range(1, 4):
            ws_jurnal.cell(row=current_row_jurnal, column=col).border = thin_border
            ws_jurnal.cell(row=current_row_jurnal, column=col).fill = title_fill
        
        # Total Debit
        cell_total_debit = ws_jurnal.cell(row=current_row_jurnal, column=4, value=total_debit)
        cell_total_debit.number_format = '"Rp"#,##0.00'
        cell_total_debit.alignment = Alignment(horizontal="right", vertical="center")
        cell_total_debit.fill = title_fill
        cell_total_debit.border = thin_border
        cell_total_debit.font = Font(bold=True)
        
        # Total Kredit
        cell_total_kredit = ws_jurnal.cell(row=current_row_jurnal, column=5, value=total_kredit)
        cell_total_kredit.number_format = '"Rp"#,##0.00'
        cell_total_kredit.alignment = Alignment(horizontal="right", vertical="center")
        cell_total_kredit.fill = title_fill
        cell_total_kredit.border = thin_border
        cell_total_kredit.font = Font(bold=True)
        
        current_row_jurnal += 2
    
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
                        cell.value = "Rp                    -"
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
                    cell.value = "Rp                    -"
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

# ============================
# 5. DEFINISI MENU (TAMBAHKAN INI!)
# ============================
# Bagian ini HARUS ada sebelum baris "if menu == ..."
menu = st.sidebar.selectbox("Navigasi", ["Input Transaksi", "Grafik", "Export Excel"])

# ============================
# 6. LOGIKA HALAMAN (UTAMA)
# ============================

if menu == "Input Transaksi":
    st.header("Halaman Input")
    # ... isi kode input Anda ...

# SEKARANG BARU BOLEH PAKAI ELIF
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
        ).properties(title="Grafik Jumlah Debit per Akun")
        st.altair_chart(chart, use_container_width=True)

    # ============================
    # SHEET 5: LAPORAN LABA RUGI
    # ============================
    ws_lr = wb.create_sheet("Laporan Laba Rugi")
    
    # Title
    ws_lr.merge_cells(start_row=1, start_column=1, end_row=1, end_column=3)
    title_cell = ws_lr.cell(row=1, column=1, value="Laporan Laba Rugi")
    title_cell.font = Font(bold=True, size=14)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    title_cell.fill = year_fill
    for col in range(1, 4):
        ws_lr.cell(row=1, column=col).border = thin_border
    
    # Data Laba Rugi
    laba_rugi_data = laporan_laba_rugi(df)
    
    row_lr = 3
    headers_lr = ["Keterangan", "Debit", "Kredit"]
    for col_num, header in enumerate(headers_lr, start=1):
        cell = ws_lr.cell(row=row_lr, column=col_num, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = header_fill
        cell.border = thin_border
    row_lr += 1
    
    # Pendapatan
    cell = ws_lr.cell(row=row_lr, column=1, value="Total Pendapatan")
    cell.alignment = Alignment(horizontal="left", vertical="center")
    cell.border = thin_border
    cell = ws_lr.cell(row=row_lr, column=2, value=laba_rugi_data["Total Pendapatan"])
    cell.number_format = '"Rp"#,##0.00'
    cell.alignment = Alignment(horizontal="right", vertical="center")
    cell.border = thin_border
    cell = ws_lr.cell(row=row_lr, column=3, value=0)
    cell.number_format = '"Rp"#,##0.00'
    cell.alignment = Alignment(horizontal="right", vertical="center")
    cell.border = thin_border
    row_lr += 1
    
    # Beban
    cell = ws_lr.cell(row=row_lr, column=1, value="Total Beban")
    cell.alignment = Alignment(horizontal="left", vertical="center")
    cell.border = thin_border
    cell = ws_lr.cell(row=row_lr, column=2, value=0)
    cell.number_format = '"Rp"#,##0.00'
    cell.alignment = Alignment(horizontal="right", vertical="center")
    cell.border = thin_border
    cell = ws_lr.cell(row=row_lr, column=3, value=laba_rugi_data["Total Beban"])
    cell.number_format = '"Rp"#,##0.00'
    cell.alignment = Alignment(horizontal="right", vertical="center")
    cell.border = thin_border
    row_lr += 1
    
    # Laba/Rugi
    cell = ws_lr.cell(row=row_lr, column=1, value="Laba/Rugi")
    cell.font = Font(bold=True)
    cell.alignment = Alignment(horizontal="left", vertical="center")
    cell.border = thin_border
    if laba_rugi_data["Laba/Rugi"] >= 0:
        cell = ws_lr.cell(row=row_lr, column=2, value=laba_rugi_data["Laba/Rugi"])
        cell.number_format = '"Rp"#,##0.00'
        cell.alignment = Alignment(horizontal="right", vertical="center")
        cell.border = thin_border
        cell = ws_lr.cell(row=row_lr, column=3, value=0)
        cell.number_format = '"Rp"#,##0.00'
        cell.alignment = Alignment(horizontal="right", vertical="center")
        cell.border = thin_border
    else:
        cell = ws_lr.cell(row=row_lr, column=2, value=0)
        cell.number_format = '"Rp"#,##0.00'
        cell.alignment = Alignment(horizontal="right", vertical="center")
        cell.border = thin_border
        cell = ws_lr.cell(row=row_lr, column=3, value=abs(laba_rugi_data["Laba/Rugi"]))
        cell.number_format = '"Rp"#,##0.00'
        cell.alignment = Alignment(horizontal="right", vertical="center")
        cell.border = thin_border
    
    ws_lr.column_dimensions['A'].width = 20
    ws_lr.column_dimensions['B'].width = 20
    ws_lr.column_dimensions['C'].width = 20

    wb.save(output)
    output.seek(0)
    return output.getvalue()
    
# ============================
# 6. GRAFIK
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
# 7. EXPORT EXCEL (MULTI SHEET)
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
