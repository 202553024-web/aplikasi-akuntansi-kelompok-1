import streamlit as st
import pandas as pd
import altair as alt
from datetime import datetime
import calendar
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

# ============================
# CONFIG
# ============================
st.set_page_config(
    page_title="Aplikasi Akuntansi",
    page_icon="💰",
    layout="wide"
)

st.markdown("""
<style>
.title { font-size:36px; font-weight:800; text-align:center; color:#1a237e; }
.subtitle { font-size:22px; font-weight:600; color:#1a237e; }
</style>
""", unsafe_allow_html=True)

st.markdown("<div class='title'>📊 Aplikasi Akuntansi</div>", unsafe_allow_html=True)

# ============================
# SESSION
# ============================
if "transaksi" not in st.session_state:
    st.session_state.transaksi = []

# ============================
# FORMAT RUPIAH
# ============================
def to_rp(x):
    return f"Rp {x:,.0f}".replace(",", ".")

# ============================
# FUNGSI DASAR
# ============================
def tambah_transaksi(tgl, akun, ket, debit, kredit):
    st.session_state.transaksi.append({
        "Tanggal": tgl,
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

# ============================
# EXPORT EXCEL LENGKAP
# ============================
def export_excel(df):
    output = io.BytesIO()
    wb = Workbook()

    thin = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    header = PatternFill("solid", fgColor="4472C4")
    title = PatternFill("solid", fgColor="B4C7E7")

    # ============================
    # SHEET LAPORAN KEUANGAN
    # ============================
    ws = wb.active
    ws.title = "Laporan Keuangan"

    df["Tanggal"] = pd.to_datetime(df["Tanggal"])
    df["Tahun"] = df["Tanggal"].dt.year
    df["Bulan"] = df["Tanggal"].dt.month

    row = 1
    for (thn, bln), g in df.groupby(["Tahun", "Bulan"]):
        ws.merge_cells(row,1,row,5)
        c = ws.cell(row,1,f"Laporan Keuangan Tahun {thn}")
        c.font = Font(bold=True,size=14)
        c.alignment = Alignment(horizontal="center")
        c.fill = title
        row+=1

        ws.merge_cells(row,1,row,5)
        c = ws.cell(row,1,f"Bulan {calendar.month_name[bln]}")
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center")
        c.fill = title
        row+=1

        for i,h in enumerate(["Tanggal","Akun","Keterangan","Debit","Kredit"],1):
            cell=ws.cell(row,i,h)
            cell.font=Font(bold=True,color="FFFFFF")
            cell.fill=header
            cell.border=thin
        row+=1

        for _,r in g.iterrows():
            ws.cell(row,1,r["Tanggal"])
            ws.cell(row,2,r["Akun"])
            ws.cell(row,3,r["Keterangan"])
            ws.cell(row,4,r["Debit"]).number_format='"Rp"#,##0.00'
            ws.cell(row,5,r["Kredit"]).number_format='"Rp"#,##0.00'
            row+=1
        row+=2

    # ============================
    # SHEET LABA RUGI
    # ============================
    ws_lr = wb.create_sheet("Laba Rugi")
    ws_lr.merge_cells(1,1,1,3)
    t = ws_lr.cell(1,1,"Laporan Laba Rugi")
    t.font = Font(bold=True,size=14)
    t.alignment = Alignment(horizontal="center")
    t.fill = title

    r=3
    for (thn, bln), g in df.groupby(["Tahun","Bulan"]):
        ws_lr.merge_cells(r,1,r,3)
        ws_lr.cell(r,1,f"{calendar.month_name[bln]} {thn}").font=Font(bold=True)
        r+=1

        pendapatan = g[g["Akun"].str.contains("Pendapatan")]["Kredit"].sum()
        beban = g[g["Akun"].str.contains("Beban")]["Debit"].sum()
        laba = pendapatan - beban

        data=[("Pendapatan",pendapatan),("Total Beban",beban),("Laba Bersih",laba)]
        for k,v in data:
            ws_lr.cell(r,1,k)
            ws_lr.cell(r,2,v).number_format='"Rp"#,##0.00'
            r+=1
        r+=1

    wb.save(output)
    output.seek(0)
    return output

# ============================
# MENU
# ============================
menu = st.sidebar.radio(
    "📌 MENU",
    ["Input Transaksi","Jurnal Umum","Buku Besar","Neraca Saldo","Laporan Laba Rugi","Grafik","Export Excel"]
)

# ============================
# INPUT
# ============================
if menu=="Input Transaksi":
    tgl = st.date_input("Tanggal",datetime.now())
    akun = st.selectbox("Akun",[
        "Kas","Modal","Pendapatan Jasa",
        "Beban Listrik","Beban Gaji","Beban Sewa"
    ])
    ket = st.text_input("Keterangan")
    d = st.number_input("Debit",0)
    k = st.number_input("Kredit",0)

    if st.button("Tambah"):
        tambah_transaksi(str(tgl),akun,ket,d,k)
        st.success("Transaksi ditambahkan")

    if st.session_state.transaksi:
        df=pd.DataFrame(st.session_state.transaksi)
        df["Debit"]=df["Debit"].apply(to_rp)
        df["Kredit"]=df["Kredit"].apply(to_rp)
        st.dataframe(df,use_container_width=True)

# ============================
# JURNAL UMUM
# ============================
elif menu=="Jurnal Umum":
    if st.session_state.transaksi:
        df=pd.DataFrame(st.session_state.transaksi)
        st.dataframe(df,use_container_width=True)

# ============================
# BUKU BESAR
# ============================
elif menu=="Buku Besar":
    if st.session_state.transaksi:
        df=pd.DataFrame(st.session_state.transaksi)
        bb=buku_besar(df)
        for a,d in bb.items():
            st.subheader(a)
            st.dataframe(d,use_container_width=True)

# ============================
# NERACA SALDO
# ============================
elif menu=="Neraca Saldo":
    if st.session_state.transaksi:
        df=pd.DataFrame(st.session_state.transaksi)
        n=df.groupby("Akun")[["Debit","Kredit"]].sum()
        n["Saldo"]=n["Debit"]-n["Kredit"]
        n=n.applymap(to_rp)
        st.dataframe(n,use_container_width=True)

# ============================
# LABA RUGI
# ============================
elif menu=="Laporan Laba Rugi":
    df=pd.DataFrame(st.session_state.transaksi)
    if not df.empty:
        pendapatan=df[df["Akun"].str.contains("Pendapatan")]["Kredit"].sum()
        beban=df[df["Akun"].str.contains("Beban")]["Debit"].sum()
        laba=pendapatan-beban
        st.table(pd.DataFrame({
            "Keterangan":["Pendapatan","Total Beban","Laba Bersih"],
            "Jumlah":[to_rp(pendapatan),to_rp(beban),to_rp(laba)]
        }))

# ============================
# GRAFIK
# ============================
elif menu=="Grafik":
    if st.session_state.transaksi:
        df=pd.DataFrame(st.session_state.transaksi)
        st.altair_chart(
            alt.Chart(df).mark_bar().encode(
                x="Akun",y="Debit"
            ),
            use_container_width=True
        )

# ============================
# EXPORT
# ============================
elif menu=="Export Excel":
    if st.session_state.transaksi:
        df=pd.DataFrame(st.session_state.transaksi)
        file=export_excel(df)
        st.download_button(
            "📥 Download Excel",
            file,
            "laporan_akuntansi_lengkap.xlsx"
        )
