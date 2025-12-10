def export_excel_multi(df):
    output = io.BytesIO()
    wb = Workbook()

    # =============================
    #   SETUP DATA BULAN & TAHUN
    # =============================
    df["Tanggal"] = pd.to_datetime(df["Tanggal"])
    df["Bulan"] = df["Tanggal"].dt.month
    df["Tahun"] = df["Tanggal"].dt.year
    df_sorted = df.sort_values("Tanggal")

    # Ambil periode pertama dan terakhir
    bulan_awal = calendar.month_name[int(df_sorted["Bulan"].iloc[0])]
    tahun_awal = df_sorted["Tahun"].iloc[0]
    bulan_akhir = calendar.month_name[int(df_sorted["Bulan"].iloc[-1])]
    tahun_akhir = df_sorted["Tahun"].iloc[-1]

    periode_text = f"Periode: {bulan_awal} {tahun_awal} - {bulan_akhir} {tahun_akhir}"

    # =====================================================
    #  SHEET 1: JURNAL UMUM (DENGAN HEADER PERIODE)
    # =====================================================
    ws1 = wb.active
    ws1.title = "Jurnal Umum"

    # Header periode
    ws1.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
    header_cell = ws1.cell(row=1, column=1, value=periode_text)
    header_cell.font = Font(bold=True, size=13)
    header_cell.alignment = Alignment(horizontal="center")

    current_row = 3  # mulai baris 3 setelah header periode

    for (tahun, bulan), grup in df_sorted.groupby(["Tahun", "Bulan"]):
        nama_bulan = calendar.month_name[bulan].upper()

        # Judul bulan per grup
        ws1.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=6)
        cell = ws1.cell(row=current_row, column=1, value=f"=== {nama_bulan} {tahun} ===")
        cell.font = Font(bold=True, size=12)
        current_row += 1

        # Header tabel
        headers = list(grup.columns.drop(["Bulan", "Tahun"]))
        for col_num, header in enumerate(headers, start=1):
            ws1.cell(row=current_row, column=col_num, value=header).font = Font(bold=True)
        current_row += 1

        # Isi tabel
        for r in dataframe_to_rows(grup.drop(["Bulan", "Tahun"], axis=1), index=False, header=False):
            for c_idx, val in enumerate(r, start=1):
                cell = ws1.cell(row=current_row, column=c_idx, value=val)
                if c_idx in [4, 5] and isinstance(val, (int, float)):
                    cell.number_format = '"Rp"#,##0'
            current_row += 1

        current_row += 2  # spasi antar bulan

    # =====================================================
    #  SHEET 2: BUKU BESAR (DENGAN HEADER PERIODE)
    # =====================================================
    ws2 = wb.create_sheet("Buku Besar")

    ws2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
    h2 = ws2.cell(row=1, column=1, value=periode_text)
    h2.font = Font(bold=True, size=13)
    h2.alignment = Alignment(horizontal="center")

    row_buku = 3

    buku = buku_besar(df)
    for akun, data in buku.items():
        ws2.merge_cells(start_row=row_buku, start_column=1, end_row=row_buku, end_column=6)
        ws2.cell(row=row_buku, column=1, value=f"== {akun.upper()} ==").font = Font(bold=True, size=12)
        row_buku += 1

        for col_num, header in enumerate(data.columns, start=1):
            ws2.cell(row=row_buku, column=col_num, value=header).font = Font(bold=True)
        row_buku += 1

        for r in dataframe_to_rows(data, index=False, header=False):
            for c_idx, val in enumerate(r, start=1):
                cell = ws2.cell(row=row_buku, column=c_idx, value=val)
                if c_idx in [4, 5, 6] and isinstance(val, (int, float)):
                    cell.number_format = '"Rp"#,##0'
            row_buku += 1

        row_buku += 2

    # =====================================================
    #  SHEET 3: NERACA SALDO (DENGAN HEADER PERIODE)
    # =====================================================
    ws3 = wb.create_sheet("Neraca Saldo")

    ws3.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
    h3 = ws3.cell(row=1, column=1, value=periode_text)
    h3.font = Font(bold=True, size=13)
    h3.alignment = Alignment(horizontal="center")

    neraca = neraca_saldo(df).reset_index()

    row = 3
    for r in dataframe_to_rows(neraca, index=False, header=True):
        for c_idx, val in enumerate(r, start=1):
            cell = ws3.cell(row=row, column=c_idx, value=val)
            if row == 3:
                cell.font = Font(bold=True)
            elif c_idx in [2, 3, 4] and isinstance(val, (int, float)):
                cell.number_format = '"Rp"#,##0'
        row += 1

    wb.save(output)
    output.seek(0)
    return output.getvalue()
