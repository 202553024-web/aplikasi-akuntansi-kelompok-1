import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def export_excel_multi(df_jurnal, buku_besar, neraca):
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Jurnal"

    # ============================
    # 1. SHEET JURNAL
    # ============================
    headers = df_jurnal.columns.tolist()
    ws1.append(headers)

    for row in df_jurnal.itertuples(index=False):
        ws1.append(list(row))

    # Style header
    for col in range(1, len(headers) + 1):
        cell = ws1.cell(row=1, column=col)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    # Autofit kolom
    for col in ws1.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            val = str(cell.value) if cell.value else ""
            max_length = max(max_length, len(val))
        ws1.column_dimensions[col_letter].width = max_length + 2

    # Border
    thin = Side(border_style="thin", color="000000")
    for row in ws1.iter_rows():
        for cell in row:
            cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)

    # ============================
    # 2. SHEET BUKU BESAR PER AKUN
    # ============================
    for akun, df_akun in buku_besar.items():
        ws = wb.create_sheet(title=f"Buku {akun[:25]}")

        ws.append(df_akun.columns.tolist())
        for row in df_akun.itertuples(index=False):
            ws.append(list(row))

        # Format header
        for col in range(1, len(df_akun.columns) + 1):
            cell = ws.cell(row=1, column=col)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center")

        # Autofit
        for col in ws.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                val = str(cell.value) if cell.value else ""
                max_length = max(max_length, len(val))
            ws.column_dimensions[col_letter].width = max_length + 2

        # Border
        for row in ws.iter_rows():
            for cell in row:
                cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)

    # ============================
    # 3. SHEET NERACA
    # ============================
    ws3 = wb.create_sheet("Neraca")

    ws3.append(neraca.columns.tolist())
    for row in neraca.itertuples(index=False):
        ws3.append(list(row))

    # Format header
    for col in range(1, len(neraca.columns) + 1):
        cell = ws3.cell(row=1, column=col)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    # Autofit
    for col in ws3.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            val = str(cell.value) if cell.value else ""
            max_length = max(max_length, len(val))
        ws3.column_dimensions[col_letter].width = max_length + 2

    # Border
    for row in ws3.iter_rows():
        for cell in row:
            cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)

    # ============================
    # 4. SAVE KE BUFFER
    # ============================
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer
