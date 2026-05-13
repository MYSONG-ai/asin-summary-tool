import io
import os
from datetime import datetime

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


def process_one(backup_bytes: bytes, asin_cat: pd.DataFrame) -> pd.DataFrame:
    backup = pd.read_excel(io.BytesIO(backup_bytes), engine='xlrd')
    agg = backup.groupby('Asin').agg(
        Product_Name=('Title', 'first'),
        总金额=('Net Sales', 'sum'),
        台数=('Quantity', 'sum'),
        total_rebate=('Rebate In Agreement Currency', 'sum'),
    ).reset_index()
    agg['Discount in USD'] = (agg['total_rebate'] / agg['台数']).round(2)
    agg['台数'] = agg['台数'].astype(int)
    agg['总金额'] = agg['总金额'].round(2)
    merged = agg.merge(asin_cat, left_on='Asin', right_on='ASIN', how='left')
    result = merged[['Asin', 'ID', 'Product_Name', 'Discount in USD', '总金额', '台数']].copy()
    result.columns = ['Product ASIN', 'ID', 'Product Name', 'Discount in USD', '总金额', '台数']
    return result.sort_values('台数', ascending=False).reset_index(drop=True)


def build_excel(backup_bytes_list: list, asin_bytes: bytes) -> bytes:
    asin_cat = pd.read_excel(io.BytesIO(asin_bytes), engine='xlrd')
    results  = [process_one(b, asin_cat) for b in backup_bytes_list]

    wb = Workbook()
    ws = wb.active
    ws.title = 'ASIN Summary'

    header_font = Font(name='Arial', bold=True, color='FFFFFF', size=10)
    header_fill = PatternFill('solid', start_color='1F4E79')
    white_fill  = PatternFill('solid', start_color='FFFFFF')
    grey_fill   = PatternFill('solid', start_color='F2F7FC')
    center = Alignment(horizontal='center', vertical='center')
    left   = Alignment(horizontal='left',   vertical='center')
    right  = Alignment(horizontal='right',  vertical='center')
    thin   = Side(style='thin', color='CCCCCC')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    headers    = ['Product ASIN', 'ID', 'Product Name', 'Discount in USD', '总金额', '台数']
    col_widths = [16, 10, 65, 18, 16, 10]

    for ci, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.font      = header_font
        cell.fill      = header_fill
        cell.alignment = center
        cell.border    = border
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[1].height = 22

    current_row = 2
    for block_idx, result in enumerate(results):
        for ri, row in result.iterrows():
            bg        = white_fill if ri % 2 == 0 else grey_fill
            body_font = Font(name='Arial', size=9)
            vals   = [row['Product ASIN'], row['ID'], row['Product Name'],
                      row['Discount in USD'], row['总金额'], row['台数']]
            aligns = [center, center, left, right, right, right]
            for ci, (val, aln) in enumerate(zip(vals, aligns), 1):
                cell           = ws.cell(row=current_row, column=ci, value=val)
                cell.font      = body_font
                cell.fill      = bg
                cell.alignment = aln
                cell.border    = border
            ws.cell(row=current_row, column=4).number_format = '$#,##0.00'
            ws.cell(row=current_row, column=5).number_format = '#,##0.00'
            ws.cell(row=current_row, column=6).number_format = '#,##0'
            ws.row_dimensions[current_row].height = 18
            current_row += 1
        if block_idx < len(results) - 1:
            current_row += 1  # blank row between blocks

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue(), sum(len(r) for r in results)


# ── UI ────────────────────────────────────────────────────────────────────────

DEFAULT_ASIN = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ASIN对应品类关系.xls')

st.title('ASIN Summary Generator')

backup_files = st.file_uploader('Upload BackupReport (.xls)', type=['xls'], accept_multiple_files=True)

asin_file = st.file_uploader(
    'Upload ASIN对应品类关系 (.xls)  —  optional, leave empty to use the default file',
    type=['xls'],
)
if not asin_file:
    if os.path.exists(DEFAULT_ASIN):
        st.caption('Using default ASIN对应品类关系.xls from the app folder.')
    else:
        st.warning('Default ASIN对应品类关系.xls not found. Please upload the file.')

if st.button('Generate Table', disabled=not backup_files or (not asin_file and not os.path.exists(DEFAULT_ASIN))):
    with st.spinner('Processing...'):
        try:
            asin_bytes = asin_file.read() if asin_file else open(DEFAULT_ASIN, 'rb').read()
            excel_bytes, row_count = build_excel([f.read() for f in backup_files], asin_bytes)
            ts = datetime.now().strftime('%Y%m%d_%H%M%S')
            st.success(f'Done — {row_count} ASINs')
            st.download_button(
                label='Download Excel',
                data=excel_bytes,
                file_name=f'asin_summary_table_{ts}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            )
        except Exception as e:
            st.error(f'Error: {e}')
