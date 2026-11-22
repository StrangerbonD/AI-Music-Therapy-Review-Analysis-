import os
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

dest_dir = r"E:\Insight\Sample_Dataset"
os.makedirs(dest_dir, exist_ok=True)

source_files = {
    1: r"E:\Insight\phase_2_sampling\research_sample_5000.csv",
    2: r"E:\Insight\samplings\sampling_2\research_sample_5000.csv",
    3: r"E:\Insight\samplings\sampling_3\research_sample_5000.csv",
    4: r"E:\Insight\samplings\sampling_4\research_sample_5000.csv",
    5: r"E:\Insight\samplings\sampling_5\research_sample_5000.csv"
}

header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
thin_border = Side(border_style="thin", color="D9D9D9")
border_all = Border(left=thin_border, right=thin_border, top=thin_border, bottom=thin_border)

for s_num, src_csv in source_files.items():
    df = pd.read_csv(src_csv)
    
    # Save CSV
    csv_out = os.path.join(dest_dir, f"sample_{s_num}_5000.csv")
    df.to_csv(csv_out, index=False, encoding="utf-8")
    
    # Save Excel (.xlsx)
    xlsx_out = os.path.join(dest_dir, f"sample_{s_num}_5000.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Sample_{s_num}_5000"
    ws.views.sheetView[0].showGridLines = True
    
    ws.append(list(df.columns))
    for col_idx in range(1, len(df.columns) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_all
    ws.row_dimensions[1].height = 24
    
    for row in df.itertuples(index=False):
        ws.append(list(row))
        
    for col in ws.iter_cols(min_row=1, max_row=min(len(df)+1, 100)):
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(min(max_len + 3, 50), 12)
        
    ws.freeze_panes = "A2"
    wb.save(xlsx_out)

print("Export completed successfully.")
