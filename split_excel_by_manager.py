import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


def split_excel_by_manager(path, manager_column='Manager'):
    df = pd.read_excel(path)
    for manager, group in df.groupby(manager_column):
        out_path = f"{manager}_report.xlsx"
        group.to_excel(out_path, index=False)
        wb = load_workbook(out_path)
        ws = wb.active
        for idx, col in enumerate(ws.columns, 1):
            max_len = 0
            for cell in col:
                if cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[get_column_letter(idx)].width = max_len + 2
        wb.save(out_path)


if __name__ == '__main__':
    import sys
    if len(sys.argv) < 2:
        print("Usage: python split_excel_by_manager.py <excel_file> [manager_column]")
    else:
        col = sys.argv[2] if len(sys.argv) > 2 else 'Manager'
        split_excel_by_manager(sys.argv[1], col)
