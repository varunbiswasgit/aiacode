"""
test_split_excel_by_manager.py — pytest test suite for split_excel_by_manager.py
Run from the repo root:
    pytest tests/test_split_excel_by_manager.py -v
"""
import sys
import pytest
import pandas as pd
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / 'scripts'))
from split_excel_by_manager import sanitize_filename, split_excel_by_manager


# ── Helpers ───────────────────────────────────────────────────────────────────

def make_xlsx(path, data, columns):
    """Write a simple DataFrame to .xlsx for use in tests."""
    df = pd.DataFrame(data, columns=columns)
    df.to_excel(path, index=False)
    return path


# ── TC-PY-01: sanitize_filename — invalid characters replaced ─────────────────

def test_sanitize_invalid_chars():
    assert sanitize_filename('Alice/Bob:Report') == 'Alice_Bob_Report'


def test_sanitize_all_invalid_chars():
    result = sanitize_filename('<>:"/\\|?*')
    assert '_' in result and '/' not in result


# ── TC-PY-02: sanitize_filename — Windows reserved names ─────────────────────

@pytest.mark.parametrize('name', [
    'CON', 'PRN', 'AUX', 'NUL',
    'COM1', 'COM5', 'COM9',
    'LPT1', 'LPT5', 'LPT9',
])
def test_sanitize_reserved_names(name):
    assert sanitize_filename(name) == 'Unknown_Manager'
    assert sanitize_filename(name.lower()) == 'Unknown_Manager'


# ── TC-PY-03: sanitize_filename — 200-character length cap ───────────────────

def test_sanitize_length_cap():
    result = sanitize_filename('A' * 300)
    assert len(result) <= 200


# ── TC-PY-04: sanitize_filename — normal name unchanged ──────────────────────

def test_sanitize_normal_name():
    assert sanitize_filename('John Smith') == 'John Smith'


# ── TC-PY-05: missing input file ─────────────────────────────────────────────

def test_missing_input_file(tmp_path):
    assert split_excel_by_manager(str(tmp_path / 'no_such_file.xlsx')) is False


# ── TC-PY-06: unsupported file type ──────────────────────────────────────────

def test_unsupported_file_type(tmp_path):
    bad_file = tmp_path / 'data.csv'
    bad_file.write_text('Manager\nAlice\n')
    assert split_excel_by_manager(str(bad_file)) is False


# ── TC-PY-07: manager column not found ───────────────────────────────────────

def test_manager_column_not_found(tmp_path):
    xlsx = make_xlsx(tmp_path / 'input.xlsx', [['Alice', 100]], ['Supervisor', 'Score'])
    assert split_excel_by_manager(str(xlsx), manager_column='Manager') is False


# ── TC-PY-08: empty DataFrame ────────────────────────────────────────────────

def test_empty_dataframe(tmp_path):
    df = pd.DataFrame(columns=['Manager', 'Score'])
    xlsx = tmp_path / 'empty.xlsx'
    df.to_excel(xlsx, index=False)
    assert split_excel_by_manager(str(xlsx), output_dir=str(tmp_path / 'out')) is False


# ── TC-PY-09: standard split — one file per manager ──────────────────────────

def test_standard_split(tmp_path):
    xlsx = make_xlsx(tmp_path / 'data.xlsx',
                     [['Alice', 90], ['Bob', 85], ['Alice', 92]], ['Manager', 'Score'])
    out_dir = str(tmp_path / 'reports')
    assert split_excel_by_manager(str(xlsx), output_dir=out_dir) is True
    names = {f.name for f in Path(out_dir).glob('*.xlsx')}
    assert 'Alice_report.xlsx' in names
    assert 'Bob_report.xlsx' in names


# ── TC-PY-10: correct rows per manager ───────────────────────────────────────

def test_correct_rows_per_manager(tmp_path):
    xlsx = make_xlsx(tmp_path / 'data.xlsx',
                     [['Alice', 90], ['Bob', 85], ['Alice', 92], ['Alice', 88]],
                     ['Manager', 'Score'])
    out_dir = tmp_path / 'reports'
    split_excel_by_manager(str(xlsx), output_dir=str(out_dir))
    assert len(pd.read_excel(out_dir / 'Alice_report.xlsx')) == 3
    assert len(pd.read_excel(out_dir / 'Bob_report.xlsx'))   == 1


# ── TC-PY-11: blank manager rows skipped ─────────────────────────────────────

def test_blank_manager_skipped(tmp_path):
    xlsx = make_xlsx(tmp_path / 'data.xlsx',
                     [['Alice', 90], [None, 85], ['', 88]], ['Manager', 'Score'])
    out_dir = tmp_path / 'reports'
    split_excel_by_manager(str(xlsx), output_dir=str(out_dir))
    files = list(out_dir.glob('*.xlsx'))
    assert len(files) == 1
    assert files[0].name == 'Alice_report.xlsx'


# ── TC-PY-12: configurable manager column name ───────────────────────────────

def test_configurable_column(tmp_path):
    xlsx = make_xlsx(tmp_path / 'data.xlsx',
                     [['Alice', 90], ['Bob', 85]], ['Supervisor', 'Score'])
    out_dir = tmp_path / 'reports'
    assert split_excel_by_manager(str(xlsx), manager_column='Supervisor',
                                  output_dir=str(out_dir)) is True
    assert 'Alice_report.xlsx' in {f.name for f in out_dir.glob('*.xlsx')}


# ── TC-PY-13: invalid characters in manager name sanitized ───────────────────

def test_manager_name_sanitized(tmp_path):
    xlsx = make_xlsx(tmp_path / 'data.xlsx', [['Alice/Bob', 90]], ['Manager', 'Score'])
    out_dir = tmp_path / 'reports'
    split_excel_by_manager(str(xlsx), output_dir=str(out_dir))
    files = list((tmp_path / 'reports').glob('*.xlsx'))
    assert files[0].name == 'Alice_Bob_report.xlsx'


# ── TC-PY-14: output directory created when it does not exist ────────────────

def test_output_dir_created(tmp_path):
    xlsx = make_xlsx(tmp_path / 'data.xlsx', [['Alice', 90]], ['Manager', 'Score'])
    out_dir = tmp_path / 'new' / 'nested' / 'reports'
    assert not out_dir.exists()
    split_excel_by_manager(str(xlsx), output_dir=str(out_dir))
    assert out_dir.exists()


# ── TC-PY-15: column width clamped between 8 and 50 ─────────────────────────

def test_column_width_clamped(tmp_path):
    from openpyxl import load_workbook
    xlsx = make_xlsx(tmp_path / 'data.xlsx',
                     [['Alice', 'X' * 200]], ['Manager', 'Notes'])
    out_dir = tmp_path / 'reports'
    split_excel_by_manager(str(xlsx), output_dir=str(out_dir))
    wb = load_workbook(out_dir / 'Alice_report.xlsx')
    for col_letter, dim in wb.active.column_dimensions.items():
        assert dim.width <= 50, f'Column {col_letter} width {dim.width} exceeds 50'
        assert dim.width >= 8,  f'Column {col_letter} width {dim.width} below 8'
