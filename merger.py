
from io import BytesIO
from pathlib import Path
import pandas as pd

SUPPORTED = {'.csv', '.xlsx'}


def read_table(file, filename):
    ext = Path(filename).suffix.lower()
    if ext not in SUPPORTED:
        raise ValueError(f'Unsupported file type: {filename}')
    if ext == '.csv':
        try:
            return pd.read_csv(file, encoding='utf-8-sig')
        except UnicodeDecodeError:
            file.seek(0)
            return pd.read_csv(file, encoding='cp1256')
    return pd.read_excel(file, engine='openpyxl')


def merge_files(files):
    """Merge file-like uploads; align columns by name, preserving origin."""
    if not files:
        raise ValueError('Upload at least one file.')
    frames = []
    for upload in files:
        name = upload.name
        if Path(name).suffix.lower() not in SUPPORTED:
            raise ValueError(f'Unsupported file type: {name}')
        upload.seek(0)
        try:
            df = read_table(upload, name)
        except Exception as exc:
            raise ValueError(f'Could not read {name}: {exc}') from exc
        if df.empty:
            raise ValueError(f'{name} contains no data rows.')
        if 'Source File' in df.columns:
            df = df.rename(columns={'Source File': 'Original Source File'})
        df.insert(0, 'Source File', name)
        frames.append(df)
    return pd.concat(frames, ignore_index=True, sort=False)


def to_excel_bytes(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Merged Data')
        sheet = writer.sheets['Merged Data']
        sheet.freeze_panes = 'A2'
        sheet.auto_filter.ref = sheet.dimensions
        from openpyxl.styles import Font, PatternFill
        from openpyxl.utils import get_column_letter
        for cell in sheet[1]:
            cell.font = Font(color='FFFFFF', bold=True)
            cell.fill = PatternFill('solid', fgColor='17365D')
        for idx, col in enumerate(df.columns, start=1):
            values = df[col].dropna().astype(str).head(500)
            width = max([len(str(col)), *[len(v) for v in values]]) + 3
            sheet.column_dimensions[get_column_letter(idx)].width = min(max(width, 14), 45)
    return output.getvalue()
