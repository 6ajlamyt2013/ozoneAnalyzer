import re

def excel_cell_to_int(cell_value):
    if cell_value is None:
        return 0
    if not isinstance(cell_value, str):
        cell_value = str(cell_value)
    cell_value = re.sub(r'[^\d.,]', '', cell_value)
    cell_value = cell_value.replace(" ", "")
    if not cell_value or cell_value == '.':
        return 0
    if ',' in cell_value or '.' in cell_value:
        cell_value = cell_value.replace(',', '.')
        return int(float(cell_value))
    else:
        return int(cell_value)

def process_excel_cell(cell_value):
    if not cell_value:
        return ""
    if isinstance(cell_value, str) and cell_value.strip() == "":
        return ""
    if isinstance(cell_value, (int, float)):
        return str(cell_value).lower()
    return str(cell_value).lower()