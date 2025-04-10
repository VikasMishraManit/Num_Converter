from openpyxl import load_workbook
import re
from dateutil.parser import parse

def is_date_string(value):
    try:
        parse(value, fuzzy=False)
        return True
    except:
        return False

def is_pure_number(value):
    return re.fullmatch(r"-?\d+(\.\d+)?", value.strip()) is not None

def convert_text_to_number_safely(file_path, sheet_name=None):
    wb = load_workbook(file_path)
    sheets = [wb[sheet_name]] if sheet_name else wb.worksheets

    for ws in sheets:
        for row in ws.iter_rows():
            for cell in row:
                val = cell.value
                if isinstance(val, str):
                    if is_pure_number(val) and not is_date_string(val):
                        # Safe to convert
                        num = float(val)
                        cell.value = int(num) if num.is_integer() else num

    wb.save(file_path)
    print("Safe numeric conversion done.")

# Example usage
convert_text_to_number_safely("your_excel_file.xlsx")
