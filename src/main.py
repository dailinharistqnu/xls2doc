# from tkinter import Tk
import pandas as pd
from docx import Document
from datetime import datetime

excel_data = pd.read_excel("src/template/EmployeeList.xlsx")
# Check xem mot cot co phai dang ngay khong
def is_date_column(col):
    return pd.api.types.is_datetime64_any_dtype(col)

# Check xem mot cot co phai la so khong
def is_num_colum(col):
    return pd.api.types.is_numeric_dtype(col)

# Format ngay voi dang ddmmyy
def format_date(date):
    if pd.notnull(date):
        return date.strftime("%d/%m/%Y")
    return ""

# Ngan cach phan nghin cho so
def format_number(num):
    if pd.notnull(num):
        return f"{num:,}".replace(",",".")
    return ""

# The hien so 0 truoc so co 1 chu so
def adjust_number(num):
    if num < 10 and num > 0:
        return "0"+str(num)
    return str(num)

def process_date(date):
    parsed_date = pd.to_datetime(date)
    return {
        "Day":adjust_number(parsed_date.day),
        "Month":adjust_number(parsed_date.month),
        "Year":adjust_number(parsed_date.year)
            }
# Loop tat ca cac dong trong excel
for index,row in excel_data.iterrows():
    document = Document('src/template/Contract.docx')
    if "Date" in row.keys():
        date_data = process_date(row["Date"])
        for key, value in date_data.items():
            row[key] = value
    formatted_row = row.copy()

    for col in excel_data.columns:
        if is_date_column(excel_data[col]):
            formatted_row[col] = format_date(row[col])
        if is_num_colum(excel_data[col]):
            formatted_row[col] = format_number(row[col])

    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            # print("Key: "+key)
            for key in row.keys():
                placeholder = f'{{{key}}}'
                if placeholder in run.text:
                # print(formatted_row[key])
                    run.text = run.text.replace(placeholder,str(formatted_row[key]))

    document.save(f'HopDong_{row["No"]}_{row["Employee"]}.docx')