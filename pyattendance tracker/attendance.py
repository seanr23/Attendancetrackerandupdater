import openpyxl

def load_workbook(file_path):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    return workbook, sheet

def add_attendance(sheet, name, date, status):
    sheet.append([name, date, status])

def save_workbook(workbook, file_path):
    workbook.save(file_path)
