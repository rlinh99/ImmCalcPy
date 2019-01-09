from openpyxl import load_workbook
from .classes import *

def read_file(name):
    wb = load_workbook(name, data_only=True)

    sheet = wb.active
    data = sheet.iter_rows()
    return data


def insert_data(rows):
    data = []
    for row in rows:
        print(type(row[0].value))

        # data.append(Record(row[0].toStromg))
    return 0


def calculate_date(data):
    return 0
