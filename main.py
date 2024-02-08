import pandas as pd
import openpyxl
import math

wb = openpyxl.load_workbook('test.xlsx')            # Loads spreadsheet
ws = wb['Output']
count = 0                                           # Counts each written row

def write_ans(x, y, i):
    ws.cell(row = i, column = 1, value = x)
    ws.cell(row = i, column = 2, value = y)

def write_err(i):
    ws.cell(row = i, column = 1, value = 'This equation has complex roots')

def calculate_roots(a, b, c, count):
    if a == 0:                                     # Verify quadratic equation
        print("Not a quadratic equation. a cannot be zero")
    else:   
        determinant = (b)**2 - 4 * a * c           # Calculate determinant

        if determinant >= 0:                       # Calculates only real roots
            x_one = (-b + math.sqrt(determinant)) / (2 * a)
            x_two = (-b - math.sqrt(determinant)) / (2 * a)      
            write_ans(x_one, x_two, count)
        else:                                      # Complex roots error handling
            write_err(count)

def create_sheet():
    wb.create_sheet('Output')
    ws['A2'] = "x1"
    ws['A2'] = "x2"

def delete_sheet():
    wb.remove(wb['Output'])

def read():
    delete_sheet()                                  # Resets previous calculations
    create_sheet()
    df = pd.read_excel('test.xlsx', sheet_name = 'Sheet1')     # Reads inputs
    for i, row in df.iterrows():                    # Fetches values from each cell
        l = row.to_list()
        a, b, c = l[0], l[1], l[2]
        calculate_roots(a, b, c, i)                    # Calculates for each row
    wb.save('test.xlsx')
read()
