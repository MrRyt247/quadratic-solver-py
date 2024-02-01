import pandas as pd
import openpyxl 
import math

df = pd.read_excel('input.xlsx')

wb = openpyxl.load_workbook('input.xlsx')
ws_output = wb['Output']
count = 0

def write_ans(x, y, i):
    ws_output.cell(row=i, column=1, value = x)
    ws_output.cell(row=i, column=2, value = y)

def write_err(i):
    ws_output.cell(row=i, column=1, value = 'This equation has complex roots')

def calculate_roots(a, b, c):
    if a == 0:                              # Verify quadratic equation
        print("Not a quadratic equation. a cannot be zero")
        return
    
    determinant = (b)**2 - 4 * a * c           # Calculate determinant\
    count += 1
    
    if determinant >= 0:                    # Real roots calcaltion
        x_one = (-b + math.sqrt(determinant)) / 2 * a
        x_two = (-b - math.sqrt(determinant)) / 2 * a
        write_ans(x_one, x_two, count)
    else:                                   # Complex roots error handling
        write_err(count)

def read():
    delete_sheet()
    create_sheet()
    for i, row in df.iterrows():
        l = row.to_list()
        a, b, c = l[0], l[1], l[2]
        calculate_roots(a, b, c)
    count = 0

read()

def create_sheet():
    wb.create_sheet('Output')
    wb.save('input.xlsx')


def delete_sheet():
    wb.remove(wb['Output'])
    wb.save('input.xlsx')

    
