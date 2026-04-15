import openpyxl as xl
from openpyxl.chart import BarChart, Reference

def process_workbook(filename):
    wb = xl.load_workbook(filename)
    # wb = xl.load_workbook('transactions.xlsx')
    sheet = wb['Foaie1']
    # cell = sheet.cell(row=1, column=1)
    # cell = sheet['a1']
    # print(cell.value)
    # print(sheet.max_row)

    for row in range(2, sheet.max_row + 1):
        cell = sheet.cell(row,3)
        # print(cell.value)
        value = cell.value
        corrected_price = float(value) * 0.9
        # print(corrected_price)
        corrected_price_cell = sheet.cell(row,4)
        corrected_price_cell.value = corrected_price

    values = Reference(sheet, min_row = 2 , max_row = sheet.max_row + 1,
              min_col = 4,
              max_col = 4)

    chart = BarChart()
    chart.add_data(values)
    sheet.add_chart(chart,'e2')
    # wb.save('transactions2.xlsx')
    wb.save(filename)

