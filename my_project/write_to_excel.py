import openpyxl

book=openpyxl.Workbook()

sheet = book.active

sheet['A2'] = 100
sheet['B3'] = 'hello'


book.save("my_book.xlsx")
book.close()
