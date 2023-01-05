import openpyxl

book = openpyxl.load_workbook('Time_Table_2.xlsx')

sheet = book.get_sheet_by_name("Table 7")  #Division B

a1 = sheet['B15']
a2 = sheet['B16']
a3 = sheet['C15']
a4 = sheet['C16']
t1 = sheet['B2']
print(t1.value+"\n")
print(a1.value , end=" "+a2.value+"\n")
print(a3.value , end=" "+a4.value)