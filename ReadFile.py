import openpyxl
wb = openpyxl.load_workbook('Sample.xlsx')
print(type(wb))
for sheetNames in wb.get_sheet_names():
    print(sheetNames)
sheet = wb.get_sheet_by_name('Sheet1')
#print(sheet['A2'].value)

row = []
row_list = []
for i in range(2,5):
    row = []
    for j in range(1,4):
        row.append(sheet.cell(row=i, column=j).value)
    row_list.append(row)

        #print(sheet.cell(row=i, column=j).value)
dict = {'Name' : '', 'Age' : 0, 'Position' : ''}

attr_names  = dict.keys()
keys = []
for each in attr_names:
    keys.append(each)

for row in row_list:
    index = 0
    for attr in row:
        dict[keys[index]] = attr
        index += 1
    print(dict)