import openpyxl
import fileinput
wb = openpyxl.load_workbook('Sample.xlsx')
print(type(wb))
for sheetNames in wb.get_sheet_names():
    print(sheetNames)
sheet = wb.get_sheet_by_name('Sheet1')
#print(sheet['A2'].value)

row = []
row_list = []
#i represents the number of rows to be read
#j represents the number of cols to be read
for i in range(2,101):
    row = []
    for j in range(1,4):
        row.append(sheet.cell(row=i, column=j).value)
    row_list.append(row)

#Dictionary to create request
dict = {'Name': '', 'Age': 0, 'Position': ''}
attr_names  = dict.keys()
keys = []


# Retrieving keys from dictionary
for each in attr_names:
    keys.append(each)

fileIndex = 0

#Row contains row of excel sheet
for row in row_list:
    index = 0

    # Creating key,value pair of row values
    # Here attr refers to cell values
    for attr in row:
        #Replacing column 2 attribute by 0
        if(index == 1 and attr == None):
            dict[keys[index]] = 0
        else:
            dict[keys[index]] = attr
        index += 1

    fileIndex += 1
    fileName = "Request" + str(fileIndex) + ".txt"

    #Creating Data Requests
    text_file = open(fileName,"w")
    text_file.write("DataObjectName" + str(dict))
    text_file.close()

    #Replacing None with Null in text file
    with fileinput.FileInput(fileName, inplace=True) as file:
        for line in file:
            print(line.replace("None","null"),end = '')


