from pandas import *
import openpyxl

fileName = input("Enter a filename in the same folder: ")
fileName2 = input("Enter a filename in the same folder: ")
sheetName = input("Enter a sheet name, leave blank if there is only one sheet : ")
columnLetter = input("Input the column letter to compare, if multiple put a comma after each e.g. A,B,C: ").upper()
print()

columnLetter = columnLetter.split(",")

if sheetName != "":
    targetFile = read_excel(f'{fileName}.xlsx', sheet_name=sheetName)
    targetFile2 = read_excel(f'{fileName2}.xlsx', sheet_name=sheetName)

else:
    targetFile = read_excel(f'{fileName}.xlsx')
    targetFile2 = read_excel(f'{fileName2}.xlsx')

targetFile.sort_values(ascending=True, by=targetFile.columns[0])
targetFile2.sort_values(ascending=True, by=targetFile2.columns[0])

if targetFile.equals(targetFile2):
    print("Complete file is the same")
else:
    print("Complete file is not the same")

for i in range(0, len(columnLetter)):
    if sheetName != "":
        targetFile = read_excel(f'{fileName}.xlsx', sheet_name=sheetName, usecols=columnLetter[i])
        targetFile2 = read_excel(f'{fileName2}.xlsx', sheet_name=sheetName, usecols=columnLetter[i])

    else:
        targetFile = read_excel(f'{fileName}.xlsx', usecols=columnLetter[i])
        targetFile2 = read_excel(f'{fileName2}.xlsx', usecols=columnLetter[i])

    targetFile.sort_values(ascending=True, by=targetFile.columns[0])
    targetFile2.sort_values(ascending=True, by=targetFile2.columns[0])

    if targetFile.equals(targetFile2):
        print(f"Column {columnLetter[i]} is the same")
    else:
        print(f"Column {columnLetter[i]} is not the same")



