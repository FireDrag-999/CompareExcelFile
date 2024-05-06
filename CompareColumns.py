from pandas import read_excel, ExcelFile
from logging import basicConfig, log, INFO
import logging

basicConfig(level=logging.INFO, filename="log", filemode="w", format="%(message)s")

# Problems: different headers cause errors

def checkSheet(sheet):
    listOfColumns = listOfColumnsPerSheet[listOfSheets.index(sheet)][1]
    targetFile = read_excel(f'{fileName}.xlsx', sheet_name=sheet)
    targetFile2 = read_excel(f'{fileName2}.xlsx', sheet_name=sheet)
    if not targetFile.empty or not targetFile2.empty:
        targetFile.sort_values(ascending=True, by=targetFile.columns[0])  # sorts by first column or only column
        targetFile2.sort_values(ascending=True, by=targetFile2.columns[0])
        if str(targetFile) == str(targetFile2):
            print(f"Sheet: {sheet} is the same"), log(level=INFO, msg=f"Sheet: {sheet} is the same")
        else:
            print(f"Sheet: {sheet} is not the same"), log(level=INFO, msg=f"Sheet: {sheet} is not the same")

        for colName in listOfColumns:
            checkColumn(sheet, colName)
    else:
        print(f'Sheet: {sheet} has no data'), log(level=INFO, msg=f'Sheet: {sheet} has no data')
    print(), log(level=INFO, msg="")


def checkColumn(sheet, colName):
    targetFile = read_excel(f'{fileName}.xlsx', usecols=[listOfColumns.index(colName)], sheet_name=sheet)
    targetFile2 = read_excel(f'{fileName2}.xlsx', usecols=[listOfColumns.index(colName)], sheet_name=sheet)

    if not targetFile.empty or not targetFile2.empty:
        targetFile.sort_values(ascending=True, by=targetFile.columns[0])
        targetFile2.sort_values(ascending=True, by=targetFile2.columns[0])

        if str(targetFile) == str(targetFile2):
            print(f"Sheet: {sheet}, Column: {colName} is the same"), log(level=INFO, msg=f"Sheet: {sheet}, Column: {colName} is the same")
        else:
            print(f"Sheet: {sheet}, Column: {colName} is not the same"), log(level=INFO, msg=f"Sheet: {sheet}, Column: {colName} is not the same")
            print(f"First file {sheet}, {colName}: has the length {len(targetFile)}"), log(level=INFO, msg=f"First file {sheet}, {colName}: has the length {len(targetFile)}")
            print(f"Second file {sheet}, {colName}: has the length {len(targetFile2)}"), log(level=INFO, msg=f"Second file {sheet}, {colName}: has the length {len(targetFile2)}")
            print(), log(level=INFO, msg="")
            if not notMatchingColumn.__contains__((sheet, colName)):
                notMatchingColumn.append((sheet, colName))
    else:
        print(f'Sheet: {sheet}, Column {colName} has no data'), log(level=INFO, msg=f'Sheet: {sheet}, Column {colName} has no data')


def checkAllRows(sheet, colName):
    counter = 0
    targetFile = read_excel(f'{fileName}.xlsx', usecols=[listOfColumns.index(colName)], sheet_name=sheet)
    targetFile2 = read_excel(f'{fileName2}.xlsx', usecols=[listOfColumns.index(colName)], sheet_name=sheet)

    for rowNum in range(0, len(targetFile)):
        if str(targetFile.values[rowNum]) != str(targetFile2.values[rowNum]):
            print(f"Sheet: {sheet}, Column: {colName}, row {rowNum + 2} is different: {targetFile.values[rowNum]} and {targetFile2.values[rowNum]}"), log(level=INFO, msg=f"Sheet: {sheet}, Column: {colName}, row {rowNum + 2} is different: {targetFile.values[rowNum]} and {targetFile2.values[rowNum]}")  # add 1 for header and 1 as it starts at 0
            counter += 1
        if counter >= maxErrorRowsShown:
            break


# MAIN
fileName = input("Enter a filename in the same folder: ")
fileName2 = input("Enter a second filename in the same folder: ")
maxErrorRowsShown = 20
listOfSheets = list(ExcelFile(f'{fileName}.xlsx').sheet_names)  # assuming that all sheet names are the same in each file
listOfColumnsPerSheet = []
notMatchingColumn = []  # clear for each new sheet
for sheet in listOfSheets:
    listOfColumns = list(read_excel(f'{fileName}.xlsx', sheet_name=sheet).columns)
    listOfColumnsPerSheet.append((sheet, listOfColumns))

checkAll = input(f"Do you want a summary of all sheets and their column? y/n: ")
print()

if checkAll == "y":
    for sheet in listOfSheets:
        checkSheet(sheet)

    if len(notMatchingColumn) != 0:
        for sheet, colName in notMatchingColumn:
            checkAllRows(sheet, colName)
            print()

choseColumn = input("Do you want to search the whole column y/n: ")
print()
while choseColumn == "y":
    print(f"Sheets: {listOfSheets}")
    sheet = input("Enter the sheet name to check (case sensitive): ")
    print()
    while not listOfSheets.__contains__(sheet):
        print("That isn't a sheet name please try again")
        sheet = input("Enter the sheet name to check (case sensitive): ")
        print()

    wholeSheet = input("Check the whole sheet y/n: ")
    if wholeSheet == "y":
        checkSheet(sheet)
    else:
        print(f"Columns for {sheet}: {listOfColumnsPerSheet[listOfSheets.index(sheet)][1]}")
        colName = input(f"Enter the column name to check from (case sensitive): ")
        print()
        while not listOfColumns.__contains__(colName):
            print("That isn't a column name please try again")
            colName = input(f"Enter the column name to check from (case sensitive): ")
            print()
        checkColumn(sheet, colName)
    choseColumn = input("Do you want to search the whole column y/n: ")
    print()

choseColumn = input("Do you want to search each row y/n: ")
print()
while choseColumn == "y":
    print(f"Sheets: {listOfSheets}")
    sheet = input("Enter the sheet name to check (case sensitive): ")
    print()
    while not listOfSheets.__contains__(sheet):
        print("That isn't a sheet name please try again")
        sheet = input("Enter the sheet name to check (case sensitive): ")
        print()

    print(f"Columns for {sheet}: {listOfColumnsPerSheet[listOfSheets.index(sheet)][1]}")
    colName = input(f"Enter the column name to check from (case sensitive): ")
    print()
    while not listOfColumns.__contains__(colName):
        print("That isn't a column name please try again")
        colName = input(f"Enter the column name to check from (case sensitive): ")
        print()

    checkAllRows(sheet, colName)
    choseColumn = input("Do you want to search each row y/n: ")
    print()

print("Terminating")