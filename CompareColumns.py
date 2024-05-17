from pandas import read_excel, ExcelFile, ExcelWriter
from logging import basicConfig, log, INFO
import logging
import os

basicConfig(level=logging.INFO, filename="log", filemode="w", format="%(message)s")

# Problems: different headers cause errors

def checkSheet(sheet):
    listOfColumns = listOfColumnsPerSheet[listOfSheets.index(sheet)][1]
    targetFile = read_excel(f'{sortedFilesPath}\\{fileName}', sheet_name=sheet)
    targetFile2 = read_excel(f'{sortedFilesPath}\\{fileName2}', sheet_name=sheet)
    if not targetFile.empty or not targetFile2.empty:
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
    listOfColumns = listOfColumnsPerSheet[listOfSheets.index(sheet)][1]
    targetFile = read_excel(f'{sortedFilesPath}\\{fileName}', usecols=[listOfColumns.index(colName)], sheet_name=sheet)
    targetFile2 = read_excel(f'{sortedFilesPath}\\{fileName2}', usecols=[listOfColumns.index(colName)], sheet_name=sheet)

    if not targetFile.empty or not targetFile2.empty:
        if str(targetFile) == str(targetFile2):
            print(f"Sheet: {sheet}, Column: {colName} is the same"), log(level=INFO, msg=f"Sheet: {sheet}, Column: {colName} is the same")
        else:
            print(f"Sheet: {sheet}, Column: {colName} is not the same"), log(level=INFO, msg=f"Sheet: {sheet}, Column: {colName} is not the same")
            print(f"First file {sheet}, {colName}: has the length {len(targetFile)}"), log(level=INFO, msg=f"First file {sheet}, {colName}: has the length {len(targetFile)}")
            print(f"Second file {sheet}, {colName}: has the length {len(targetFile2)}"), log(level=INFO, msg=f"Second file {sheet}, {colName}: has the length {len(targetFile2)}")
            try:
                print(f"First file {sheet}, {colName}: has the sum {sum(targetFile.values)}"), log(level=INFO, msg=f"First file {sheet}, {colName}: has the sum {sum(targetFile.values)}")
                print(f"Second file {sheet}, {colName}: has the sum {sum(targetFile2.values)}"), log(level=INFO, msg=f"Second file {sheet}, {colName}: has the sum {sum(targetFile2.values)}")
            except TypeError:
                print(f"Sheet: {sheet}, {colName}: is text"), log(level=INFO, msg=f"Sheet: {sheet}, {colName}: is text")

            print(), log(level=INFO, msg="")
            if not notMatchingColumn.__contains__((sheet, colName)):
                notMatchingColumn.append((sheet, colName))
    else:
        print(f'Sheet: {sheet}, Column {colName} has no data'), log(level=INFO, msg=f'Sheet: {sheet}, Column {colName} has no data')


def checkAllRows(sheet, colName):
    listOfColumns = listOfColumnsPerSheet[listOfSheets.index(sheet)][1]
    counter = 0
    targetFile = read_excel(f'{sortedFilesPath}\\{fileName}', usecols=[listOfColumns.index(colName)], sheet_name=sheet)
    targetFile2 = read_excel(f'{sortedFilesPath}\\{fileName2}', usecols=[listOfColumns.index(colName)], sheet_name=sheet)

    for rowNum in range(0, len(targetFile)):
        try:
            if str(targetFile.values[rowNum]) != str(targetFile2.values[rowNum]):
                print(f"Sheet: {sheet}, Column: {colName}, row {rowNum + 2} is different: {targetFile.values[rowNum]} and {targetFile2.values[rowNum]}"), log(level=INFO, msg=f"Sheet: {sheet}, Column: {colName}, row {rowNum + 2} is different: {targetFile.values[rowNum]} and {targetFile2.values[rowNum]}")  # add 1 for header and 1 as it starts at 0
                counter += 1
            if counter >= maxErrorRowsShown:
                break
        except TypeError:
            break


# MAIN
# creates the folders to store files
filesPath = os.getcwd() + "\\files"
sortedFilesPath = os.getcwd() + "\\sortedFiles"
if not os.path.exists(filesPath):
    os.mkdir(filesPath)
    print(f"Please add the files to the files folder: ")
    exit()

if not os.path.exists(sortedFilesPath):
    os.mkdir(sortedFilesPath)

listOfFiles = []
count = 2
for file in os.listdir(filesPath):
    if file.endswith(".xlsx") and count > 0:
        listOfFiles = listOfFiles + [file]
        count -= 1

#  load all files
fileName = listOfFiles[0]
fileName2 = listOfFiles[1]

sort = input(f"Do you want to sort the files by the first column in ascending order? y/n: ")

if sort == "y":
    targetFile = read_excel(f'{filesPath}\\{fileName}')
    targetFile2 = read_excel(f'{filesPath}\\{fileName2}')
    listOfSheets = list(ExcelFile(f'{filesPath}\\{fileName}').sheet_names)  # assuming that all sheet names are the same in each file
    listOfSheets2 = list(ExcelFile(f'{filesPath}\\{fileName2}').sheet_names)  # assuming that all sheet names are the same in each file

    # takes each file and sorts each sheet by first column then saves under sortedFiles folder
    with ExcelWriter(f'{sortedFilesPath}\\{fileName}') as writer:
        for sheet in listOfSheets:
            targetFile = read_excel(f'{filesPath}\\{fileName}', sheet_name=sheet)
            targetFile.sort_values(ascending=True, by=targetFile.columns[0], inplace=True)
            targetFile.to_excel(writer, sheet_name=sheet, index=False)

    with ExcelWriter(f'{sortedFilesPath}\\{fileName2}') as writer:
        for sheet in listOfSheets2:
            targetFile2 = read_excel(f'{filesPath}\\{fileName2}', sheet_name=sheet)
            targetFile2.sort_values(ascending=True, by=targetFile2.columns[0], inplace=True)
            targetFile2.to_excel(writer, sheet_name=sheet, index=False)
else:
    sortedFilesPath = filesPath  # doesn't use the sorted files folder

maxErrorRowsShown = 20  # amount of rows that don't match shown per column
listOfSheets = list(ExcelFile(f'{sortedFilesPath}\\{fileName}').sheet_names)  # assuming that all sheet names are the same in each file
listOfColumnsPerSheet = []
notMatchingColumn = []  # clear for each new sheet
for sheet in listOfSheets:
    listOfColumns = list(read_excel(f'{sortedFilesPath}\\{fileName}', sheet_name=sheet).columns)
    listOfColumnsPerSheet.append((sheet, listOfColumns))

checkAll = input(f"Do you want a summary of all sheets and their column comparing {fileName} and {fileName2}? y/n: ")
print()

if checkAll == "y":
    for sheet in listOfSheets:
        checkSheet(sheet)

    if len(notMatchingColumn) != 0:
        for sheet, colName in notMatchingColumn:
            checkAllRows(sheet, colName)
            print()

choseColumn = input("Do you want to search the whole column y/n: ")
while choseColumn == "y":
    print()
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
        while not {listOfColumnsPerSheet[listOfSheets.index(sheet)][1]}.__contains__(colName):
            print("That isn't a column name please try again")
            colName = input(f"Enter the column name to check from (case sensitive): ")
            print()
        checkColumn(sheet, colName)
    choseColumn = input("Do you want to search the whole column y/n: ")

choseColumn = input("Do you want to search each row y/n: ")
while choseColumn == "y":
    print()
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
    while not {listOfColumnsPerSheet[listOfSheets.index(sheet)][1]}.__contains__(colName):
        print("That isn't a column name please try again")
        colName = input(f"Enter the column name to check from (case sensitive): ")

    checkAllRows(sheet, colName)
    choseColumn = input("Do you want to search each row y/n: ")
    print()

print("Terminating")