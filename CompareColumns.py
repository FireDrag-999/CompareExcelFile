from pandas import read_excel, ExcelFile


def checkSheet(sheet):
    print()
    targetFile = read_excel(f'{fileName}.xlsx', sheet_name=sheet, skiprows=amountOfHeaderRows)
    targetFile2 = read_excel(f'{fileName2}.xlsx', sheet_name=sheet, skiprows=amountOfHeaderRows)
    if targetFile.empty or not targetFile2.empty:
        targetFile.sort_values(ascending=True, by=targetFile.columns[0])  # sorts by first column or only column
        targetFile2.sort_values(ascending=True, by=targetFile2.columns[0])
        if targetFile.equals(targetFile2):
            print(f"Sheet: {sheet} is the same")
        else:
            print(f"Sheet: {sheet} is not the same")

        for colName in listOfColumns:
            checkColumn(sheet, colName)
    else:
        print(f'Sheet: {sheet} has no data')
    print()

def checkColumn(sheet, colName):
    targetFile = read_excel(f'{fileName}.xlsx', usecols=[listOfColumns.index(colName)], skiprows=amountOfHeaderRows, sheet_name=sheet)
    targetFile2 = read_excel(f'{fileName2}.xlsx', usecols=[listOfColumns.index(colName)], skiprows=amountOfHeaderRows, sheet_name=sheet)

    if not targetFile.empty or targetFile2.empty:
        targetFile.sort_values(ascending=True, by=targetFile.columns[0])
        targetFile2.sort_values(ascending=True, by=targetFile2.columns[0])

        if targetFile.equals(targetFile2):
            print(f"Sheet: {sheet}, Column: {colName} is the same")
        else:
            print(f"Sheet: {sheet}, Column: {colName} is not the same")
            if not notMatchingColumn.__contains__((sheet, colName)):
                notMatchingColumn.append((sheet, colName))
    else:
        print(f'Sheet: {sheet}, Column {colName} has no data')

def checkAllRows(sheet, colName):
    counter = 0
    targetFile = read_excel(f'{fileName}.xlsx', usecols=[listOfColumns.index(colName)], skiprows=amountOfHeaderRows, sheet_name=sheet)
    targetFile2 = read_excel(f'{fileName2}.xlsx', usecols=[listOfColumns.index(colName)], skiprows=amountOfHeaderRows, sheet_name=sheet)

    for rowNum in range(0, len(targetFile)):
        if targetFile.values[rowNum] != targetFile2.values[rowNum]:
            print(f"Sheet: {sheet}, Column: {colName}, row {rowNum + 2} is different: {targetFile.values[rowNum]} and {targetFile2.values[rowNum]}")  # add 1 for header and 1 as it starts at 0
            counter += 1
        if counter >= maxErrorRowsShown:
            break

#
#  problems to fix: error when n, n are inputted
#  prevent duplicates in notMatchingColumns

fileName = input("Enter a filename in the same folder: ")
fileName2 = input("Enter a filename in the same folder: ")
amountOfHeaderRows = 1  # assuming that there is only one row of headers
maxErrorRowsShown = 20
listOfSheets = list(ExcelFile(f'{fileName}.xlsx').sheet_names)  # assuming that all sheet names are the same in each file
notMatchingColumn = []  # clear for each new sheet

choseColumn = input("Do you want to specifically choose a column to search column by column y/n: ")
print()
while choseColumn == "y":
    print(f"Sheets: {listOfSheets}")
    sheet = input("Enter the sheet name to check (case sensitive): ")
    print()
    while not listOfSheets.__contains__(sheet):
        print("That isn't a sheet name please try again")
        sheet = input("Enter the sheet name to check (case sensitive): ")
        print()

    listOfColumns = list(read_excel(f'{fileName}.xlsx', sheet_name=sheet).columns)  # assuming that all column names are the same in each file
    wholeSheet = input("Check the whole sheet y/n: ")
    if wholeSheet == "y":
        checkSheet(sheet)
    else:
        print(f"Columns for {sheet}: {listOfColumns}")
        colName = input(f"Enter the column name to check from (case sensitive): ")
        print()
        while not listOfColumns.__contains__(colName):
            print("That isn't a column name please try again")
            colName = input(f"Enter the column name to check from (case sensitive): ")
            print()
        checkColumn(sheet, colName)
    choseColumn = input("Do you want to specifically choose a column to search column by column y/n: ")
    print()

skipToRows = input(f"""Do you want to search through all sheets
These are the current (sheet,column) that aren't matching: {notMatchingColumn}
If above is blank you must declare your own columns to check row by row y/n: """)
print()

if skipToRows == "y":
    for sheet in listOfSheets:
        listOfColumns = list(read_excel(f'{fileName}.xlsx', sheet_name=sheet).columns)
        checkSheet(sheet)

choseColumn = input("Do you want to specifically choose a column to search row by row y/n: ")
print()
while choseColumn == "y":
    print(f"Sheets: {listOfSheets}")
    sheet = input("Enter the sheet name to check (case sensitive): ")
    print()
    while not listOfSheets.__contains__(sheet):
        print("That isn't a sheet name please try again")
        sheet = input("Enter the sheet name to check (case sensitive): ")
        print()

    listOfColumns = list(read_excel(f'{fileName}.xlsx', sheet_name=sheet).columns)  # assuming that all column names are the same in each file
    print(f"Columns {listOfColumns}")
    colName = input(f"Enter the column name to check from (case sensitive): ")
    print()
    while not listOfColumns.__contains__(colName):
        print("That isn't a column name please try again")
        colName = input(f"Enter the column name to check from (case sensitive): ")
        print()

    checkAllRows(sheet, colName)
    choseColumn = input("Do you want to specifically choose a column to search row by row y/n: ")
    print()

if len(notMatchingColumn) != 0:
    for sheet, colName in notMatchingColumn:
        checkAllRows(sheet, colName)

print("Terminating")