from pandas import read_excel, ExcelFile


fileName = input("Enter a filename in the same folder: ")
fileName2 = input("Enter a filename in the same folder: ")
amountOfHeaderRows = 1  # assuming that there is only one row of headers
listOfColumns = list(read_excel(f'{fileName}.xlsx').columns)  # assuming that all column names are the same in each file
listOfSheets = list(ExcelFile(f'{fileName}.xlsx').sheet_names)  # assuming that all sheet names are the same in each file

for sheet in listOfSheets:
    notMatchingColumn = []  # clear for each new sheet
    print()
    targetFile = read_excel(f'{fileName}.xlsx', sheet_name=sheet, skiprows=amountOfHeaderRows)
    targetFile2 = read_excel(f'{fileName2}.xlsx', sheet_name=sheet, skiprows=amountOfHeaderRows)
    targetFile.sort_values(ascending=True, by=targetFile.columns[0])  # sorts by first column or only column
    targetFile2.sort_values(ascending=True, by=targetFile2.columns[0])

    if targetFile.equals(targetFile2):
        print(f"Sheet: {sheet} is the same")
    else:
        print(f"Sheet: {sheet} is not the same")

    for colName in listOfColumns:
        targetFile = read_excel(f'{fileName}.xlsx', usecols=[listOfColumns.index(colName)], skiprows=amountOfHeaderRows, sheet_name=sheet)
        targetFile2 = read_excel(f'{fileName2}.xlsx', usecols=[listOfColumns.index(colName)], skiprows=amountOfHeaderRows, sheet_name=sheet)
        targetFile.sort_values(ascending=True, by=targetFile.columns[0])
        targetFile2.sort_values(ascending=True, by=targetFile2.columns[0])

        if targetFile.equals(targetFile2):
            print(f"Sheet: {sheet}, Column: {colName} is the same")
        else:
            print(f"Sheet: {sheet}, Column: {colName} is not the same")
            notMatchingColumn.append(colName)
    print()

    skip = input(f"""Hit enter to search each row for
Sheet: {sheet}, Columns: {notMatchingColumn}
Type 'skip' to go to the next sheet without searching by rows: """).lower()

    if skip != "skip":
        for colName in notMatchingColumn:
            targetFile = read_excel(f'{fileName}.xlsx', usecols=[listOfColumns.index(colName)], skiprows=amountOfHeaderRows, sheet_name=sheet)
            targetFile2 = read_excel(f'{fileName2}.xlsx', usecols=[listOfColumns.index(colName)], skiprows=amountOfHeaderRows, sheet_name=sheet)

            for num in range(0, len(targetFile)):
                if targetFile.values[num] != targetFile2.values[num]:
                    print(
                        f"Sheet: {sheet}, Column: {colName}, row {num + 2} is different: {targetFile.values[num]} and {targetFile2.values[num]}")  # add 1 for header and 1 as it starts at 0
