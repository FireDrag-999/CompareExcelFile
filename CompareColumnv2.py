from pandas import read_excel, ExcelFile, ExcelWriter
from logging import basicConfig, log, INFO
import logging
import os


def sortFiles():
    for sheet in listOfSheets:
        firstFile = firstFileSheetsStored[listOfSheets.index(sheet)]
        secondFile = secondFileSheetsStored[listOfSheets.index(sheet)]
        firstFile.sort_values(ascending=True, by=firstFile.columns[0], inplace=True)
        secondFile.sort_values(ascending=True, by=secondFile.columns[0], inplace=True)


def compareSheets():  # compares each entire row, then filters down to the incorrect columns in each row
    for sheet in listOfSheets:
        global currentSheet
        currentSheet = sheet
        errors = False
        firstFile = firstFileSheetsStored[listOfSheets.index(sheet)]
        secondFile = secondFileSheetsStored[listOfSheets.index(sheet)]

        if not firstFile.empty or not secondFile.empty:
            for row in firstFile.index:
                if not str(firstFile.values[row]) == str(secondFile.values[row]):
                    if not errors:
                        errors = True
                        listOfSheetsWithErrors.append(sheet)
                        print(f"{sheet} is not the same"), log(level=INFO, msg=f"{sheet} is not the same")
                        print(), log(level=INFO, msg="")

                    for column in firstFile.columns:
                        if firstFile[column][row] != secondFile[column][
                            row]:  # add 2 to each row, 1 for the header row, 1 because python starts at 0
                            print(f"{sheet}, {column}, row: [{row + 2}] is not the same, value: [{firstFile[column][row]} to {secondFile[column][row]}]"), log(level=INFO, msg=f"{sheet}, {column}, row: [{row + 2}] is not the same, value: [{firstFile[column][row]} to {secondFile[column][row]}]")
                            listOfErrors.append([sheet, column, row])  # stores all errors

                            for (sheetName, columnName,
                                 errorCount) in listOfColumnsErrorCount:  # increment the error count per different column errors
                                if sheet == sheetName and column == columnName:
                                    tempErrorCount = errorCount
                                    tempErrorCount += 1
                                    listOfColumnsErrorCount[
                                        listOfColumnsErrorCount.index((sheetName, columnName, errorCount))] = (
                                    sheetName, columnName, tempErrorCount)

            if not errors:
                print(f"{sheet} is the same"), log(level=INFO, msg=f"{sheet} is the same")
                print(), log(level=INFO, msg="")
        print(), log(level=INFO, msg="")

        # firstFileHighlightedSheetsStored.append(firstFile.style.apply(highlight, axis=None))  # highlight each error
        # secondFileHighlightedSheetsStored.append(secondFile.style.apply(highlight, axis=None))

        firstFile.style.apply(highlight, axis=1).to_excel('test.xlsx')
    printSummery()


def printSummery():  # summery of previous information
    print("SUMMERY: "), log(level=INFO, msg="SUMMERY: ")
    print(), log(level=INFO, msg="")

    for sheet in listOfSheets:
        firstFile = firstFileSheetsStored[listOfSheets.index(sheet)]
        secondFile = secondFileSheetsStored[listOfSheets.index(sheet)]

        if sheet in listOfSheetsWithErrors:
            print(f"{sheet} is not the same"), log(level=INFO, msg=f"{sheet} is not the same")
            print(), log(level=INFO, msg="")

        else:
            print(f"{sheet} is the same"), log(level=INFO, msg=f"{sheet} is not the same")
            print(), log(level=INFO, msg="")

        for column in firstFile.columns:
            try:  # prevents a situation where you sum a different datatype to integer/float
                firstFileSum = sum(firstFile[column].values)
                secondFileSum = sum(secondFile[column].values)

            except TypeError:
                firstFileSum = "not a number"
                secondFileSum = "not a number"

            for (sheetName, columnName, errorCount) in listOfColumnsErrorCount:
                if sheetName == sheet and columnName == column:  # finds the saved error count which is updated in compareSheets() and prints other info
                    if errorCount == 0:
                        print(f"{column} is the same, row count: {len(firstFile[column].index)} to {len(secondFile[column].index)}"), log(level=INFO, msg=f"{column} is the same, row count: {len(firstFile[column].index)} to {len(secondFile[column].index)}")
                        print(f"sum: {firstFileSum} to {secondFileSum}"), log(level=INFO, msg=f"sum: {firstFileSum} to {secondFileSum}")
                        print(f"error count: {errorCount}"), log(level=INFO, msg=f"error count: {errorCount}")
                        print(), log(level=INFO, msg="")

                    else:
                        print(f"{column} is not the same, row count: {len(firstFile[column].index)} to {len(secondFile[column].index)}"), log(level=INFO, msg=f"{column} is not the same, row count: {len(firstFile[column].index)} to {len(secondFile[column].index)}")
                        print(f"sum: {firstFileSum} to {secondFileSum}"), log(level=INFO, msg=f"sum: {firstFileSum} to {secondFileSum}")
                        print(f"error count: {errorCount}"), log(level=INFO, msg=f"error count: {errorCount}")
                        print(), log(level=INFO, msg="")


def highlight(x):
    df = x.copy()

    for (sheet, column, row) in listOfErrors:
        if sheet == currentSheet:
            df.loc[row, column] = ['background-color: red']
    return df


def createExcelFiles():
    if not os.path.exists(highlightedFilesPath):
        os.mkdir(highlightedFilesPath)

    createdFiles = False

    while not createdFiles:
        try:  # takes each file and sorts each sheet by first column then saves under sortedFiles folder
            with ExcelWriter(f'{highlightedFilesPath}\\{firstFileName}', engine="openpyxl") as writer:
                for sheet in listOfSheets:
                    firstFileHighlighted = firstFileHighlightedSheetsStored[listOfSheets.index(sheet)]
                    firstFileHighlighted.to_excel(writer, sheet_name=sheet, index=False)

            with ExcelWriter(f'{highlightedFilesPath}\\{secondFileName}', engine="openpyxl") as writer:
                for sheet in listOfSheets:
                    secondFileHighlighted = secondFileHighlightedSheetsStored[listOfSheets.index(sheet)]
                    secondFileHighlighted.to_excel(writer, sheet_name=sheet, index=False)

            createdFiles = True

        except PermissionError:
            print("Please close all the files in the highlighted folder")
            input("Press any key to try again: ")


def getExcelFiles():
    if not os.path.exists(filesPath):  # creates folders
        os.mkdir(filesPath)
        print(f"Please add the excel files to the files folder: {filesPath}")
        exit()

    listOfFiles = []
    count = 0

    for file in os.listdir(filesPath):  # pull the first two Excel files from the files folder
        if file.endswith(".xlsx") and count < 2:
            listOfFiles = listOfFiles + [file]
            count += 1

    fileNotPresent = True

    while fileNotPresent:  # prevents errors
        try:
            firstFileName = listOfFiles[0]
            secondFileName = listOfFiles[1]
            fileNotPresent = False

        except IndexError:
            print(f"Please add the excel files to the files folder: {filesPath}")
            input("Press any key once the file is in the folder ")

    fileOpen = True

    while fileOpen:  # prevents errors
        try:
            open(f'{filesPath}\\{firstFileName}', "r")
            open(f'{filesPath}\\{secondFileName}', "r")
            fileOpen = False

        except PermissionError:
            print("Permission error, please close the file.")
            input("Press any key once the file is closed: ")
            print()

    return firstFileName, secondFileName


def readAndStoreExcelFiles():
    for sheet in listOfSheets:  # reading and storing each sheet separately
        firstFile = read_excel(f'{filesPath}\\{firstFileName}', sheet_name=sheet)
        secondFile = read_excel(f'{filesPath}\\{secondFileName}', sheet_name=sheet)
        firstFileSheetsStored.append(firstFile)
        secondFileSheetsStored.append(secondFile)


def logsSetup():
    if not os.path.exists(logsPath):
        os.mkdir(logsPath)
    basicConfig(level=logging.INFO, filename=f"{logsPath}\\{firstFileName[0:-5]} log.txt", filemode="w", format="%(message)s")  # create log file, [0: -5] removes the .xlsx


def listOfColumnsErrorCountSetup():
    for sheet in listOfSheets:
        firstFile = firstFileSheetsStored[listOfSheets.index(sheet)]
        for column in firstFile.columns:
            listOfColumnsErrorCount.append((sheet, column, 0))


# Variables
filesPath = os.getcwd() + "\\files"
logsPath = os.getcwd() + "\\logs"
highlightedFilesPath = os.getcwd() + "\\highlightedFiles"

currentSheet = ""

firstFileSheetsStored = []
secondFileSheetsStored = []
firstFileHighlightedSheetsStored = []
secondFileHighlightedSheetsStored = []
listOfErrors = []
listOfColumnsErrorCount = []
listOfSheetsWithErrors = []

firstFileName, secondFileName = getExcelFiles()
logsSetup()

print()
print("All differences are given in the format [sheet name], [column name], row: [row number] is not the same, value: [first file value] to [second file value]"), log(level=INFO, msg="All differences are given in the format [sheet name], [column name], row: [row number] is not the same, value: [first file value] to [second file value]")
print(), log(level=INFO, msg="")

print(f"First file: {firstFileName}"), log(level=INFO, msg=f"First file: {firstFileName}")
print(f"Second file: {secondFileName}"), log(level=INFO, msg=f"Second file: {secondFileName}")
print(), log(level=INFO, msg="")

if input("Do you want to sort the files by the first column in ascending order y/n: ").lower() == "y":
    sortFiles()
print()

listOfSheets = list(ExcelFile(f'{filesPath}\\{firstFileName}').sheet_names)  # assuming that all sheet names are the same in each file
readAndStoreExcelFiles()
listOfColumnsErrorCountSetup()
compareSheets()
# createExcelFiles()

print("Terminating")
