import os
import shutil
import openpyxl
import datetime
from datetime import date
import pandas as pd
from sqlalchemy import create_engine


section1 = '\n\n***********************************************************\n\n'

##########################################################################
#########################                        #########################
######################### Author: Muneeb Qureshi #########################
#########################                        #########################
##########################################################################


##########################################################################
##                                                                      ##
##        This function returns txt file names both old and new         ##
##                                                                      ##
##########################################################################

def FileNames():

    # Date pattern for old name

    oDate = date.today().strftime('%Y%m%d')

    # Date pattern for old name

    nDate = date.today().strftime('%m%d%y')

    # Text file old name

    oFile = 'UNIPAY' + oDate + '.txt'

    # Text file new name

    nFile = 'UNI' + nDate + '.txt'

    return oFile, nFile


##########################################################################
##                                                                      ##
##        This function moves template invoice to current folder.       ##
##                                                                      ##
##########################################################################

def GetFiles(path1, path2, date):

    # Copies Template file to current folder

    oldName1 = 'Invoice_Template.xlsx'

    prevFile1 = path1 + oldName1

    newName1 = 'UNI-SCS Remit Invoice ' + date + '.xlsx'

    newFile1 = path2 + newName1

    shutil.copy(prevFile1, newFile1)

    # Moves txt file to current folder

    oldName2, newName2 = FileNames()

    prevFile2 = path1 + oldName2

    newFile2 = path2 + newName2

    shutil.move(prevFile2, newFile2)


##########################################################################
##                                                                      ##
##         SCS Payment file is a text file with '|' as delimeter        ##
##        The data from text file is a long string which needs          ##
##                          to be structured.                           ##
##                                                                      ##
##        The for loop executes over each character in string.          ##
##        Each character is appended to elem untill char == '|'         ##
##        Once char == '|' means elem is completed it is then           ##
##        added to a list named row, elem keeps adding to row           ##
##        until char == '\n' which means row is complete. Row           ##
##                    is then appended to rowsList.                     ##
##                                                                      ##
##        This function returns dataframe containing all rows           ##
##               also returns the rowCount in dataframe.                ##
##                                                                      ##
##########################################################################

def ReadFile(path, txtFile):

    with open(path + txtFile, 'r') as file:
        
            txtData = file.read()

    # Total characters in text file

    charTotal = len(txtData)

    elem = ''
    row = []
    rowsList = []

    for i in range(charTotal):

        if txtData[i] == '|':

            row.append(elem)
            elem = ''
        
        elif txtData[i] == '\n':
            
            row.append(elem)
            rowsList.append(row)
            row = []
            elem = ''

        else:

            elem = elem + txtData[i]

            if i == charTotal - 1:

                row.append(elem)
                rowsList.append(row)


    # Above for loop creates a list of list containing all data. Headers and 
    # last row are saved separately and removed. Remaining data is converted 
    # to a dataframe.

    headers = rowsList[0]
    lastRow = rowsList[-1]

    # All data except headers and last row

    data = rowsList[1:-1]
    
    # List of Column names
    
    colNames = ['delete1', 'delete2', 'delete3', 'SCS Account Number', 'delete4', 'Payment Date',
                'Payment Amount', 'Agency Fee','Payment Type', 'delete5', 'delete6', 'delete7',
                'SIF or PIF?']

    # Creating dataframe using List of Lists named data

    df = pd.DataFrame(data, columns=colNames)
    
    colDelete = ['delete1', 'delete2', 'delete3', 'delete4', 'delete5', 'delete6', 'delete7']
    
    df = df.drop(colDelete, axis=1)

    # Converting Column Payment Amount and Agency Fee to float
    
    df['Payment Amount'] = df['Payment Amount'].astype(float)
    
    df['Agency Fee'] = df['Agency Fee'].astype(float)

    # Calculating Due SCS Column
    
    dueSCS = df['Payment Amount'] - df['Agency Fee']
    
    # Inserting Due SCS Column in dataframe at 4th place

    df.insert(4, 'Due SCS', dueSCS)

    # Formatting Numeric Columns
    
    numCols = ['Payment Amount', 'Agency Fee', 'Due SCS']
    
    df[numCols] = df[numCols].apply(lambda x: x.round(2))
    
    df[numCols] = df[numCols].map(lambda x: "{:.2f}".format(x))

    # Gets the number of rows in dataframe

    rowCount = df.shape[0]

    if rowCount != int(lastRow[1]):

        print('\n' + 'Value in last row does not equal row Count.')
    
    df.sort_values('SCS Account Number', inplace=True)
    
    return df, rowCount


##########################################################################
##                                                                      ##
##        This function takes start and end date and pulls data         ##
##        from database & returns a dataframe containing 5 cols         ##
##                                                                      ##
##########################################################################

def ReadSQL(startDate, endDate):

    engine = create_engine("mssql+pyodbc://read_only:Neustar01@unifin-sql/tiger?driver=ODBC+Driver+17+for+SQL+Server")

    queryDF = "SELECT dbr_cli_ref_no AS 'SCS Account Number', DBR_NAME1 AS 'Consumer Name', TRS_AMT AS 'Payment Amount', \
               TRS_COMM_AMT AS 'Agency Fee', (TRS_AMT - TRS_COMM_AMT) AS 'Due SCS', dbr_no, dbr_status, \
	          TRS_TRUST_CODE, CONVERT(VARCHAR, TRS_TRX_DATE_O, 101) AS [Transaction Date] \
              FROM cds.trs trs \
              INNER JOIN cds.dbr dbr ON trs.TRS_DBR_NO = dbr.DBR_NO \
              WHERE DBR_CLIENT like 'SCS%' \
              AND TRS_TRX_DATE_O BETWEEN '" + startDate + "' AND '" + endDate + "' AND \
              trs_amt <> 0 AND TRS_TRUST_CODE not in ('2', '5', '14', '30', '31', '32', '33', '34', '35') \
              ORDER BY dbr_cli_ref_no"

    df = pd.read_sql(queryDF, engine, dtype=str)
    
    cols = ['SCS Account Number', 'Consumer Name', 'Payment Amount', 'Agency Fee', 'Due SCS']
    
    numCols = ['Payment Amount', 'Agency Fee', 'Due SCS']
    
    df[numCols] = df[numCols].astype(float)
    
    df[numCols] = df[numCols].map(lambda x: "{:.2f}".format(x))
    
    df.sort_values('SCS Account Number', inplace=True)

    return df[cols]


##########################################################################
##                                                                      ##
##         This function takes dataframe & returns a dictionary         ##
##            with Acc No as Keys and ConsumerName as values.           ##
##                                                                      ##
##########################################################################

def ConsumerName(df):
    
    nameDict = {}
    
    for index, row in df.iterrows():
        
        nameDict[row['SCS Account Number']] = row['Consumer Name']
    
    return nameDict


# --------------------------------Paths to Folders-------------------------------- #


# Current date in folder name format

cDate = date.today().strftime('%m-%d-%Y')
cYear = date.today().strftime('%Y')

# Path to current remits folder

oPath = f'{os.getcwd()}\\'

# Path to current date folder

cPath = f'{oPath}\\Remit Invoices\\{cYear}\\{cDate}\\'


# ---------------------------------Start/End Date--------------------------------- #


endDate = date.today() - datetime.timedelta(days=1)
startDate = endDate - datetime.timedelta(days=6)

# Dates in string format for SQL query

sDate = startDate.strftime('%m-%d-%Y')
eDate = endDate.strftime('%m-%d-%Y')

print(section1)

while True:
    
    inp = input('Do you want to run report on custom dates (Y/N): ')
    
    if inp == 'Y':

        print(section1)

        sDate = input('Enter the beginning date in format mm-dd-yyyy: ')
        eDate = input('Enter the closing date in format mm-dd-yyyy: ')
        break
    
    elif inp == 'N': break
        
    else: print('\nInvalid Input\n')

print(section1)
print(f'Report Start Date: {sDate}'.center(60))
print(f'Report End Date: {eDate}'.center(60))


# -------------------------Creates folder of Today's Date------------------------- #


os.makedirs(cPath)


# ------------------Moving Text File & Invoice to current Folder------------------ #


GetFiles(oPath, cPath, cDate)


# ---------------------------------Runs SQL Query--------------------------------- #


# Calls Read SQL function to get dataframe

dfSQL = ReadSQL(sDate, eDate)

# Gets row count of dataframe from SQL

rowCount1 = dfSQL.shape[0]


# --------------------Reads Text File & Gets dataframe, rowCount------------------ #


# Gets text file names old and new one

otxtFile, ctxtFile = FileNames()

# Returns dataframe and rowCount

dfTextFile, rowCount2 = ReadFile(cPath, ctxtFile)


# --------------------Compares Row Count of both Text File & SQL------------------ #


if rowCount1 != rowCount2:

    print('\n' + 'Row Count of Text File and SQL does not match.')

else:

    print(section1)
    print(f'Number of Accounts Today: {rowCount1}'.center(60))


# ---------------------Gets names Dictionary from SQL dataframe------------------- #


namesDict = ConsumerName(dfSQL)
nameList = []

# Loops over each row of text file dataframe and creates 
# a list of name to be attached as column in dataframe.

for index, row in dfTextFile.iterrows():
    
    key = row['SCS Account Number']
    name = namesDict[key]
    nameList.append(name)

# Saves the data file in folder as data to compare

dfTextFile.insert(1, 'Consumer Name', nameList)
dfTextFile.to_csv(cPath + 'data.csv', index=False)