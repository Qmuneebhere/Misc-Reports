import datetime
import os
from sqlalchemy import create_engine
import pandas as pd
from datetime import date
import csv


section1 = '\n\n***********************************************************\n\n'

# Takes dataframe, removes SYN & Returns the new dataframe

def RemoveSYN(df):

    # ASCII code for SYN character

    synChr = '\x16'

    df = df.replace(synChr, "", regex=True)

    return df


# Function to check if a row contains special character

def Contains(row, removeList):

    # Joins a row into string

    row_string = ', '.join(map(str, row.values))

    for char in removeList:

        if row_string.count(char) > 0: return True
    
    else: return False


# Removes rows containing incorrect data

def RemoveChr(df, fileType):

    if fileType == 1: 
        
        # Sets the list of incorrect characters for Letter file

        remove = invalidLT

    elif fileType == 2: 
        
        remove = invalidEM
    
    # Gets a series of indexes having special characters

    incorrect = df.apply(lambda x: Contains(x, remove), axis=1)

    # Gets count of incorrect rows

    count = incorrect.sum()

    return df[~incorrect], count


# Checks date in each file and print resuls

def CheckDates(df, fileType):

    print(section1)

    print(f'Dates found in {fileType} Files: '.center(60) + '\n')

    dates = df['AttemptDate'].unique()

    for date in dates:

        print(date.center(60))


# Reads csv File and returns data in file as Lists of Lists

def ReadCsv(name):

    with open(name, 'r') as file:

        # Creates a csv reader object

        reader = csv.reader(file)

        dataList = [row for row in reader]

    return dataList


# --------------Lists of special characters that aren't allowed in LT, EM files--------------- #


invalidLT = ['~', '`', '@', '!', '$', '%', '^', '*', '(', ')',
             '+', '=', '[', ']', '{', '}', '|', ';', "'", '?', '<', '>']

invalidEM = ['~', '`', '!', '#', '$', '%', '^', '&', '*', '(', ')', '+',
             '=', '[', ']', '{', '}', '|', '\\', ':', ';', "'", '"', '?',
             '<', '>', '/']


# ---------------------------Prompts User for custom Date report------------------------------ #


print(section1)

while True:
    
    inp = input('Do you want to run report on custom dates (Y/N): ')
    
    if inp == 'Y' or inp == 'N': 
        
        if inp == 'Y': 
            
            custom = True

            customstartDate = input('\nEnter the beginning date in format mm/dd/yyyy: ')
            customendDate = input('\nEnter the closing date in format mm/dd/yyyy: ')

        else: custom = False
            
        break
        
    else: print('\nInvalid Input\n')


# --------------------------------Today's Date and Day No.------------------------------------ #


currDate = date.today()                  # Today's date
currDay = currDate.strftime('%w')        # Day number 0 - 6 (0 is sunday)
currYear = currDate.strftime('%Y')        # Day number 0 - 6 (0 is sunday)
yyyymmdd = currDate.strftime('%Y%m%d')   # Date in yyyymmdd format


# ------------------------------Names of Files to be saved------------------------------------ #


emName = 'EM.csv'
ltName = 'LT.csv'
phName = 'PH.csv'
fileName = 'ATTMPTUNI' + yyyymmdd + '.csv'


# --------------------------------Date of previous Sunday------------------------------------- #


endDate = date.today() - datetime.timedelta(days=int(currDay))


# --------------------------------Date of previous Monday------------------------------------- #


startDate = endDate - datetime.timedelta(days=6)


# -------------------------Dates as strings (Previous Mon & Sun)------------------------------ #


startDate = "'" + startDate.strftime("%m/%d/%Y") + "'"
endDate = "'" + endDate.strftime("%m/%d/%Y") + "'"

if custom:

    startDate = f"'{customstartDate}'"
    endDate = f"'{customendDate}'"

print('\n')
print(f'Start Date: {startDate}'.center(60))
print(f'End Date: {endDate}'.center(60))


# -------------------------------SQLAlchemy engine for MSSQL---------------------------------- #


engine = create_engine(
    "mssql+pyodbc://read_only:Neustar01@unifin-sql/"
    "tiger?driver=ODBC+Driver+17+for+SQL+Server")


##########################################################################################
##                                                                                      ##
##                             Query for fetching LT data                               ##
##                                                                                      ##
##########################################################################################


queryLT = "SELECT 'LT' AS RecordType, DBR_CLI_REF_NO AS BframeID, UDW_FLD1 AS RemoteID," \
          " convert(varchar, DAT_TRX_DATE_O, 112) as AttemptDate, " \
          "IIF(DAT_FORMS = 'E02' OR DAT_FORMS = 'E21', '7', '6') as AttemptMethodID," \
          "'5' as AttemptResultID," \
          "REPLACE(ADR_ADDR1, '', '') as Address1," \
          "REPLACE(ADR_ADDR2, '''', '') as Address2, ADR_CITY as City, ADR_STATE as State," \
          "RIGHT('00000' + CAST(REPLACE(LEFT(ADR_ZIP_CODE,5), '-', '') AS VARCHAR(5)),5) as ZipCode," \
          "IIF(DAT_LETTER_NO = '021', 2, IIF(DAT_LETTER_NO = '001', 3, IIF(DAT_LETTER_NO = '032', 4, " \
          "IIF(DAT_LETTER_NO IN ('ZRC', '034'), 5, 3)))) as LetterTypeID FROM CDS.DAT " \
          "INNER JOIN CDS.DBR on DBR_NO = DAT_DBR_NO " \
          "INNER JOIN CDS.UDW on UDW_DBR_NO = DAT_DBR_NO and UDW_SEQ = '0JC' " \
          "INNER JOIN CDS.ADR on ADR_DBR_NO = DAT_DBR_NO and ADR_SEQ_NO = '01' " \
          "WHERE DAT_TYPE = 'W' AND DAT_ACTION_CODE='LTR' AND DAT_FORMS NOT in ('E02','E21') " \
          "AND DAT_TRX_DATE_O BETWEEN " + startDate + " and " + endDate + " AND DBR_CLIENT LIKE 'JCAP%'"


##########################################################################################
##                                                                                      ##
##                             Query for fetching EM data                               ##
##                                                                                      ##
##########################################################################################


queryEM = "SELECT 'EM' as RecordType, DBR_CLI_REF_NO as BframeID," \
          "UDW_FLD1 as RemoteID, convert(varchar, Offer_Date, 112) as AttemptDate," \
          "'7' as AttemptMethodID, '5' as AttemptResultID, '2' as LetterTypeID," \
          "Offer_Email as EmailAddress, '0' as ArrangementMade FROM UFN.OFFERS " \
          "INNER JOIN CDS.DBR on DBR_NO = Account_Number " \
          "INNER JOIN CDS.UDW on UDW_DBR_NO = Account_Number and UDW_SEQ = '0JC' " \
          "WHERE Offer_Method = 2 AND Offer_Email IS NOT NULL " \
          "AND Offer_Date BETWEEN " + startDate + " and " + endDate + " AND DBR_CLIENT LIKE 'JCAP%'"


##########################################################################################
##                                                                                      ##
##                             Query for fetching PH data                               ##
##                                                                                      ##
##########################################################################################


queryPH = "SELECT 'PH' as 'Record Type', DBR_CLI_REF_NO as 'Bframe ID', UDW_FLD1 as 'Remote ID'," \
          "FORMAT(CallDateTime,'yyyyMMdd') as 'AttemptDate'," \
          "IIF(TypeOfConnect = 1, 4, 1) as 'Attempt Method ID'," \
          "IIF(RESULT_CODE IN ('017','050','051','052','053','054','055','040','041','042'),1," \
          "IIF(RESULT_CODE IN ('012'),2,IIF(RESULT_CODE IN ('015'),3,IIF(RESULT_CODE IN ('020'),4,5)))) as 'Attempt Result ID'," \
          "IIF(TypeOfConnect = 1, caller_id_number, destination_number) as 'Phone Number Attempted'," \
          "FORMAT(CallDateTime,'hhmm') as 'Attempt Start Time'," \
          "FORMAT(DATEADD(ss,ISNULL(CallDuration,0),CallDateTime),'hhmm') as 'Attempts End Time'," \
          "IIF(RESULT_CODE IN ('040','041','042'),1,0) as 'Arrangement Made Indicator' " \
          "FROM DCCollectOne.dbo.UnifinCDR WITH(NOLOCK) " \
          "INNER JOIN tiger.CDS.DBR WITH(NOLOCK) on AccountNumber = DBR_NO " \
          "INNER JOIN tiger.CDS.UDW WITH(NOLOCK) on AccountNumber = UDW_DBR_NO and UDW_SEQ = '0JC' " \
          "WHERE CallDateTime BETWEEN " + startDate + " and " + endDate + " AND DBR_CLIENT LIKE 'JCAP%' " \
          "and LEN(IIF(TypeOfConnect = 1, caller_id_number, destination_number)) = 10 " \
                                                                          "UNION " \
          "SELECT  'PH' as RecordType, DBR_CLI_REF_NO as BframeID, UDW_FLD1 as RemoteID," \
          "convert(varchar, DAT_TRX_DATE_O, 112) as AttemptDate, '3' as AttemptMethodID, '1' as AttemptResultID," \
          "Offer_Phone as 'Phone Number Attempted', CONCAT(DAT_TRX_HH,DAT_TRX_MM) as 'Attempt Start Time'," \
          "CONCAT(DAT_TRX_HH,DAT_TRX_MM) as 'Attempts End Time', '0' as 'Arrangement Made Indicator' " \
          "FROM CDS.DAT WITH(NOLOCK) " \
          "INNER JOIN CDS.DBR WITH(NOLOCK) on DBR_NO = DAT_DBR_NO INNER JOIN CDS.UDW WITH(NOLOCK) on UDW_DBR_NO = DAT_DBR_NO and UDW_SEQ = '0JC' " \
          "LEFT JOIN UFN.Offers WITH(NOLOCK) on Account_Number = DAT_DBR_NO AND DAT_TRX_DATE_O=Offer_Date " \
          "WHERE DAT_ACTION_CODE='SMS' and Offer_Method=3 and Offer_Phone is not NULL " \
          "AND DAT_TRX_DATE_O BETWEEN " + startDate + " and " + endDate + " AND DBR_CLIENT LIKE 'JCAP%'"


# ---------------------------Path to folder where Files are saved----------------------------- #


curr_Dir = os.getcwd()


print(section1)
print('Fetching Data from SQL'.center(60))
print(section1)


# ----------------------------------EM, LT, PH dataframes------------------------------------- #


dfLT = pd.read_sql(queryLT, engine, dtype=str)

print(f'LT dataframe created. Count:{len(dfLT)}'.center(60))

dfEM = pd.read_sql(queryEM, engine, dtype=str)

print(f'EM dataframe created. Count:{len(dfEM)}'.center(60))


dfPH = pd.read_sql(queryPH, engine, dtype=str)

print(f'PH dataframe created. Count:{len(dfPH)}'.center(60))


# ------------------------------Does QA of EM & LT dataframes--------------------------------- #


print(section1)

if dfLT.shape[0] != 0:
    dfLT, countLT = RemoveChr(dfLT, 1)
    print(f'Incorrect rows in LT Data: {countLT}'.center(60))

if dfEM.shape[0] != 0:
    dfEM, countEM = RemoveChr(dfEM, 2)
    print(f'Incorrect rows in EM Data: {countEM}'.center(60))


# --------------------------------SYN Removal & Date Check------------------------------------ #

if dfEM.shape[0] != 0:
    dfEM = RemoveSYN(dfEM)
    CheckDates(dfEM, 'EM')

if dfLT.shape[0] != 0:
    dfLT = RemoveSYN(dfLT)
    CheckDates(dfLT, 'LT')

if dfPH.shape[0] != 0:
    dfPH = RemoveSYN(dfPH)
    CheckDates(dfPH, 'PH')


##########################################################################################
##                                                                                      ##
##               Dataframes having different number of columns can't be                 ##
##               merged together. Therefore, files are saved as csv and                 ##
##                    then these csv files are read using ReadCsv.                      ##
##                                                                                      ##
##########################################################################################


os.chdir(path=curr_Dir)

dfEM.to_csv(emName, index=False, header=False)
dfLT.to_csv(ltName, index=False, header=False)
dfPH.to_csv(phName, index=False, header=False)


# ----------------------Allows user to check files before proceeding-------------------------- #


print(section1)
print('PH, EM, LT files are created.'.center(60) + '\n')

while True:

    inp = input('Concatenate them? (Y): ')

    if inp == 'Y': break
    else: print('\nInvalid Input.\n')


# ----------------------------------Reads PH, EM, LT files------------------------------------ #


emData = ReadCsv(emName)
ltData = ReadCsv(ltName)
phData = ReadCsv(phName)


# -----------------------------------Create Todays Folder------------------------------------- #

# Path to todays folder

curr_Folder = f"{curr_Dir}\\{currYear}\\{currDate.strftime('%m-%d-%Y')}\\"
os.makedirs(curr_Folder)


##########################################################################################
##                                                                                      ##
##                 ReadCsv returns data in the form of list of lists.                   ##
##                 These lists are then concatenated to prepare final                   ##
##                 File which is then saved as ATTMPTUNIyyyymmdd.csv                    ##
##                                                                                      ##
##########################################################################################


# List for adding Header row

header = [['HDR', yyyymmdd, 'UNI', '', '', '', '', '', '', '']]

data = header + phData + ltData + emData

with open(curr_Folder + fileName, 'w', newline='') as file:

    writer = csv.writer(file)
    writer.writerows(data)
