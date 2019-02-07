import gspread
from oauth2client.service_account import ServiceAccountCredentials
from fogbugz import FogBugz
from s import getcolumns, getapikey, getpath, getday0, getoldestcasesnoupdate, getprofilter, getoldestresolvedcases
from datetime import datetime
import time

class FogbugzCase:
   def __init__(self, bugId, title, personAssignedTo, status, lastUpdated):
       self.bugId = bugId
       self.title = title
       self.personAssignedTo = personAssignedTo
       self.status = status
       self.lastUpdated = lastUpdated


def initgooglesheets():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
    client = gspread.authorize(creds)
    sheet = client.open("General Kaunas Stats").get_worksheet(1)
    return sheet


def getcaseslist(fb, query):
    print("Started gathering data...")
    start = time.time()
    resultXML = fb.search(q=query, cols=getcolumns())
    end = time.time()
    print('Finished gathering data. Time spent: ' + str(end-start))

    resultXMLList = list(resultXML.cases.childGenerator())
    resultListOfCases = list()
    for c in resultXMLList:
        caseObject = FogbugzCase(c.ixBug.string, c.sTitle.string, c.sPersonAssignedTo.string, c.sStatus.string,
                                 datetime.strptime(c.dtLastUpdated.string, CONST_DATE_FORMAT))
        resultListOfCases.append(caseObject)

    resultListOfCases.sort(key=lambda r: r.lastUpdated)

    return resultListOfCases


def writeoldestcases(sheet, listOfCases, row, col):
    print('Adding data to Sheets: 5 oldest cases')
    i = 0
    for c in listOfCases:
        if i > 4:
            break
        sheet.update_cell(row + i, col, c.bugId)
        sheet.update_cell(row + i, col + 1, c.title)
        sheet.update_cell(row + i, col + 2, c.personAssignedTo)
        sheet.update_cell(row + i, col + 3, c.status)
        sheet.update_cell(row + i, col + 4, c.lastUpdated.strftime(FORMAT_OF_DATE))
        i = i + 1


def writefiltercount(firstTimeWriting, sheet, dateOfData, value, row, col, goal):
    print('Adding data to Sheets: Case count')
    if not firstTimeWriting:
        dateOfLastDataString = sheet.cell(row, col).value
        dateOfLastData = datetime.strptime(dateOfLastDataString, FORMAT_OF_DATE)
        if dateOfData.date() != dateOfLastData.date():
            appendrowfordata(sheet)
    sheet.update_cell(row, col, dateOfData.strftime(FORMAT_OF_DATE))
    sheet.update_cell(row, col + 1, value)
    sheet.update_cell(row, col + 2, str(goal))


def writeday0age(sheet, oldestCaseDate, row, col, goal):
    print('Adding data to Sheets: Day0 oldest case age')
    oldestCaseAge = (datetime.now() - oldestCaseDate).days
    sheet.update_cell(row, col + 4, str(oldestCaseAge))
    sheet.update_cell(row, col + 5, str(goal))


def writeoldestcases(sheet, listOfCases, row, col):
    print('Adding data to Sheets: Oldest cases without an update')
    i = 0
    for c in listOfCases:
        if i > 4:
            break
        sheet.update_cell(row + i, col, c.bugId)
        sheet.update_cell(row + i, col + 1, c.title)
        sheet.update_cell(row + i, col + 2, c.personAssignedTo)
        sheet.update_cell(row + i, col + 3, c.status)
        sheet.update_cell(row + i, col + 4, c.lastUpdated.strftime(FORMAT_OF_DATE))
        i = i + 1


def writeoldestcasenoupdateage(sheet, oldestCaseDate, row, col, goal):
    print('Adding data to Sheets: Oldest cases without update age')
    oldestCaseAge = (datetime.now() - oldestCaseDate).days
    sheet.update_cell(row, col, datetime.now().strftime(FORMAT_OF_DATE))
    sheet.update_cell(row, col + 1, str(oldestCaseAge))
    sheet.update_cell(row, col + 2, str(goal))


def appendrowfordata(sheet):
    sheet.insert_row([], dataRowStartIndex)


def main():
    sheets = initgooglesheets()

    # Fogbugz stuff
    fb = FogBugz(getpath(), getapikey(), api_version=8)

    # ----------------------------------------------------------------------------------------------------------
    # Getting Day0
    print('Query set to day0')
    query = getday0()
    resultListOfCases = getcaseslist(fb, query)

    firstTimeWriting = False
    listLength = str(len(resultListOfCases))
    oldestCaseDate = resultListOfCases[0].lastUpdated

    # Calling all the information gathering functions
    writeoldestcases(sheets, resultListOfCases, 3, 1)
    writefiltercount(firstTimeWriting, sheets, datetime.now(), listLength, 10, 1, 500)
    writeday0age(sheets, oldestCaseDate, 10, 1, 7)
    print('Day 0 query finished')
    # ----------------------------------------------------------------------------------------------------------

    # ----------------------------------------------------------------------------------------------------------
    # Continuing non day0 queues
    # -----------------------------------------------------------------------------------------------------------
    # Getting oldest cases assigned to us without update
    print('Query set to oldest case without update')
    query = getoldestcasesnoupdate()
    resultListOfCases = getcaseslist(fb, query)
    oldestCaseDate = resultListOfCases[0].lastUpdated

    # Calling all the information gathering functions
    writeoldestcases(sheets, resultListOfCases, 3, 8)
    writeoldestcasenoupdateage(sheets, oldestCaseDate, 10, 8, 10)
    print('Oldest cases without an update query finished')
    # ----------------------------------------------------------------------------------------------------------

    # ----------------------------------------------------------------------------------------------------------
    # Getting pro filter cases
    print('Query set to pro filter')
    query = getprofilter()
    resultListOfCases = getcaseslist(fb, query)
    listLength = str(len(resultListOfCases))

    # Calling all the information gathering functions
    writeoldestcases(sheets, resultListOfCases, 3, 15)
    writefiltercount(True, sheets, datetime.now(), listLength, 10, 15, 0)
    print('Pro filter query finished')
    # ----------------------------------------------------------------------------------------------------------
    # Getting oldest resolved cases assigned to us
    print('Query set to oldest cases resolved and assigned to us')
    query = getoldestresolvedcases()
    resultListOfCases = getcaseslist(fb, query)
    oldestCaseDate = resultListOfCases[0].lastUpdated

    # Calling all the information gathering functions
    writeoldestcases(sheets, resultListOfCases, 3, 22)
    writeoldestcasenoupdateage(sheets, oldestCaseDate, 10, 22, 30)
    print('Oldest resolved cases without an update query finished')
    # ----------------------------------------------------------------------------------------------------------
    # ----------------------------------------------------------------------------------------------------------
    # ----------------------------------------------------------------------------------------------------------

    print('Data writing to Sheets finished')
    print("Good to go")


CONST_DATE_FORMAT = '%Y-%m-%dT%H:%M:%SZ'
FORMAT_OF_DATE = '%Y-%m-%d %H:%M:%S'
dataRowStartIndex = 10

main()





