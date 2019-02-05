import gspread
from oauth2client.service_account import ServiceAccountCredentials
from fogbugz import FogBugz
from logins import getcolumns, getapikey, getpath, getquery1
from datetime import datetime

class FogbugzCase:
   def __init__(self, bugId, title, personAssignedTo, status, lastUpdated):
       self.bugId = bugId
       self.title = title
       self.personAssignedTo = personAssignedTo
       self.status = status
       self.lastUpdated = lastUpdated

# Google API Stuff
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)
sheet = client.open("FogBugzGoals").get_worksheet(0)
sheet2 = client.open("FogBugzGoals").get_worksheet(1)

# Fogbugz stuff
fb = FogBugz(getpath(), getapikey(), api_version=8)
print("Starting")
results = fb.search(q=getquery1(), cols=getcolumns()).prettify()
print("Bye")

resultListOfCases = list()

print('Transforming XML to list of objects')

for c in results:
    caseObject = FogbugzCase(c.ixBug.string, c.sTitle.string, c.sPersonAssignedTo.string, c.sStatus.string,
                             datetime.strptime(c.dtLastUpdated.string, CONST_DATE_FORMAT))
    resultListOfCases.append(caseObject)

resultListOfCases.sort(key=lambda r: r.lastUpdated)
#print(results)






