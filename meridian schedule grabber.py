from __future__ import print_function

import datetime
import os.path
import math

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError



def classTypeOf(classString):
    humstrings = ['hum','Hum','HUM','humanities','Humanities','HUMANITIES']
    mststrings = ['mst','Mst','MST']
    espstrings = ['NovE','NH','I1','I2','A1','A2','IMS']
    frnstrings = ['NovF']
    artstrings = ['PE']
    lunstrings = ['Lunch','MSMUNch','HSMUNch']
    comstrings = ['Community Groups','SLAB','Grade Meeting','Health','Mtg']
    for i in humstrings:
        if i in classString:
            return 'hum'
    for i in mststrings:
        if i in classString:
            return 'mst'
    for i in espstrings:
        if i in classString:
            return 'esp'
    for i in frnstrings:
        if i in classString:
            return 'frn'
    for i in lunstrings:
        if i in classString:
            return 'lun'
    for i in comstrings:
        if i in classString:
            return 'com'
    for i in artstrings:
        if i in classString:
            return 'art'
    return 'elc'

def nIntoK(n,k):
    widthsList = [math.ceil(k/n) for i in range(n)]
    overshootsby = (n * math.ceil(k/n)) - k
    if overshootsby < 1:
        return widthsList
    elif overshootsby < len(widthsList):
        for i in range(len(widthsList)):
            if overshootsby > 0:
                widthsList[i] -= 1
                overshootsby -= 1
        return widthsList
    else:
        print("""
     ______
    /      \
    | | |  |
    |  /   |
    |  _   |
    \_____/



    """)
            

                
        
    

def mergeCellsRequest(startRowIn,endRow,startColumn,endColumn,sheetID,name,rgb,fontSize=18):
    startRow = startRowIn
    if startRowIn < 10:
        startRow = 10
    request = [
              {
                  "updateCells": {
                      "rows": {
                          "values": {
                              "userEnteredValue": {
                                  "stringValue": name
                              },
                          },
                          
                      },
                      "fields": "userEnteredValue",

                      "range": {
                          "sheetId": 0,
                          "startRowIndex": int(startRow)-1,
                          "endRowIndex": int(endRow)-1,
                          "startColumnIndex": int(startColumn),
                          "endColumnIndex": int(endColumn+1)
                      }
                    }
              },
              {
                  "mergeCells": {
                      "range": {
                          "sheetId": 0,
                          "startRowIndex": int(startRow)-1,
                          "endRowIndex": int(endRow)-1,
                          "startColumnIndex": int(startColumn),
                          "endColumnIndex": int(endColumn+1)
                          },
                      "mergeType": "MERGE_ALL"
                  }
              },
              {
                  "updateCells": {
                      "range": {
                            "sheetId": 0,
                            "startRowIndex": int(startRow)-1,
                            "endRowIndex": int(endRow)-1,
                            "startColumnIndex": int(startColumn),
                            "endColumnIndex": int(endColumn+1)
                            },
                      "fields": "userEnteredFormat"
                  }
              },
              {
                  "updateBorders": {
                      "range": {
                          "sheetId": '0',
                          "startRowIndex": int(startRow)-1,
                          "endRowIndex": int(endRow)-1,
                          "startColumnIndex": int(startColumn),
                          "endColumnIndex": int(endColumn+1)
                      },
                      "top": {
                          "style": "SOLID",
                          "width": 2,
                      },
                      "bottom": {
                          "style": "SOLID",
                          "width": 2,
                      },
                      "left": {
                          "style": "SOLID",
                          "width": 2,
                      },
                      "right": {
                          "style": "SOLID",
                          "width": 2,
                      }
            
                   }
                },
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": 0,
                            "startRowIndex": int(startRow)-1,
                            "endRowIndex": int(endRow)-1,
                            "startColumnIndex": int(startColumn),
                            "endColumnIndex": int(endColumn+1)
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": {
                                    "red": rgb[0],
                                    "green": rgb[1],
                                    "blue": rgb[2]
                                },
                                "verticalAlignment": "MIDDLE",
                                "horizontalAlignment" : "CENTER",
                                "wrapStrategy" : "WRAP",
                                "textFormat": {
                                "foregroundColor": {
                                    "red": 0,
                                    "green": 0,
                                    "blue": 0
                                    },
                                "fontSize": fontSize,
                                "bold": True
                                }
                            }
                        },
                        "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)"
                    }
                },
              ]

    return request



def clearCellsRequest(startRow,endRow,startCol,endCol):
    request = [{
                  "updateCells": {
                      "range": {
                            "sheetId": 0,
                            "startRowIndex": int(startRow)-1,
                            "endRowIndex": int(endRow)-1,
                            "startColumnIndex": int(startCol),
                            "endColumnIndex": int(endCol+1)
                            },
                      "fields": "userEnteredValue"
                  }
              },
               {
                  "updateCells": {
                      "range": {
                            "sheetId": 0,
                            "startRowIndex": int(startRow)-1,
                            "endRowIndex": int(endRow)-1,
                            "startColumnIndex": int(startCol),
                            "endColumnIndex": int(endCol+1)
                            },
                      "fields": "userEnteredFormat"
                  }
              },
               {
                  "unmergeCells": {
                      "range": {
                          "sheetId": 0,
                          "startRowIndex": int(startRow)-1,
                          "endRowIndex": int(endRow)-1,
                          "startColumnIndex": int(startCol),
                          "endColumnIndex": int(endCol+1)
                          }
                  }
              }]
    return request
    

def A1Notation(x,y):
    alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    colNameList = [i for i in alphabet] + [a + b for a in alphabet for b in alphabet]
    return colNameList[x]+str(y)

def addBlock(divisionIn,startDT,endDT,classname,classtype,spreadsheet,sheetID, width = 3, offset = 0,debug = False):

    #Convert DT start and end into a cell start point and cell end point
    fifteenminutes = datetime.timedelta(0,900)
    blocklength = (datetime.datetime.combine(datetime.date.min, endDT.time()) - datetime.datetime.combine(datetime.date.min, startDT.time())) / fifteenminutes
   # print(blocklength)
    startCol = offset + 1 + startDT.weekday()*19 #Monday 0, Sunday 6

    startCol += 3 * (['Div1X','Div1Y','Div2','Div3X','Div3Y','Div4'].index(divisionIn))
    endCol = startCol + width - 1
    startDT = startDT.time()
    startrow = -54 + (((datetime.datetime.combine(datetime.date.min, startDT) - (datetime.datetime.combine(datetime.date.min, datetime.time(8,45)))) / fifteenminutes) * 4)
    #print(startrow)
    endrow = startrow + (blocklength * 4)
    #if startrow < 10:
    #    startrow = 10
    if debug:
        print(startrow)
        print(endrow)
        print(startCol)
        print(endCol)

    #Make color background key
    newclasstype = str(classtype)
    if blocklength > 3 and classtype == 'elc':
        newclasstype = 'art'
    if 'SREPT' in classname:  #The distinction between srp and elc only matters here, only for color coding
        newclasstype = 'srp'

    colors = {
        'hum': [210,190,220],
        'mst': [115,145,200],
        'esp': [255,155,155],
        'frn': [250,220,220],
        'lun': [245,245,180],
        'art': [225,240,200],
        'com': [200,200,175],
        'elc': [145,215,230],
        'srp': [255,185,110]
        }
    
    
    rgb = [i/255 for i in colors[newclasstype]]
    
    


    spreadrange = A1Notation(startCol,int(startrow)) + ':' + A1Notation(endCol,int(endrow))
    # Translate info into requests

    requests = []
    fontSizeSetting = 24
    if width < 3:
        fontSizeSetting = 18
    if width < 2:
        fontSizeSetting = 10


        
    requests += mergeCellsRequest(startrow,endrow,startCol,endCol,sheetID,classname,rgb,fontSizeSetting)

    thebody = {
            'requests': requests
           }
    return requests
    

def fifteenMinuteRounder(datetimeIn):
    datetimeOut = datetimeIn + datetime.timedelta(minutes = (int(datetimeIn.minute/15) * 15) - datetimeIn.minute)
    return datetimeOut
        
    
    
    

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly','https://www.googleapis.com/auth/spreadsheets']



creds = None
if os.path.exists('token.json'):
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    
# If there are no (valid) credentials available, let the user log in.
# The only "user" logging into this should be schedulemeridianacademy@gmail.com, and that account only needs to log in once
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open('token.json', 'w') as token:
        token.write(creds.to_json())


CalendarIds = {'Div1X':'c_0p4e4hqdsimok0idv3n3o7o4e4@group.calendar.google.com','Div1Y':'c_tv2ued7u07pr46maqqhup6fmio@group.calendar.google.com',
               'Div2':'c_vecnd9qg4vgfb7dlcg9d5qbcuc@group.calendar.google.com','Div3X':'c_fphvrf8m2c9mektsis4uunbne4@group.calendar.google.com',
               'Div3Y':'c_ihj68ne0m1o5b7ac00acqgd8fo@group.calendar.google.com','Div4':'c_jfmdg03l5dlnjslp8pkbcmbci0@group.calendar.google.com'}

classes = ['Div1X','Div1Y','Div2','Div3X','Div3Y','Div4']
totalClassList = []

service = build('calendar', 'v3', credentials=creds)




now = datetime.datetime.utcnow().isoformat() + 'Z'  # 'Z' indicates UTC time
fifteenminutes = datetime.timedelta(0,900)
for division in classes:
    today = datetime.date.today()
    if today.weekday() > 4:
        monday = today + datetime.timedelta(days=-today.weekday(), weeks=1)
    else:
        monday = today - datetime.timedelta(days=today.weekday())
    mondayStart = datetime.datetime.combine(monday, datetime.time(8,45))
    fridayEnd = mondayStart + datetime.timedelta(days=6)
    monday = mondayStart.isoformat() + 'Z'
    after = fridayEnd
    events_result = service.events().list(calendarId=CalendarIds[division], timeMin=monday,
                                              maxResults=250, singleEvents=True,
                                              orderBy='startTime').execute()
    


    events = events_result.get('items', [])



    for event in events:
        #print(dict(event).keys())
        #print(event['summary'])
        startTime = datetime.datetime.strptime(event['start']['dateTime'],'%Y-%m-%dT%H:%M:%SZ')
        if startTime > after:
            continue
        endTime = datetime.datetime.strptime(event['end']['dateTime'],'%Y-%m-%dT%H:%M:%SZ')
        #print(endTime)
        summary = event['summary']
        classType = classTypeOf(summary)
        #print(classType)
        #print('\n\n\n')
        totalClassList += [{'division':[division],'start':startTime,'end':endTime,'name':summary,'type':classType}]


fifteenminutes = datetime.timedelta(0,900)
service = build('sheets','v4',credentials=creds)


spreadsheet = service.spreadsheets()

spreadsheetID = '1uWZSUKB23TBHBIzKOFBVLSk6TIVnq1-t1E9RBsdVPVY'
middleSchool = ['Div1X','Div1Y','Div2']
highSchool = ['Div3X','Div3Y','Div4']
masterRequest = []

for i,item in enumerate(totalClassList):    #Deal with irregularities
    if item['name'] == 'Athlete Activism':
        totalClassList[i]['division'] = ['Div1X', 'Div1Y', 'Div2']
    afterschool = item['start'].time() >= datetime.time(19,15)
    if afterschool:
        totalClassList[i]['division'] = ['Div1X', 'Div1Y', 'Div2','Div3X','Div3Y','Div4']
    if (item['name'] == 'JRPS - BG' or item['name'] == 'AP Calc - JA' or item['name'] == 'College - SP') and (item['end'] - item['start'] < datetime.timedelta(minutes=90)):
        totalClassList[i]['division'] = ['Div3X','Div3Y','Div4']
        print('updated ' + str(item['start']))
    if item['name'] == 'MST Field Trip to Arboretum':
        totalClassList.pop(i)



indicesToRemove = []
for i,Class in enumerate(totalClassList):  #Remove duplicates
    if not i in indicesToRemove:
        for otheri,otherClass in enumerate(totalClassList):
            if [Class['start'],Class['name']] == [otherClass['start'],otherClass['name']] and i != otheri:
                indicesToRemove += [otheri]

                Class['division'] += [dontuseiforcomprehensions for dontuseiforcomprehensions in otherClass['division'] if dontuseiforcomprehensions not in Class['division']] #Check this line if code doesn't work
                Class['division'].sort()
        

indicesToRemove = [*set(indicesToRemove)]
indicesToRemove.sort(reverse = True)
for i in indicesToRemove:
    totalClassList.pop(i)
        



    

classClusters = []
clusteredIndices = []
for i,Class in enumerate(totalClassList):   #Make clusters
    if not i in clusteredIndices:
        
        classClusters += [[Class]]
        clusteredIndices += [i]
        for i2,Class2 in enumerate(totalClassList):
            if i2 == i or i2 in clusteredIndices:
                continue #Ok, what's the big idea, huh? You think you're so smart, don't you. Did you write this? I didn't think so.
            overlapping = Class['start'] >= Class2['start'] and Class['start'] < Class2['end'] or Class2['start'] >= Class['start'] and Class2['start'] < Class['end']
            
            if (Class['division'] == Class2['division']) and overlapping:
                clusteredIndices += [i2]
                classClusters[-1] += [Class2]


#   IMPORTANT FOR USE OF PROGRAM (if anyone besides me ever reads this lol. Josh? Emily?)
#
#   Meridian's schedule is erratic and inconsistent,
#   and trying to codify it is a nightmare. The code
#   that I already have lays out as many general rules 
#   as I could think of, but usually every week will have
#   one or two blocks that break the program. The "solution"
#   to this is to find which block is messsing up the 
#   program, and then delete that block and add it in
#   manually in Google Sheets. To do this, use a slice of
#   classClusters using the "for i in classClusters" line
#   below: "for i in classClusters[0:30]" then run the program.
#   If it gives you an error, make the slice smaller. Once you've
#   found the index of the slice, use the below two lines
#   (commented out) to remove that index from the list.
#   Then, remove the slice index from the "for i in classClusters"
#   line and run the program in full. Go to the Google Sheet
#   and edit it. You can use the Meridian Schedule account
#   to do it. The password is "M3r1d1@N" and the email is
#   "schedulemeridianacademy@gmail.com".
#
#
#   FAQ:
#
#   Q: Amos, this program sucks. Isn't the point of it that it's automated?
#   A: I'm busy applying to colleges. This program works well enough, and
#      you can program in new rules for how to handle odd blocks if you
#      really want to.
#
#   Q: The index that breaks the code seems to be a normal block, I don't
#      understand why it would break the program?
#   A: Try going down one index.
#
#   Q: This code is very imperfect. What patterns in particular have you
#      neglected to account for?
#   A:
#       1) Blocks with overlapping divisions that do not share all of their divisions
#          (like model UN lunch and any other lunch block)
#
#       2) Electives in a language/elective block section that are only for one division
#          (you could fix this one by deciding whether theyre MS or HS, then adding the
#           missing divisions)
#               i) I've added a special case for JRPS and AP Calc, but this is an incomplete solution.
#
#       3) Field trips (I don't think it's worth trying to fix this one, there are too many
#          issues. Surprisingly, they often work with no problems)
#
#       4) Blocks that split other blocks (Field trips or long blocks that go through what would normally be lunch blocks for divisions other than Div 4 and 1X)
#


print(classClusters[59])
classClusters.pop(59)

print(classClusters[62])
classClusters.pop(62)

print(classClusters[75])
classClusters.pop(75)



for i in classClusters:
    sections = len(i)
    columns = 3 * len(i[0]['division'])
    startDiv = min(i[0]['division'])

    widths = nIntoK(sections,columns)
    widths.sort(reverse=True)
    for index,classSection in enumerate(i):

        masterRequest += addBlock(startDiv,classSection['start'],classSection['end'],classSection['name'],classSection['type'],spreadsheet,spreadsheetID, widths[index], sum(widths[:index]))


clearRequest = []
clearRequest += clearCellsRequest(10,170,1,18)
clearRequest += clearCellsRequest(10,170,20,37)
clearRequest += clearCellsRequest(10,170,39,56)
clearRequest += clearCellsRequest(10,170,58,75)
clearRequest += clearCellsRequest(10,170,77,94)
                
            
thebody = {
    'requests' : clearRequest
    }

spreadsheet.batchUpdate(
                        spreadsheetId=spreadsheetID,
                        body=thebody).execute()


thebody = {
    'requests' : masterRequest
    }
            

    

spreadsheet.batchUpdate(
                        spreadsheetId=spreadsheetID,
                        body=thebody).execute()




