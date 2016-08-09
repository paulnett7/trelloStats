import requests
import smtplib
import json
import re
import datetime
import dateutil.parser
import sys
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

#set variables from command line
if len(sys.argv) != 2:
    print('Please input password for email.')
    sys.exit()
password = sys.argv[1]

#Paul Nettleton's API Key and Authorization token
apiKey = 'a22c6d69c7c3264f73545cd46b8a83e6'
authToken = '37779fcdad1f60c10f028e8ffb4a098f38ed4379552eb9cf07f247641b72b726'

#Address for each Trello board
trelloBoards = {
    'ActionItems' : '4fca5759faccb1f81a2d42cc',
    'TDPrinter' : '4fce0e04083b7943044e1ecc',
    'Projects' : '4fce0de5083b7943044e1342',
    'PCB' : '50b91b7596fe70950100fff4'
    }

trelloItems = {
    'board' : 'board',
    'card' : 'card'
    }

#Base HTML for email
email_form = """\
    <html>
    <head></head>
    <body style="font-size: 100%; font-family: Helvetica, Verdana, sans-serif; color: #404040; line-height: 144%; width: 600px; padding: 1em; ">
        <h1 style="color: #2387DC; border-bottom: 1px solid #D6DFE6; padding: 20px 0; text-align: center">think[box] weekly awards</h1>
        <div style="overflow: auto; width: 100%; padding: 1em 0em; border-bottom: 1px solid #D6DFE6;">
	    <img src="http://engineering.case.edu/thinkbox/sites/engineering.case.edu.thinkbox/files/images/busiest-bee.png" alt="Busiest Bee" height="125" width="125" style="float: left; margin: 0; padding: 0; height: 125px; width: 125px;" />
	    <ol style="list-style-position: inside; margin: 1em 0 0 0; padding: 0; width: 450px; float: right;">
	       <span style="color: #086EC3"><strong>Busiest Bees</strong> (most completed action items)</span>
           {0}
	    </ol>
        </div>
        <div style="overflow: auto; width: 100%; padding: 1em 0em; border-bottom: 1px solid #D6DFE6;">
	       <img src="http://engineering.case.edu/thinkbox/sites/engineering.case.edu.thinkbox/files/images/3d-royalty.png" alt="3D Royalty" height="125" width="125" style="float: left; margin: 0; padding: 0; height: 125px; width: 125px;" />
           <ol style="list-style-position: inside; margin: 1em 0 0 0; padding: 0; width: 450px; float: right;">
                <span style="color: #086EC3"><strong>3D Printer Royalty</strong> (handled the most 3D Printing jobs)</span>
                {1}
           </ol>
        </div>
        <div style="overflow: auto; width: 100%; padding: 1em 0em; border-bottom: 1px solid #D6DFE6;">
	        <img src="http://engineering.case.edu/thinkbox/sites/engineering.case.edu.thinkbox/files/images/moldiest-card.png" alt="Moldiest Card Completed" height="125" width="125" style="float: left; margin: 0; padding: 0; height: 125px; width: 125px;" />
	        <ul style="list-style-position: inside; margin: 1em 0 0 0; padding: 0; width: 450px; float: right;">
		        <span style="color: #086EC3"><strong>Moldiest Card Completed</strong> (completed oldest card)</span>
		        <li>{2}</li>
	        </ul>
        </div>
        <div style="overflow: auto; width: 100%; padding: 1em 0em;">
	        <img src="http://engineering.case.edu/thinkbox/sites/engineering.case.edu.thinkbox/files/images/project-completed.png" alt="Moldiest Card Completed" height="125" width="125" style="float: left; margin: 0; padding: 0; height: 125px; width: 125px;" />
	        <ul style="list-style-position: inside; margin: 1em 0 0 0; padding: 0; width: 450px; float: right;">
		        <span style="color: #086EC3"><strong>Projects Completed</strong> (this week\'s activity)</span>
		        {3}
	        </ul>
        </div>
    </body>
    </html>
    """
#Templates in case nothing interesting is happening on Trello
itemTemplate = '<li>{0}</li>'
noBees = '<li>The bees were sleeping this week. Let\'s do better next time!</li>'
noTDP = '<li>Nothing was happening in the 3D printers this week. :(</li>'
noProjects = '<li>No projects completed this week - what gives? Next week for sure.</li>'

#sends email using Paul Nettleton's CWRU email
def mailMe(emailBody):
    server = smtplib.SMTP('smtp.cwru.edu', 25)
    server.ehlo()
    server.starttls()
    server.ehlo()
    server.login('prn15', password)

    sender = 'prn15@case.edu'
    recipient = 'prn15@case.edu'
    msg = MIMEMultipart()
    msg['Subject'] = 'Weekly Trello Awards'
    msg['From'] = sender
    msg['To'] = recipient

    msg.attach(MIMEText(emailBody, 'html'))
    server.sendmail(sender, recipient, msg.as_string())
    server.quit()
    return

#sends a REST GET request to Trello, returns JSON object of response
def callTrello(trelloItem, ID, args):
    url = 'https://api.trello.com/1/' + trelloItem + '/' + ID + '/' + args
    url += '&key=' + apiKey
    url += '&token=' + authToken

    response = requests.get(url)
    return json.loads(response.text)

#tally's the amount each person has done
def tallyMembers(actions):
    memberTally = {}
    for i in range(len(actions)):
        try:
            memberTally[actions[i]['memberCreator']['fullName']] += 1
        except KeyError:
            memberTally[actions[i]['memberCreator']['fullName']] = 1
    memberArray = []
    for mem in memberTally:
        memberArray.append([mem, memberTally[mem]])
    return memberArray

#returns the top people
def getTop(actions):
    memberTallies = tallyMembers(actions)
    memberTallies = sorted(memberTallies, key=lambda mem: mem[1])
    memberTallies.reverse()
    return memberTallies

#removes the full times from the top processed cards
def removeFullTimes(members):
    fullTimes = ["Ian Charnas", "Benjamin Guengerich", "Ruth D'Emilia", "Raymond Krajci", "Marcus Brathwaite"]
    removeIndices = []

    i = 0
    while i < len(members):
        record = members[i]
        for person in fullTimes:
            if record[0] == person:
                members.remove(record)
                i -= 1
        i += 1
    return members

#formats the stuff to put in the email
def formatString(*args):
    digits = re.compile('{(\d+)}')

    strList = []
    searchString = args[0]
    while digits.search(searchString):
        index = digits.search(searchString).end()
        strList.append(searchString[:index])
        searchString = searchString[index:]

    finalStr = ''
    i = 1
    for string in strList:
        try:
            finalStr += digits.sub(args[i], string) #string.replace(digits, args[i])
        except IndexError:
            finalStr += string
        i += 1
    finalStr += searchString
    return finalStr

def formatItems(itemList, numItems, defaultResponse):
    if len(itemList) == 0:
        return defaultResponse
    returnList = ''

    i = 0
    for item in itemList:
        returnList += formatString(itemTemplate, str(item[0]) + " - " + str(item[1]))
        i += 1
        if numItems != -1 and i >= numItems:
            break
    return returnList

#returns a date a week in the past
def getWeekAgo():
    return datetime.datetime.combine(datetime.date.fromordinal(datetime.date.today().toordinal() - 7), datetime.datetime.utcnow().time())

#returns a date a month in the past
def getMonthAgo():
    monthAgo = datetime.datetime.today()
    try:
        monthAgo = monthAgo.replace(month= monthAgo.month - 1)
    except ValueError:
        monthAgo = monthAgo.replace(month=12, year= monthAgo.year - 1)
    return monthAgo

#returns the top 3D printers
def getTopTD():
    trello3DArgs = "?actions=all" + "&actions_since=" + getWeekAgo().isoformat()[:-3] + 'Z' + "&action_fields=idMemberCreator" + "&actions_limit=1000"
    threeDActions = callTrello(trelloItems['board'], trelloBoards['TDPrinter'], trello3DArgs)
    topMems = getTop(threeDActions['actions'])
    return topMems

#returns top Action Items people
def getTopActions():
    trelloActionArgs = "?actions=updateCard:closed" + "&actions_since=" + getWeekAgo().isoformat()[:-3] + 'Z' + "&actions_limit=1000" + "&action_fields=idMemberCreator"
    actionItems = callTrello(trelloItems['board'], trelloBoards['ActionItems'], trelloActionArgs)
    topMems = getTop(actionItems['actions'])
    return topMems

#returns the oldest card
def getMoldiest(actions):
    trelloMoldyCardArgs = "?actions=createCard" + "&action_fields=all" + "&members=true" + "&member_fields=fullName"
    moldiestCard = callTrello(trelloItems['card'], actions[0]['data']['card']['id'], trelloMoldyCardArgs)
    i = 1
    while i < len(actions):
        currentCard = callTrello(trelloItems['card'], actions[i]['data']['card']['id'], trelloMoldyCardArgs)
        if currentCard['closed'] == False:
            i += 1
            continue
        mDate = dateutil.parser.parse(moldiestCard['actions'][0]['date'])
        cDate = dateutil.parser.parse(moldiestCard['actions'][0]['date'])
        if cDate < mDate:
            moldiestCard = currentCard
        i += 1
    return moldiestCard

#returns oldest card in string format
def getMoldiestCard():
    trelloMoldyArgs = "?actions=updateCard:closed" + "&actions_since=" + getWeekAgo().isoformat()[:-3] + 'Z' + "&actions_limit=1000" + "&action_fields=data, idMemberCreator"
    closedCards = callTrello(trelloItems['board'], trelloBoards['ActionItems'], trelloMoldyArgs)
    moldiestCard = getMoldiest(closedCards['actions'])
    memberString = ''
    for i in range(len(moldiestCard['members'])):
        if i == 0: memberString = moldiestCard['members'][i]['fullName']
        else: memberString += ',' +  moldiestCard['members'][i]['fullName']
    card = [moldiestCard['name'], memberString]
    return card

def getCompletedProjects():
    trelloProjectArgs = "?actions=updateCard:closed" + "&actions_since=" + getWeekAgo().isoformat()[:-3] + 'Z' + "&actions_limit=1000" + "&action_fields=data,idMemberCreator"
    projectActions = callTrello(trelloItems['board'], trelloBoards['Projects'], trelloProjectArgs)['actions']
    projectNames = []

    for i in range(len(projectActions)):
        projectNames.extend([projectActions[i]['data']['card']['name'], projectActions[i]['memberCreator']['fullName']])
    return projectNames


#Main Program
actionTop = getTopActions()
TDPTop = getTopTD()
moldiestCard = getMoldiestCard()
compProjects = getCompletedProjects()

actionTop = removeFullTimes(actionTop)
TDPTop = removeFullTimes(TDPTop)

busyList = formatItems(actionTop, 3, noBees)
tdpList = formatItems(TDPTop, 3, noTDP)
projectList = formatItems(compProjects, -1, noProjects)

emailBody = formatString(email_form, busyList, tdpList, moldiestCard[0] + ' - ' + moldiestCard[1], projectList)
mailMe(emailBody)
