import pyperclip, requests, bs4, json
import re
from openpyxl import Workbook


def getParticipants(results):
    participants = []
    output = "id|StartTime|Name|Comment|slot\n"
    for x, y in results["DATA"]["participants"].items():
        i = y[0]
        n = {}
        n['id'] = x
        n['starttime'] = i['starttime']
        n['firstname'] = i['firstname']
        n['lastname'] = i['lastname']
        n['comment'] = i['mycomment']
        n['slotid'] = i['slotitemID']
        participants.append(n)

        output +=  x + "|" + y[0]['starttime'] + "|" + y[0]['firstname'] + " " + y[0]['lastname'] + "|" + y[0]['mycomment'] + "|" + str(y[0]['slotitemID']) + "\n"
    return participants
    # pyperclip.copy(output)
    # print("all done")
    # print(output)
def getSlots(results, participants):
    schedule = []
    line = "date|time|slotitemid|itemid\n"
    for x, y in results["DATA"]["slots"].items():
       
      
        # print("-------------------------")
        # print(x, ":", y)
        thisDate = results["DATA"]["slots"][x]["starttime"] + "-" + results["DATA"]["slots"][x]["endtime"]
        for item in results["DATA"]["slots"][x]["items"]:
            slot = {}
            slot['range'] = thisDate
            slot['item'] = item['item']
            slot['slotitemid'] = item['slotitemid']
            slot['itemid'] = item['itemid']
            #TODO: searc through participants to get participant data
            slot['participant'] = None
            for p in participants:
                if (p['slotid'] == item['slotitemid']):
                    slot['participant'] = p
            schedule.append(slot)
            line += thisDate + "|" + item['item'] + "|" + str(item['slotitemid']) + "|" + str(item['itemid']) + "\n"
     
    
        # output +=  x + "|" + y[0]['starttime'] + "|" + y[0]['firstname'] + " " + y[0]['lastname'] + "|" + y[0]['mycomment'] + "\n"
    return schedule
def getScheduleString(schedule):
    output = "range\ttime\tfirstName\tlastName\tcomment\n"
    for slot in schedule:
        p = slot['participant']
        if (p == None):
            line = "\t".join([slot['range'], slot['item'], "--", "--","--"])
        else:
            line = "\t".join([slot['range'], slot['item'], p['firstname'], p['lastname'], p['comment']])
        output += line + "\n"
    return output
# link = input("What is the link to your signup genius page?")
# x = re.search('"urlid":"[.+"]')
# print(x)
def pushToSpreadsheet(schedule, workbookName, url=""):
    wb = Workbook()
    sheet = wb.active

    output = "range\ttime\tfirstName\tlastName\tcomment\n"
    sheet['A1'] = "Signup Genius UrL: " 
    sheet['B1'] = url
    sheet['A2'] = "Range"
    sheet["B2"] = "Time"
    sheet["C2"] = "First Name"
    sheet["D2"] = "Last Name"
    sheet["E2"] = "Comment"
    counter = 3
    for slot in schedule:
        p = slot['participant']
        sheet["A" + str(counter)] = slot['range']
        sheet["B" + str(counter)] = slot['item']
        if (p == None):
            sheet['C' + str(counter)] = "--"
            sheet['D' + str(counter)] = "--"
            sheet['E' + str(counter)] = "--"
            
        else:
            sheet['C' + str(counter)] = p['firstname']
            sheet['D' + str(counter)] = p['lastname']
            sheet['E' + str(counter)] = p['comment']
        counter += 1
    wb.save(workbookName + ".xlsx") 
def getSignups(url, teacherName, createSheets=True):
    print(url)
    print("------")
    x = re.findall("/go/(.+)#?", url.strip())
    print(x)

    postobj = '{"forSignUpView":true,"urlid":"' + x[0] + '","portalid":0}'
    res = requests.post("https://www.signupgenius.com/SUGboxAPI.cfm?go=s.getSignupInfo", postobj)
    # print(res.text)

    results = json.loads(res.text)
    participants = getParticipants(results)
    schedule = getSlots(results, participants)
    if createSheets:
        pushToSpreadsheet(schedule, teacherName, url)
    else:
        pyperclip.copy(getScheduleString(schedule))
        print("Your results are ready to paste into your favorite spreadsheet")

def pullMultiplesFromClipboard():
    urls = pyperclip.paste()
    print(urls)
    print("test")
    lines = urls.split("\n")
    for line in lines:
        item = line.split("\t")
        if (len(item) == 2 and len(item[1])>5):
            getSignups(item[1], item[0].strip())           


print("Welcome to the Sign Up Checker")
print("This program will pull sign ups from Sign Up Genius and put them in a spreadsheet")
print("--------------------------------------------------------------------------")
print("I can do multiples. Copy from a spreadsheet with columns (Teacher Name) and (SignUp Link)")
url = input("what is the url for your signup Genius? or (C) if it's multiples in the clipborad ")

if (url == "C" or url == "c"):
    pullMultiplesFromClipboard()
else:
    getSignups(url, "", False)
exit(0)
# url = input("what is the url for your signup Genius?")
# x = re.findall("/go/(.+)", url)
# print(x)
# postobj = '{"forSignUpView":true,"urlid":"' + x[0] + '","portalid":0}'
# res = requests.post("https://www.signupgenius.com/SUGboxAPI.cfm?go=s.getSignupInfo", postobj)
# # print(res.text)

# results = json.loads(res.text)
# participants = getParticipants(results)
# schedule = getSlots(results, participants)
# output = getScheduleString(schedule)
# pyperclip.copy(output)
# print("Results are ready to paste into your favorite spreadsheet")
# pushToSpreadsheet(schedule, "McLain")


# res.raise_for_status()

# soup = bs4.BeautifulSoup(res.text)
# print(soup)
# dv = soup.select('table')
# print(dv)



