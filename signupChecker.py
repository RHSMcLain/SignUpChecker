import pyperclip, requests, bs4, json
import re


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
url = input("what is the url for your signup Genius?")
x = re.findall("/go/(.+)#", url)

postobj = '{"forSignUpView":true,"urlid":"' + x[0] + '","portalid":0}'
res = requests.post("https://www.signupgenius.com/SUGboxAPI.cfm?go=s.getSignupInfo", postobj)
# print(res.text)

results = json.loads(res.text)
participants = getParticipants(results)
schedule = getSlots(results, participants)
output = getScheduleString(schedule)
pyperclip.copy(output)
print("Results are ready to paste into your favorite spreadsheet")

# res.raise_for_status()

# soup = bs4.BeautifulSoup(res.text)
# print(soup)
# dv = soup.select('table')
# print(dv)



