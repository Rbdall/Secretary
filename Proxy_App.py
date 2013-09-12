'''
Created on Aug 11, 2013

@author: Ryan
'''
import gspread, getpass

user = "###"
attendanceURL = "###"
proxyURL = "###"

    
def sucessfulConnection(truth):
    global passwordEntered
    passwordEntered = truth
    global loginAchieved
    loginAchieved = truth

def getPassword():
    if passwordEntered == False:
        global password 
        password = getpass.getpass()

def attemptLogin():
    if loginAchieved == False:
        print("Accessing Google as user: " + user)
        getPassword()
        numIncorrect = 0
        try:
            if(numIncorrect == 2):
                print("Three invalid logins. Exiting")
                exit()
            attempt = gspread.login(user, password)
            sucessfulConnection(True)
            global login
            login = attempt
        except Exception:
            print("Incorrect password.")
            numIncorrect += 1
            getPassword()
    
def getAbsent():
    attemptLogin()
    sheet = login.open_by_url(attendanceURL)
    worksheet = sheet.get_worksheet(0)
    names = worksheet.col_values(1)[2:]
    attendance = worksheet.col_values(2)[2:]
    while attendance.__len__() < names.__len__():
        attendance.append(None)
    x = 0
    absentList = []
    for name in names:
        if(attendance[x] != None):
            absentList.append(name)
        x += 1 
    return absentList

def getProxyList():
    attemptLogin()
    sheet = login.open_by_url(proxyURL)
    worksheet = sheet.get_worksheet(0)
    names = worksheet.col_values(1)[2:]
    proxy1 = worksheet.col_values(2)[2:]
    proxy2 = worksheet.col_values(3)[2:]
    proxy3 = worksheet.col_values(4)[2:]
    proxyList = []
    x = 0
    for name in names:
        proxies = [proxy1[x], proxy2[x], proxy3[x]]
        entry = [name, proxies]
        proxyList.append(entry)
        x += 1
    return proxyList

def returnProxies(attendance, proxies):
    results = []
    for absent in attendance:
        for people in proxies:
            if(people[0] == absent):
                if((people[1])[0] not in attendance):
                    entry = [absent, (people[1])[0]]
                    results.append(entry)
                    break
                if((people[1])[1] not in attendance):
                    entry = [absent, (people[1])[1]]
                    results.append(entry)
                    break
                if((people[1])[2] not in attendance):
                    entry = [absent, (people[1])[2]]
                    results.append(entry)
                    break
                entry = [absent, "Nobody"]
                results.append(entry)
    for entries in results:
        print(entries[1] + " is a proxy for " + entries[0])
        
def runTool():
    sucessfulConnection(False)
    returnProxies(getAbsent(), getProxyList())       

if __name__ == '__main__':
    runTool()