from pip._vendor import requests
import os
import csv
import time
os.chdir(r"c:\Users\miria\Downloads")
#Dealing with CSV file called PopuliDonations
f = open('BloomerangExport.csv', 'r')
reader = csv.reader(f)
mylist = list(reader)
f.close()

alreadyHouseholds = []
row = 0
mylist = [["Sort Name","Recognition Name","His Mobile Phone Number","Her Mobile Phone Number","Home Phone Number","Primary Email Address", "Primary Street","Primary City","Primary State","Primary ZIP Code","Acct #"]]
upto = 0
while upto < 78:
    print(upto)
    response3 = requests.get('https://api.bloomerang.co/v2/households?skip='+ str(50*upto)+'&take=50',
    headers={
    'X-API-KEY': '***''
    },
    json={
    }
    )
    for person in (response3.json()["Results"]):
        added = False
        shouldadd = False
        sortName = recName = hiscell = hercell = homephone = email = street=city=state=zip=acct=""
        
        householdlist = []
        try:
            for p in person['MemberIds']:
                head = False
                response1 = requests.get('https://api.bloomerang.co/v2/constituent/' + str(p),
                headers={
                'X-API-KEY': '***''
                },
                json={
                }
                )
                for thing in response1.json()['CustomValues']:
                    if thing["FieldId"] == 18432:
                        for t in thing['Values']:
                            try:
                                if ((t['Id'] == 19472) or (t['Id'] == 19471)) and (response1.json()['PrimaryAddress']["State"]) == "Maryland":
                                    shouldadd = True
                                    print(response1.json())
                                    lastname = (response1.json()['LastName'])
                                    acct = response1.json()['AccountNumber']
                                    if (response1.json()['IsHeadOfHousehold']):
                                        head = True
                                    try:
                                        email =response1.json()['PrimaryEmail']['Value']
                                    except:
                                        pass
                            
                                    street = response1.json()['PrimaryAddress']["Street"]
                                    city = response1.json()['PrimaryAddress']["City"]
                                    state= response1.json()['PrimaryAddress']["State"]
                                    zip= response1.json()['PrimaryAddress']["PostalCode"]
                                    for phone in (response1.json()["PhoneIds"]):
                                        phoneresponse = requests.get('https://api.bloomerang.co/v2/phone/' + str(phone),
                                        headers={
                                        'X-API-KEY': '***''
                                        },
                                        json={
                                        }
                                        )
                                        if (phoneresponse.json()["Type"]) == "Home":
                                            homephone = (phoneresponse.json()["Number"])
                                            print(homephone)
                                        if ((phoneresponse.json()["Type"]) == "Mobile") and (head):
                                            print(hiscell)
                                            hiscell = (phoneresponse.json()["Number"])
                                        elif ((phoneresponse.json()["Type"]) == "Mobile"):
                                            print(hercell)
                                            hercell = (phoneresponse.json()["Number"])

                                    householdlist= [person['SortName'] ,person['RecognitionName'], hiscell ,hercell , homephone , email ,street,city,state,zip,acct]
                            except:
                                pass
        except:
            pass
        if added == False and shouldadd == True:
            print(householdlist)
            mylist.append(householdlist)
            added = True
                        
                        
                        
        
    
    file3 = open('BloomerangExport.csv', 'w', newline = '')
    csv_writer = csv.writer(file3)
    csv_writer.writerows(mylist)
    #print(mylist)
    file3.close()
    upto = upto+1

