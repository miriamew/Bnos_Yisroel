from pip._vendor import requests
import os
import csv
os.chdir(r"c:\Downloads")
f = open('BloomerangExport.csv', 'r')
reader = csv.reader(f)
mylist = list(reader)
f.close()'''

alreadyHouseholds = []
row = 0
mylist = [["Name","Last Name","His Cell Phone","Her Cell Phone","Home Phone", "His Email","Her Email","Address"]]
upto = 0
while upto < 1:
    response3 = requests.get('https://api.bloomerang.co/v2/households?skip='+ str(50*upto)+'&take=50',
    headers={
    'X-API-KEY': '***'c'
    },
    json={
        
    }
    )

    for person in (response3.json()["Results"]):
        lastname = hiscell = hercell = homephone = hisemail = heremail = address = ""
        householdlist = []
        
        #print(person)
        for p in person['MemberIds']:
            head = False
            response1 = requests.get('https://api.bloomerang.co/v2/constituent/' + str(p),
            headers={
            'X-API-KEY': '***''
            },
            json={
                
            }
            )
            lastname = (response1.json()['LastName'])
            print(response1.json())
            if (response1.json()['IsHeadOfHousehold']):
                head = True
            try:
                address = (response1.json()['PrimaryAddress']["Street"] + " "+response1.json()['PrimaryAddress']["City"]+" "+response1.json()['PrimaryAddress']["State"])
                if (head):
                    hisemail = (response1.json()['PrimaryEmail']["Value"])
                else:
                    heremail = (response1.json()['PrimaryEmail']["Value"])
                #print(response1.json()['PrimaryPhone']["Number"])
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
                    if ((phoneresponse.json()["Type"]) == "Mobile") and (head):
                        hiscell = (phoneresponse.json()["Number"])
                    elif ((phoneresponse.json()["Type"]) == "Mobile"):
                        hercell = (phoneresponse.json()["Number"])
        

                '''for phone in (response1.json()["EmailIds"]):
                    phoneresponse = requests.get('https://api.bloomerang.co/v2/email/' + str(phone),
                    headers={
                    'X-API-KEY': '***''
                    },
                    json={
                        
                    }
                    )
                    print(phoneresponse.json())'''
                '''for phone in (response1.json()["AddressIds"]):
                    phoneresponse = requests.get('https://api.bloomerang.co/v2/address/' + str(phone),
                    headers={
                    'X-API-KEY': '***''
                    },
                    json={
                        
                    }
                    )
                    print(phoneresponse.json())'''
            except:
                pass
        householdlist= [person['RecognitionName'], lastname, hiscell, hercell, homephone, hisemail, heremail, address]
        print(mylist)
        mylist.append(householdlist)
        
        
    
    '''file3 = open('BloomerangExport.csv', 'w', newline = '')
    csv_writer = csv.writer(file3)
    csv_writer.writerows(mylist)
    #print(mylist)
    file3.close()
    upto = upto+1'''

