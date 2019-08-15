import sys 
import redcap
import pandas as pd
import numpy as np
import re 

def call_record(i): #function that prints out record ID
    print('record:', i)

if __name__=='__main__': #takes in one argument in the command prompt where I declare the record ID
    i = str(sys.argv[1])
    call_record(i)

wcmAffiliates = ['@med.cornell.edu', '@myhunter.cuny.edu', '@cornell.edu', '@hss.edu', '@hunter.cuny.edu', '@HSS.edu', '@HSS.EDU', '@mskcc.org', '@amcny.org', '@nyp.org', '@NYP.org', '@burke.org']
personalemails= ['@yahoo.com', '@gmail.com', '@hotmail.com', '@mail.com', '@aol.com', '@icloud.com', '@me.com', '@mac.com'] #a list used to check if users provided personal emails
project = redcap.Project('https://redcap.ctsc.weill.cornell.edu/redcap_protocols/api/', 'C0F7DAD56835B9DAA53CAC272A060E4F')

num = project.export_records(records=[i], fields=['numberofusers'])
numofusers = int(num[0]['numberofusers'])

def User(record):
    username  = project.export_records(records=[i], fields=['email'], format='df')
    email = username['email'].to_string()
    domain = re.search("@[\w.]+", email).group()
    return domain

def Sponsor(record):
    sponsor = project.export_records(records=[i], fields=['sponsoremail'])
    string = sponsor[0]['sponsoremail']
    semail = re.search("@[\w.]+", string).group()
    return semail

def SponsorName(record):
    sponsor = project.export_records(records=[i], fields=['sponsoremail'])
    string = sponsor[0]['sponsoremail']
    semail = re.search("@[\w.]+", string).group()
    if semail=="@med.cornell.edu" or semail=="@nyp.org":
        semail2 = string.split("@")[0]
    else: 
        semail2 = string 
    return semail2

if (User(i) != "@med.cornell.edu" and User(i) != "@nyp.org") and User(i) not in personalemails and Sponsor(i) in wcmAffiliates:
    username  = project.export_records(records=[i], fields=['email'], format='df')
    username.columns= ['Username']
    user = project.export_records(records=[i], fields=['firstname', 'lastname', 'email', 'sponsoremail', 'userexpiration', 'user_notes'], format = 'df')
    user.insert(loc=3, column='Institution ID', value=np.nan)
    user.columns = ['First Name', 'Last Name', 'Email address', 'Institution ID','Sponsor username', 'Expiration', 'Comments']
    user.loc[user['Sponsor username'] != '', 'Sponsor username'] = SponsorName(i)
    result = pd.concat([username, user], axis=1)
    result.to_csv(r'Cats2Template.csv', index=False)
else:
    print("ERROR! Sponsor's email is not part of WCM affiliates: {0} and/or \n username's email is WCM/NYP/personal email: {1}".format(Sponsor(i), User(i)))

def Users(record):
    if numofusers > 1: 
        for j in range(2, numofusers+1, 1):
            j = str(j)
            email = 'email_'
            fname = 'firstname_'
            lname = 'lastname_'
            userexpiration= 'userexpiration_'
            username = project.export_records(records=[i], fields=[email+j], format='df')
            string = username[email+j]
            string = string.to_string()
            domain = re.search("@[\w.]+", string).group()
            return domain

if numofusers > 1: 
    for j in range(2, numofusers+1, 1):
        j = str(j)
        email = 'email_'
        fname = 'firstname_'
        lname = 'lastname_'
        username = project.export_records(records=[i], fields=[email+j], format='df')
        userexpiration= 'userexpiration_'
        if (Users(i) != "@med.cornell.edu" and Users(i) != "@nyp.org") and User(i) not in personalemails and Sponsor(i) in wcmAffiliates:
            username.columns= ['Username']
            user = project.export_records(records=[i], fields=[fname+j, lname+j, email+j,'sponsoremail', userexpiration+j, 'user_notes'], format='df')
            user.insert(loc=3, column='Institution ID', value=np.nan)
            user.loc[user['sponsoremail'] != '', 'sponsoremail'] = SponsorName(i)
            result = pd.concat([username, user], axis=1)
            with open('Cats2Template.csv', 'a') as fd:
                result.to_csv(fd, header=False, index=False)
        else: 
            print("ERROR! Sponsor's email is not part of WCM affiliates: {0} and/or \n username's email is WCM/NYP/personal email: {1}".format(Sponsor(i), Users(i)))