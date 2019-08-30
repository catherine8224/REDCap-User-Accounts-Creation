#-----------------------------------------------------------------------------------MADE BY CATHERINE EVA NG---------------------------------------------------------------------
#------------This will first take the data saved in csv file and upload it on REDCap through "Create users (bulk upload)". Then it will fill out the comments field of the Username Request Forms project (notified sponsor, completed, etc.) in the Username Request Form. It will then email the sponsors about the status of the user accounts--------------

import sys 
import redcap
import pandas as pd
import numpy as np
import re 
from selenium import webdriver
import base64
import csv 
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

def call_record(i): #function that prints out record ID
    print('record:', i)

if __name__=='__main__': #takes in one argument in the command prompt where I declare the record ID
    i = str(sys.argv[1])
    call_record(i)

institAffiliates = [domain names of email addresses] #a list used to check if sponsors are part of your institution, if not, email them. This is only if the institution may use a special username.

personalemails= ['yahoo.com', 'gmail.com', 'hotmail.com', 'mail.com', 'aol.com', 'icloud.com', 'me.com', 'mac.com'] #a list used to check if users provided personal emails

project = redcap.Project('[API URL of the REDCap site]', [API_key]) #You can find the API url from API Playground, which gives you examples of how to use REDCap API. Never share your API key with anyone. 

user = (project.export_records(records=[i], fields=['email', 'firstname', 'lastname'])) #finds the user's email, first and last name. Whatever you variable name is for those fields!
email = user[0]['email'] #finds email address of all users' emails
email = email.split('@') #splits user emails where @ symbol is 

num = project.export_records(records=[i], fields=['numberofusers'])
numofusers = int(num[0]['numberofusers'])

def User(record): #finds the domain name of the email of the first user
    username  = project.export_records(records=[i], fields=['email'], format='df')
    email = username['email'].to_string()
    domain = re.search("@[\w.]+", email).group()
    return domain

def Sponsor(record): #parses out the sponsor's name from a separate institution
    sponsor = project.export_records(records=[i], fields=['sponsoremail'])
    string = sponsor[0]['sponsoremail']
    semail = string.split('@')[1]
    return semail

def SponsorName(record): #parses out the sponsor's local-part of the email address of the institution
    sponsor = project.export_records(records=[i], fields=['sponsoremail']) #this is because in our institution there is a specific id that is used in the first part of the email address
    string = sponsor[0]['sponsoremail']
    semail = re.search("@[\w.]+", string).group()
    if semail=="[domain name of email address]" or semail=="[domain name of email address]":
        semail2 = string.split("@")[0]
    else: 
        semail2 = string
    return semail2

#CREATE USER ACCOUNTS IN CONTROL CENTER OF REDCAP AND FILL IN INFORMATION OF STATUS OF USER ACCOUNTS IN THE USERNAME REQUEST FORM
if (User(i) != "[domain name of email address]" and User(i) != "[domain name of email address]") and User(i) not in personalemails and Sponsor(i) in institAffiliates:
    username  = project.export_records(records=[i], fields=['email'], format='df')
    username.columns= ['Username']
    user = project.export_records(records=[i], fields=['firstname', 'lastname', 'email', 'sponsoremail', 'userexpiration', 'user_notes'], format = 'df') #or whatever your variable names for those specific things are
    user.insert(loc=3, column='Institution ID', value=np.nan)
    user.columns = ['First Name', 'Last Name', 'Email address', 'Institution ID','Sponsor username', 'Expiration', 'Comments']
    result = pd.concat([username, user], axis=1)
    result.to_csv(r'Template.csv', index=False)
    user = "[REDCap username]"
    pwd = base64.b64decode(b"[REDCap associated password]")
    pwd = pwd.decode("utf-8")
    driver = webdriver.Chrome('system path to your chromedriver.exe') #Create a automated test webpage e.g. C:\\Users\\cath\\Documents\\chromedriver_win32\\chromedriver.exe
    driver.get("[link to go to add users as a bulk upload]") #Enter REDCap link to upload new REDCap users
    elem = driver.find_element_by_id("username")
    elem.send_keys(user)
    elem = driver.find_element_by_id("password")
    elem.send_keys(pwd)
    driver.find_element_by_id("login_btn").click() #Log into REDCap
    submit = driver.find_element_by_name("fname").send_keys("[path to your csv file]") #path to your file.csv e.g. C:\\Users\\cath\\Documents\\REDCap Automated Email\\CatTemplate.csv
    f = open('Template.csv')
    csv_f = csv.reader(f)
    next(csv_f)
    for csv in csv_f:
         domain = csv[0].split('@')
         print(domain[1])
         if domain[1] not in institAffiliates: 
             driver.find_element_by_name("allow_create_db").click()
    driver.find_element_by_name("submit").click() #Submits the user accounts to be made as REDCap acocunts
    if driver.find_elements_by_xpath("//div[@class='red']/img"): #this will look at the result page after submitting the new user accounts
        driver.get("[Finds the link to the project with the id parameters to be the one as called by the system arguments, e.g. first part of link + i + last part of link]")
        radio = driver.find_element_by_id("label-status-1")
        radio.click()
        if driver.find_element_by_id("label-status_notifysp-1"):
            radio2 = 'Element exists'
        if radio2 == 'Element exists': 
            driver.find_element_by_id("label-status_notifysp-1").click()
        complete = driver.find_element_by_xpath("//select[@name='request_status_complete']/option[text()='Complete']")
        complete.click()
        driver.find_element_by_id("comments").send_keys("User already has account - emailed user and sponsor")
        submit = driver.find_element_by_id("submit-btn-saverecord")
        submit.click()
    else: 
        driver.get("[Finds the link to the project with the id parameters to be the one as called by the system arguments, e.g. first part of link + i + last part of link]")
        radio = driver.find_element_by_id("label-status-1")
        radio.click()
        radio2 = driver.find_element_by_id("label-status_notifysp-1")
        radio2.click()
        complete = driver.find_element_by_xpath("//select[@name='request_status_complete']/option[text()='Complete']")
        complete.click()
        submit = driver.find_element_by_id("submit-btn-saverecord")
        submit.click()
elif email[1] == "[domain name of institutional email address]" or email[1] == "[domain name of institutional email address]":
    print("ERROR! Sponsor's email is not part of the institutional affiliates: {0} and/or \n username's email is [Institution] /personal email: {1}".format(Sponsor(i), User(i)))
    #this will fill in the comments box of the Username Request Form if the email is part of the institution--meaning they can already log into REDCap. No need to request to be a REDCap user.
    driver.get("[Finds the link to the project with the id parameters to be the one as called by the system arguments, e.g. first part of link + i + last part of link]")
    radio = driver.find_element_by_id("label-status-1")
    radio.click()
    if driver.find_element_by_id("label-status_notifysp-1"):
        radio2 = 'Element exists'
    if radio2 == 'Element exists': 
        driver.find_element_by_id("label-status_notifysp-1").click()
    complete = driver.find_element_by_xpath("//select[@name='request_status_complete']/option[text()='Complete']")
    complete.click()
    driver.find_element_by_id("comments").send_keys("User has institutional ID - emailed user/sponsor ") #The text to fill into the comments box of the User Request Form
    submit = driver.find_element_by_id("submit-btn-saverecord")
    submit.click()
elif Sponsor(i) not in institAffiliates:
    print("ERROR! Sponsor's email is not part of [institutional] affiliates: {0}".format(Sponsor(i)))
    #this will fill in the comments box of the Username Request Form if the sponsor is not from the institution
    driver.get("https://redcap.edu/redcap_protocols/redcap_v8.10.20/DataEntry/index.php?pid=3122&id=" + i + "&page=request_status&event_id=15049&instance=1")
    radio = driver.find_element_by_id("label-status-1")
    radio.click()
    if driver.find_element_by_id("label-status_notifysp-1"):
        radio2 = 'Element exists'
    if radio2 == 'Element exists': 
        driver.find_element_by_id("label-status_notifysp-1").click()
    complete = driver.find_element_by_xpath("//select[@name='request_status_complete']/option[text()='Complete']")
    complete.click()
    driver.find_element_by_id("comments").send_keys("Sponsor is not part of [institutional] affiliate-emailed sponsor") #The text to fill into the comments box of the User Request Form
    submit = driver.find_element_by_id("submit-btn-saverecord")
    submit.click()

def Users(record): #after first user, gets the name of subsequent user's name and email 
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
        if (Users(i) != "[domain name of email address]" and Users(i) != "[domain name of email address]") and User(i) not in personalemails and Sponsor(i) in institAffiliates:
            username.columns= ['Username']
            user = project.export_records(records=[i], fields=[fname+j, lname+j, email+j,'sponsoremail', userexpiration+j, 'user_notes'], format='df')
            user.insert(loc=3, column='Institution ID', value=np.nan)
            result = pd.concat([username, user], axis=1)
            with open('Template.csv', 'a') as fd:
                result.to_csv(fd, header=False, index=False)
            driver = webdriver.Chrome('[system path to your chromedriver.exe]') #Create a automated test webpage e.g. C:\\Users\\cath\\Documents\\chromedriver_win32\\chromedriver.exe
            driver.get("[link to go to add users as a bulk upload]") #Enter REDCap link to upload new REDCap users
            submit = driver.find_element_by_name("fname").send_keys("[path to csv file]") #types in the first name information from the csv file e.g. C:\\Users\\cath\\Documents\\REDCap Automated Email\\Template.csv
            f = open('Template.csv')
            csv_f = csv.reader(f)
            next(csv_f)
            for csv in csv_f:
                domain = csv[0].split('@')
                print(domain[1])
                if domain[1] not in institAffiliates: 
                    driver.find_element_by_name("allow_create_db").click()
            driver.find_element_by_name("submit").click() #Submits the user accounts to be made as REDCap acocunts
            image_elements = driver.find_elements_by_xpath("//div[@class='red']/img")
            for image in image_elements:
                img_src = image.get_attribute("src") #spits the image source for either exclamation mark or check mark or something
            if img_src == 'https://redcap.edu/redcap_protocols/redcap_v8.10.20/Resources/images/exclamation.png':
                driver.get("[Finds the link to the project with the id parameters to be the one as called by the system arguments, e.g. first part of link + i + last part of link]")
                radio = driver.find_element_by_id("label-status-1")
                radio.click()
                if driver.find_element_by_id("label-status_notifysp-1"):
                    radio2 = 'Element exists'
                if radio2 == 'Element exists': 
                    driver.find_element_by_id("label-status_notifysp-1").click()
                complete = driver.find_element_by_xpath("//select[@name='request_status_complete']/option[text()='Complete']")
                complete.click()
                driver.find_element_by_id("comments").send_keys("User already has account - emailed user and sponsor")
                submit = driver.find_element_by_id("submit-btn-saverecord")
                submit.click()
        elif (email[1] == "[domain name of email address]" or email[1] == "[domain name of email address]"):
            print("ERROR! Username's email is [institutional] personal email: {0}".format(Users(i)))
            #this will fill in the comments box of the Username Request Form 
            driver.get("[Finds the link to the project with the id parameters to be the one as called by the system arguments, e.g. first part of link + i + last part of link]")
            radio = driver.find_element_by_id("label-status-1")
            radio.click()
            if driver.find_element_by_id("label-status_notifysp-1"):
                radio2 = 'Element exists'
            if radio2 == 'Element exists': 
                driver.find_element_by_id("label-status_notifysp-1").click()
            complete = driver.find_element_by_xpath("//select[@name='request_status_complete']/option[text()='Complete']")
            complete.click()
            driver.find_element_by_id("comments").send_keys("User has institutional ID-emailed user/sponsor ") #The text to fill into the comments box of the User Request Form
            submit = driver.find_element_by_id("submit-btn-saverecord")
            submit.click()
        elif Sponsor(i) not in institAffiliates:
            print("ERROR! Sponsor's email is not part of [institutional] affiliates: {0}".format(Sponsor(i)))
            #this will fill in the comments box of the Username Request Form if the sponsor is not from the institution
            driver.get("[Finds the link to the project with the id parameters to be the one as called by the system arguments, e.g. first part of link + i + last part of link]")
            radio = driver.find_element_by_id("label-status-1")
            radio.click()
            if driver.find_element_by_id("label-status_notifysp-1"):
                radio2 = 'Element exists'
            if radio2 == 'Element exists': 
                driver.find_element_by_id("label-status_notifysp-1").click()
            complete = driver.find_element_by_xpath("//select[@name='request_status_complete']/option[text()='Complete']")
            complete.click()
            driver.find_element_by_id("comments").send_keys("Sponsor is not part of [institutional] affiliate-emailed sponsor") #The text to fill into the comments box of the User Request Form
            submit = driver.find_element_by_id("submit-btn-saverecord")
            submit.click()

#SENDING EMAILS OUT TO SPONSORS TO NOTIFY THEM OF THE STATUS OF THE USER ACCOUNTS
sponsors = project.export_records(records=[i], fields=['numberofusers', 'sponsorname', 'sponsoremail']) #exports record data from the record ID I called 
sponsorname = project.export_records(records=[i], fields = ['sponsorname'], format='df') #exports record data of sponsor's name
sponsorname = sponsorname.to_string(index=False, header=False) #converts sponsorname from unicode to string

user_dict = {} 
user = (project.export_records(records=[i], fields=['email', 'firstname', 'lastname']))
user_dict.update([('user_1', (user[0]['firstname']+" "+user[0]['lastname']+" (username: "+user[0]['email']+")"))]) # puts all data from 'user' into dictionary-> {user_1: firstname, lastname, email}
numofusers = int(sponsors[0]['numberofusers']) #converts string of 'numberofusers' to integer

sponsoremail = project.export_records(records=[i], fields=['sponsoremail'], format='df') #exports record data of sponsor's email
sponsemail = sponsoremail.to_string(index=False, header=False) #converts sponsoremail from unicdoe to string
sponsoremail = sponsemail.split('@')[1] #splits sponsemail where @ symbol is

if sponsoremail not in institAffiliates: #checks if sponsor is part of [institutional] affiliated institution
    name = user[0]['firstname'] + ' ' + user[0]['lastname'] #get the users' first and last name
    subject = "User Account(s) Requests"
    recipient = "[email address of user]"
    sender = "[email address of sender]"
    cc = "[email address of anyone you want to cc]"
    password = base64.b64decode(b'[whatever you password is]') #Your associated password to enter REDCap
    password = password.decode("utf-8")

    message = MIMEMultipart("alternative")
    message["From"] = sender 
    message["To"] = recipient
    message["Subject"] = subject
    message["Cc"] = cc
    html = """\
    <html>
        <body>
            <p style= "color:#1E497D; font-family: Calibri, sans-serif;">Dear {sponsorname}, <br>
                <br>Thank you for filling out the username request form for {name} and others.</br>
                I noticed that you have an email address that is not affiliated with [institutional]  or are partners with them.
                You must fill in a sponsor that is a member of [institutional] partner institutions.
                <br>Thank you! </br>
                <br>Best,</br>
                Catherine Ng </br>
            </p>
        </body>
    </html>""".format(**locals())
    part2 = MIMEText(html,"html")
    message.attach(part2)
    text = message.as_string()

    s = smtplib.SMTP("smtp.office365.com", 587)
    s.ehlo()
    s.starttls()
    s.login(sender, password)
    s.sendmail(sender, recipient, text)
    print('done!')
    s.quit()

email = user[0]['email'] #finds email address of all users' emails
email = email.split('@') #splits user emails where @ symbol is 

#if there is only one user request 
if numofusers == 1:
    if email[1] in personalemails: 
        name = user[0]['firstname'] + ' ' + user[0]['lastname']
        subject = "User Account(s) Requests"
        recipient = "[email address of user]"
        sender = "[email address of sender]"
        cc = "[email address of anyone you want to cc]"
        password = base64.b64decode(b'[whatever you password is]') #Your associated password to enter REDCap
        password = password.decode("utf-8")

        message = MIMEMultipart("alternative")
        message["From"] = sender
        message["To"] = recipient
        message["Subject"] = subject
        message["Cc"] = cc
        html = """\
        <html>
            <body>
                <p style= "color:#1E497D; font-family: Calibri, sans-serif;"> Dear {sponsorname}, <br>
                    <br>Thank you for submitting the REDCap Username Request form for {name}.</br>
                    
                    For {name}'s own safety, please do not provide a personal email address, but an institutional email address. 
                    You may re-submit a new user request form with his/her institutional email address, or provide the email address so I can edit the username request form for you.
                    <br>Please let me know if you have any further questions.</br>
                    
                    <br>Best,</br>
                    Catherine Ng </br>
                </p>
            </body>
        </html>""".format(**locals())
        part2 = MIMEText(html,"html")
        message.attach(part2)
        text = message.as_string()

        s = smtplib.SMTP("smtp.office365.com", 587)
        s.ehlo()
        s.starttls()
        s.login(sender, password)
        s.sendmail(sender, recipient, text)
        print('done!')
        s.quit()
    elif email[1] == "[domain name of email address]" or email[1] == "[domain name of email address]":
        name = user[0]['firstname'] + ' ' + user[0]['lastname']
        user_items = []
        for item in user_dict.values():
            user_items.append(item) 
        items = ["{}\n".format(value) for (key,value) in user_dict.items()]
        items = "".join(items)
        print(items)
        subject = "User Account(s) Requests"
        recipient = "[email address of user]"
        sender = "[email address of sender]"
        cc = "[email address of anyone you want to cc]"
        password = base64.b64decode(b'[whatever you password is]') #Your associated password to enter REDCap
        password = password.decode("utf-8")

        message = MIMEMultipart("alternative")
        message["From"] = sender
        message["To"] = recipient
        message["Subject"] = subject
        message["Cc"] = cc
        html = """\
        <html>
            <body>
                <p style= "color:#1E497D; font-family: Calibri, sans-serif;">Dear {sponsorname}, <br>
                <br>Thank you for submitting the REDCap Username Request form for {name}.</br>
                <br>Since {items} <font color=#1E497D>is from [institution] , by default, he/she can access REDCap with his/her [institutional] ID and associated password.</br>

                <br>He/she can log in to REDCap at <a href = "https://redcap.edu/redcap_protocols/index.php">https://redcap.ctsc.weill.cornell.edu/redcap_protocols/index.php</a>  with his/her ID username 
                (the seven-character username consisting of the first section of your [institutional] email address) & password. </br>
                <br><br>Best,</br></br>

                Catherine Ng</font></br>
                </p>
            </body>
        </html>""".format(**locals())
        part2 = MIMEText(html,"html")
        message.attach(part2)
        text = message.as_string()
        s = smtplib.SMTP("smtp.office365.com", 587)
        s.ehlo()
        s.starttls()
        s.login(sender, password)
        s.sendmail(sender, recipient, text)
        print('done!')
        s.quit()
    else:
        user_items = []
        for item in user_dict.values():
            user_items.append(item)
        non_WCM = [x for x in user_items if "[domain name of email address]" not in x and "[domain name of email address]" not in x]
        user_table = "<html><table><tr><th><font color=#1E497D>New User(s)</th></tr>"

        for item in non_WCM:
            table_items = "<tr><td><font face='calibri' color=#1E497D>"+item+ "</font></tr></td>"
            user_table = user_table+table_items

        user_table = user_table+"</font></table></html>"
        hs = open("listOfUsers.html",'w')
        hs.write(user_table)
        items = ["{}, ".format(value) for (key,value) in user_dict.items()]
        items = "".join(items)

        subject= "User Account Request"
        recipient = "[email address of user]"
        sender = "[email address of sender]"
        cc = "[email address of anyone you want to cc]"
        password = base64.b64decode(b'[whatever you password is]') #Your associated password to enter REDCap
        password = password.decode("utf-8")

        message = MIMEMultipart("alternative")
        message["From"] = sender
        message["To"] = recipient
        message["Subject"] = subject
        message["Cc"] = cc

        html = """\
        <html>
            <body>
                <p style = "color:#1E497D; font-family: Calibri, sans-serif;">Dear {1}, <br>
                    <br>The REDCap account for:</br>{0} 
                    <br><font color=#1E497D>has been created. You may now add him/her to your project(s) using his/her username (email address).</br>
                    Please let me know if you have any questions.
                    <br><br>Thank you!</br></br>
                    <br>Best,</br>
                    Catherine Ng </br></font>
                </p>
            </body>
        </html>
        """.format(user_table, sponsorname)
        part2 = MIMEText(html,"html")
        message.attach(part2)
        text = message.as_string()

        s = smtplib.SMTP("smtp.office365.com", 587)
        s.ehlo()
        s.starttls()
        s.login(sender, password)
        s.sendmail(sender, recipient, text)
        print('done!')
        s.quit()

#if there is more than one user request
if numofusers > 1: 
    for j in range(2, numofusers+1, 1):
        j = str(j)
        email = 'email_'
        fname = 'firstname_'
        lname = 'lastname_'
        user = project.export_records(records=[i], fields=[email+j, fname+j, lname+j])
        user_dict.update( [('user_'+j, (user[0][fname+j]+" "+user[0][lname+j]+" (username: "+user[0][email+j]+")"))] )
        secemail = user[0][email+j]
        secemail = secemail.split('@')
        if secemail[1] in personalemails:
            name = user[0][fname+j] + ' ' + user[0][lname+j]
            print("Personal Email:" + name)
            
            subject = "User Account(s) Requests"
            recipient = "[email address of user]"
            sender = "[email address of sender]"
            cc = "[email address of anyone you want to cc]"
            password = base64.b64decode(b'[whatever you password is]') #Your associated password to enter REDCap
            password = password.decode("utf-8")

            message = MIMEMultipart("alternative")
            message["From"] = sender
            message["To"] = recipient
            message["Subject"] = subject
            message["Cc"] = cc
            html = """\
            <html>
                <body>
                    <p style= "color:#1E497D; font-family: Calibri, sans-serif;"> Dear {sponsorname}, <br>
                        <br>Thank you for submitting the REDCap Username Request form for {name}.</br>
                        
                        For {name}'s own safety, please do not provide a personal email address, but an institutional email address. 
                        You may re-submit a new user request form with his/her institutional email address, or provide the email address so I can edit the username request form for you.
                        <br>Please let me know if you have any further questions.</br>
                        
                        <br>Best,</br>
                        Catherine Ng </br>
                    </p>
                </body>
            </html>""".format(**locals())
            part2 = MIMEText(html,"html")
            message.attach(part2)
            text = message.as_string()

            s = smtplib.SMTP("smtp.office365.com", 587)
            s.ehlo()
            s.starttls()
            s.login(sender, password)
            s.sendmail(sender, recipient, text)
            print('done!')
            s.quit()
            
    if secemail[1] == "[domain name of email address]" or secemail[1] == "[domain name of email address]": #if user's email is [institutional] email them!
        name = user[0][fname+j] + ' ' + user[0][lname+j] #takes firstname and lastname of user to use in email
        print("Institutional Email: " + name)
        
        subject = "User Accounts Requests"
        recipient = "[email address of user]"
        sender = "[email address of sender]"
        cc = "[email address of anyone you want to cc]"
        password = base64.b64decode(b'[whatever you password is]') #Your associated password to enter REDCap
        password = password.decode("utf-8")

        message = MIMEMultipart("alternative")
        message["From"] = sender
        message["To"] = recipient
        message["Subject"] = subject
        message["Cc"] = cc
        html = """\
        <html>
            <body>
                <p style= "color:#1E497D;font-family: Calibri, sans-serif;">Dear {sponsorname}, <br>
                <br>Since {name} <font color=#1E497D>is from [Institution], by default, he/she can access REDCap with his/her Institutional ID and associated password.</br>

                <br>He/she can log in to REDCap at <a href = "https://redcap.edu">https://redcap.edu</a>  with his/her CWID (the seven-character username consisting of the first section of your [institutional] email address) & password. </br>
                <br><br>Best,</br></br>

                Catherine Ng</font></br>
                </p>
            </body>
        </html>""".format(**locals())
        part2 = MIMEText(html,"html")
        message.attach(part2)
        text = message.as_string()
        s = smtplib.SMTP("smtp.office365.com", 587)
        s.ehlo()
        s.starttls()
        s.login(sender, password)
        s.sendmail(sender, recipient, text)
        print('done!')
        s.quit()  
        
    user_items = []
    users_items= []
    for item in user_dict.values():
        user_items.append(item)
    non_WCM = [x for x in user_items if "[domain name of email address]" not in x and "[domain name of email address]" not in x and "yahoo.com" not in x and "gmail.com" not in x]
    user_table = "<html><table><tr><th><font color=#1E497D>New User(s)</th></tr>"

    for item in non_WCM:
        table_items = "<tr><td><font color=#1E497D face='calibri'>"+item+ "</font></tr></td>"
        user_table = user_table+table_items

    user_table = user_table+"</font></table></html>"

    hs = open("listOfUsers.html",'w')
    hs.write(user_table)
    
    subject= "User Accounts Request"
    recipient = "[email address of user]"
    sender = "[email address of sender]"
    cc = "[email address of anyone you want to cc]"
    password = base64.b64decode(b'[whatever you password is]') #Your associated password to enter REDCap
    password = password.decode("utf-8")

    message = MIMEMultipart("alternative")
    message["From"] = sender
    message["To"] = recipient
    message["Subject"] = subject
    message["Cc"] = cc

    html = """\
    <html>
        <body>
            <p style = "font-family: Calibri, sans-serif; color:#1E497D;">Dear {1}, <br>
                <br>The REDCap account(s) for:</br>{0} 
                <br><font color=#1E497D>have been created. You may now add them to your project(s) using their username (email address).</br>
                Please let me know if you have any questions.
                <br><br>Thank you!</br></br>
                <br>Best,</br>
                Catherine Ng </br></font>
            </p>
        </body>
    </html>
    """.format(user_table, sponsorname)
    part2 = MIMEText(html,"html")
    message.attach(part2)
    text = message.as_string()

    s = smtplib.SMTP("smtp.office365.com", 587) #send message via own SMTP server
    s.ehlo() #identify ourselves, prompting server for supported features 
    s.starttls()
    s.login(sender, password)
    s.sendmail(sender, recipient, text)
    print('done!')
    s.quit()