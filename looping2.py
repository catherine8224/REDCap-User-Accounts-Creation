import sys
import redcap
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import base64 #binary to text encoding schemes that represent binary data in ASCII string format. encodes string into base 64

def call_record(i): #function that prints out record ID
    print('record:', i)

if __name__=='__main__': #takes in one argument in the command prompt where I declare the record ID
    i = str(sys.argv[1])
    call_record(i)
    
project = redcap.Project(
    'https://redcap.ctsc.weill.cornell.edu/redcap_protocols/api/', 'C0F7DAD56835B9DAA53CAC272A060E4F')

wcmAffiliates = ['med.cornell.edu', 'myhunter.cuny.edu', 'cornell.edu', 'hss.edu', 'hunter.cuny.edu', 
                       'HSS.edu', 'HSS.EDU', 'mskcc.org', 'amcny.org', 'nyp.org', 'NYP.org', 'burke.org','cumc.columbia.edu'] #a list used to check if sponsors are part of WCM institution, if not, email them
personalemails= ['yahoo.com', 'gmail.com', 'hotmail.com', 'mail.com', 'aol.com', 'icloud.com', 'me.com', 'mac.com'] #a list used to check if users provided personal emails

sponsors = project.export_records(records=[i], fields=['numberofusers', 'sponsorname', 'sponsoremail']) #exports record data from record ID I called 
sponsorname = project.export_records(records=[i], fields = ['sponsorname'], format='df') #exports record data of sponsor's name
sponsorname = sponsorname.to_string(index=False, header=False) #converts sponsorname from unicode to string

sponsoremail = project.export_records(records=[i], fields=['sponsoremail'], format='df') #exports record data of sponsor's email
sponsemail = sponsoremail.to_string(index=False, header=False) #converts sponsoremail from unicdoe to string
sponsoremail = sponsemail.split('@')[1] #splits sponsemail where @ symbol is

user_dict = {} 
user = (project.export_records(records=[i], fields=['email', 'firstname', 'lastname']))
'''user_dict.update([('firstname_1', user[0]['firstname']),('lastname_1', user[0]['lastname']), ('email_1', user[0]['email'])])'''
user_dict.update([('user_1', (user[0]['firstname']+" "+user[0]['lastname']+" (username: "+user[0]['email']+")"))]) # puts all data from 'user' into dictionary-> {user_1: firstname, lastname, email}
numofusers = int(sponsors[0]['numberofusers']) #converts string of 'numberofusers' to integer

if sponsoremail not in wcmAffiliates: #checks if sponsor is part of WCM affiliated institution
    name = user[0]['firstname'] + ' ' + user[0]['lastname'] #get the users' first and last name
    subject = "User Account(s) Requests"
    recipient = "can2010@med.cornell.edu"
    sender = "can2010@med.cornell.edu"
    cc = "redcap@med.cornell.edu"
    password = base64.b64decode(b'UG9vcHNAIzQ=')
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
                I noticed that you have an email address that is not affiliated with WCM/NYP or are partners with them.
                You must fill in a sponsor that is a member of WCM CTSC partner institutions (e.g. WCM/NYP, Cornell, MSKCC, HSS, Hunter College, AMC, Burke, etc). 
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
        recipient = "can2010@med.cornell.edu"
        sender = "can2010@med.cornell.edu"
        cc = "redcap@med.cornell.edu"
        password = base64.b64decode(b'UG9vcHNAIzQ=')
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
    elif email[1] == "med.cornell.edu" or email[1] == "nyp.org":
        name = user[0]['firstname'] + ' ' + user[0]['lastname']
        user_items = []
        for item in user_dict.values():
            user_items.append(item) 
        items = ["{}\n".format(value) for (key,value) in user_dict.items()]
        items = "".join(items)
        print(items)
        subject = "User Account(s) Requests"
        recipient = "can2010@med.cornell.edu"
        sender = "can2010@med.cornell.edu"
        cc= "cathygreat828@gmail.com"
        password = base64.b64decode(b'UG9vcHNAIzQ=')
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
                <br>Since {items} <font color=#1E497D>is from WCM/NYP, by default, he/she can access REDCap with his/her WCM/NYP CWID and associated password.</br>

                <br>He/she can log in to REDCap at <a href = "https://redcap.ctsc.weill.cornell.edu/redcap_protocols/index.php">https://redcap.ctsc.weill.cornell.edu/redcap_protocols/index.php</a>  with his/her CWID (the seven-character username consisting of the first section of your WCM email address) & password. </br>
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
        non_WCM = [x for x in user_items if "med.cornell.edu" not in x and "nyp.org" not in x]
        user_table = "<html><table><tr><th><font color=#1E497D>New User(s)</th></tr>"

        for item in non_WCM:
            table_items = "<tr><td><font face='calibri' color=#1E497D>"+item+ "</font></tr></td>"
            user_table = user_table+table_items

        user_table = user_table+"</font></table></html>"
        hs = open("listOfUsers.html",'w')
        hs.write(user_table)
        items = ["{}, ".format(value) for (key,value) in user_dict.items()]
        items = "".join(items)
        #sponsorname = "".join(sponsorname)

        subject= "User Account Request"
        recipient = 'can2010@med.cornell.edu'
        sender= 'can2010@med.cornell.edu'
        cc = 'cathygreat828@gmail.com'
        password = base64.b64decode(b'UG9vcHNAIzQ=')
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
            #name = user[0]['firstname'] + ' ' + user[0]['lastname']
            '''
            subject = "User Account(s) Requests"
            recipient = "can2010@med.cornell.edu"
            sender = "can2010@med.cornell.edu"
            cc = "cathygreat828@gmail.com"
            password = base64.b64decode(b'UG9vcHNAIzQ=')
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
            '''
    if secemail[1] == "med.cornell.edu" or secemail[1] == "nyp.org": #if user's email is WCM/NYP email them!
        name = user[0][fname+j] + ' ' + user[0][lname+j] #takes firstname and lastname of user to use in email
        print("WCM/NYP Email: " + name)
        '''
        subject = "User Accounts Requests"
        recipient = "can2010@med.cornell.edu"
        sender = "can2010@med.cornell.edu"
        cc= "cathygreat828@gmail.com"
        password = base64.b64decode(b'UG9vcHNAIzQ=')
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
                <br>Since {name} <font color=#1E497D>is from WCM/NYP, by default, he/she can access REDCap with his/her WCM/NYP CWID and associated password.</br>

                <br>He/she can log in to REDCap at <a href = "https://redcap.ctsc.med.cornell.edu">https://redcap.ctsc.med.cornell.edu</a>  with his/her CWID (the seven-character username consisting of the first section of your WCM email address) & password. </br>
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
        '''
    user_items = []
    users_items= []
    for item in user_dict.values():
        user_items.append(item)
    #print(user_items)  ['Samuel Pierre (username: spierre@gheskio.org)', 'Bajo Joseph (username: josephmarieb88@gmail.com)', 'Yvetot Joseph (username: yvj4001@med.cornell.edu)']
    non_WCM = [x for x in user_items if "med.cornell.edu" not in x and "nyp.org" not in x and "yahoo.com" not in x and "gmail.com" not in x]
    #personalemails= ['yahoo.com', 'gmail.com', 'hotmail.com', 'mail.com', 'aol.com', 'icloud.com', 'me.com', 'mac.com']
    user_table = "<html><table><tr><th><font color=#1E497D>New User(s)</th></tr>"

    for item in non_WCM:
        table_items = "<tr><td><font color=#1E497D face='calibri'>"+item+ "</font></tr></td>"
        user_table = user_table+table_items

    user_table = user_table+"</font></table></html>"

    hs = open("listOfUsers.html",'w')
    hs.write(user_table)
    #items = ["{}, ".format(value) for (key,value) in user_dict.items()]
    #items = "".join(items)
    #sponsorname = "".join(sponsorname)
    #print(items)
    
    subject= "User Accounts Request"
    recipient = 'can2010@med.cornell.edu'
    sender= 'can2010@med.cornell.edu'
    cc = 'cathygreat828@gmail.com'
    password = base64.b64decode(b'UG9vcHNAIzQ=')
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
                <br><font color=#1E497D>have been created. You may now add them to your project(s) using  their username (email address).</br>
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
    #s.Display(False)
    s.sendmail(sender, recipient, text)
    print('done!')
    s.quit()
    