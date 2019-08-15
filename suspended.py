import sys
import redcap
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import base64 #binary to text encoding schemes that represent binary data in ASCII string format. encodes string into base 64
import win32com.client, datetime

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) #this is the inbox. 
messages = inbox.Items
#message = messages.Getlast
#body_content = message.body
#print(body_content)

y = datetime.datetime(2019, 7, 18, 11, 3, 50)

y = (datetime.date.today () - datetime.timedelta (days=0)).strftime("%y-%m-%d")
print(y)

m = "19-05-01"

for message in messages:
    if message.UnRead == True and message.SentOn.strftime("%y-%m-%d")==y:
        print(message.subject)

#Reactivating a Suspended User--GOOD!!
for message in messages:
    if message.UnRead == True and message.SentOn.strftime("%y-%m-%d")==y and ("suspend" in message.body or "suspend" in message.subject):
        subject = "Reactivating Suspended User"
        recipient = "can2010@med.cornell.edu"
        sender = "can2010@med.cornell.edu"
        #sender = message.Sender.Address
        sendername= message.Sender
        #cc = "redcap@med.cornell.edu"
        password = base64.b64decode(b'UG9vcHNAIzQ=')
        password = password.decode("utf-8")

        message = MIMEMultipart("alternative")
        message["From"] = sender 
        message["To"] = recipient
        message["Subject"] = subject
        #message["Cc"] = cc
        html = """\
        <html>
            <body>
                <p style= "color:#1E497D; font-family: Calibri, sans-serif;">Dear {sendername}, <br>
                    <br>Your account has been suspended, but I have unsuspended you. However, your access to each research project will remain expired. 
                    
                    You will need to contact the project administrator/project coordinators to remove the expiration date I have put on your project access.
                    
                    <br>Please note: For non-WCM/NYP users, REDCap accounts are automatically suspended after six months of inactivity.</br>
                    </br>
                    <br>Please let me know if you have any other further questions. </br>

                    <br>Best, </br> 
                    Catherine Ng
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

#REDCap Training Session--GOOD!!
for message in messages: 
    if message.UnRead == True and message.SentOn.strftime("%y-%m-%d")==y and ("training" in message.body):
        subject = "REDCap Training Session"
        recipient = "can2010@med.cornell.edu"
        sender = "can2010@med.cornell.edu"
        #sender = message.Sender.Address
        sendername= message.Sender
        #cc = "redcap@med.cornell.edu"
        password = base64.b64decode(b'UG9vcHNAIzQ=')
        password = password.decode("utf-8")

        message = MIMEMultipart("alternative")
        message["From"] = sender 
        message["To"] = recipient
        message["Subject"] = subject
        #message["Cc"] = cc
        html = """\
        <html>
            <body>
                <p style= "color:#1E497D; font-family: Calibri, sans-serif;">Dear {sendername}, <br>
            <br>Besides logging into REDCap and clicking on “Training Videos” that can be found on the top of the page, which can help you build the data collection surveys </br>
            <br>on your own, you can also schedule a REDCap meeting with Jessie, my manager. I have cc’d her to this email.</br> 

            <br>You are also welcome to request service that will help you build your surveys as well.</br>

            <br>Let me know which you would prefer.</br>

            <br>Best, </br>
            Catherine Ng 
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

#NYP Password Reset
#and "@med.cornell.edu" in message.Sender
for message in messages:
    if message.UnRead == True and message.SentOn.strftime("%y-%m-%d")==y and ("error" in message.body or "unable to gain access" in message.body or "unable to log" in message.body or "does not work" in message.body):
        subject = "REDCap Training Session"
        recipient = "can2010@med.cornell.edu"
        sender = "can2010@med.cornell.edu"
        #sender = message.Sender.Address
        sendername= message.Sender
        #cc = "redcap@med.cornell.edu"
        password = base64.b64decode(b'UG9vcHNAIzQ=')
        password = password.decode("utf-8")

        message = MIMEMultipart("alternative")
        message["From"] = sender 
        message["To"] = recipient
        message["Subject"] = subject
        #message["Cc"] = cc
        html = """\
        <html>
            <body>
                <p style= "color:#1E497D; font-family: Calibri, sans-serif;">Dear {sendername}, <br>
            It seems like the NYP credentials did not sync over to Cornell for some reason. </br>
            <br>You should be able to access REDCap after 10 minutes after initiating a password reset at https://pod.nyp.org/nyp/psf.exe#2</br>
            <br>Please let me know if there are any more issues.</br>

            <br>Best,</br>
            Catherine Ng
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

#Non WCM/NYP External User--GOOD!!
for message in messages:
    if message.UnRead == True and message.SentOn.strftime("%y-%m-%d")==y and ("add them to a project" in message.body or "add the following users" in message.body or "make accounts" in message.body or "create accounts" in message.body):
        subject = "REDCap Training Session"
        recipient = "can2010@med.cornell.edu"
        sender = "can2010@med.cornell.edu"
        #sender = message.Sender.Address
        sendername= message.Sender
        #cc = "redcap@med.cornell.edu"
        password = base64.b64decode(b'UG9vcHNAIzQ=')
        password = password.decode("utf-8")

        message = MIMEMultipart("alternative")
        message["From"] = sender 
        message["To"] = recipient
        message["Subject"] = subject
        #message["Cc"] = cc
        html = """\
        <html>
            <body>
                <p style= "color:#1E497D; font-family: Calibri, sans-serif;">Dear {sendername}, <br>
            <br>In order to add an external (non-WCM/NYP) user to REDCap, a username request form must be filled out at https://redcap.ctsc.weill.cornell.edu/redcap_protocols/</br>
            <br>On the website, there is a button link called “Request REDCap User Account,” where it takes you to the username request form.</br>
 
            Please let me know if you have any further questions. 
 
            <br>Best,</br>
            Catherine Ng
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

#Giving a User Project Access
for message in messages:
    if message.UnRead == True and message.SentOn.strftime("%y-%m-%d")==y and ("gain access" in message.body or "request access" in message.body or "project access" in message.body or "add me as a user" in message.body):
        subject = "REDCap Training Session"
        recipient = "can2010@med.cornell.edu"
        sender = "can2010@med.cornell.edu"
        #sender = message.Sender.Address
        sendername= message.Sender
        #cc = "redcap@med.cornell.edu"
        password = base64.b64decode(b'UG9vcHNAIzQ=')
        password = password.decode("utf-8")

        message = MIMEMultipart("alternative")
        message["From"] = sender 
        message["To"] = recipient
        message["Subject"] = subject
        #message["Cc"] = cc
        html = """\
        <html>
            <body>
                <p style= "color:#1E497D; font-family: Calibri, sans-serif;">Dear {sendername}, <br>
            <br>Someone with User Rights in the project should be able to add you under “User Rights” using your email address. You can reach out to any of the study coordinators to add you to their project.</br>
 
            <br>Best,</br>
            Catherine Ng
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

#API Rights--GOOD!!
for message in messages:
    if message.UnRead == True and message.SentOn.strftime("%y-%m-%d")==y and "requests an API token" in message.subject:
        subject = "API Rights"
        recipient = "can2010@med.cornell.edu"
        sender = "can2010@med.cornell.edu"
        #sender = message.Sender.Address
        sendername= message.Sender
        #cc = "redcap@med.cornell.edu"
        password = base64.b64decode(b'UG9vcHNAIzQ=')
        password = password.decode("utf-8")

        message = MIMEMultipart("alternative")
        message["From"] = sender 
        message["To"] = recipient
        message["Subject"] = subject
        #message["Cc"] = cc
        html = """\
        <html>
            <body>
                <p style= "color:#1E497D; font-family: Calibri, sans-serif;">Dear {sendername}, <br>
            <br>Just a reminder, you must have approval to collect data on mobiles on your IRB. 
            
            <br>Thank you!</br>
 
            <br>Best,</br>
            Catherine Ng</br>
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

#If user's email is incorrect
for message in messages:
    if message.UnRead == True and message.SentOn.strftime("%y-%m-%d")==y and "Undeliverable: REDCap access granted" in message.subject:
        subject = "Error in User Account Request"
        recipient = "can2010@med.cornell.edu"
        sender = "can2010@med.cornell.edu"
        #sender = message.Sender.Address
        sendername= message.Sender
        #cc = "redcap@med.cornell.edu"
        password = base64.b64decode(b'UG9vcHNAIzQ=')
        password = password.decode("utf-8")

        message = MIMEMultipart("alternative")
        message["From"] = sender 
        message["To"] = recipient
        message["Subject"] = subject
        #message["Cc"] = cc
        html = """\
        <html>
            <body>
                <p style= "color:#1E497D; font-family: Calibri, sans-serif;">Dear {sendername}, <br>
            <br>Thank you for submitting a REDCap username request form.

            <br>We were unable to create REDCap account access for you because there was an issue with his/her email that was provided to us- [email address]. As of now, we are unable to create an account for him/her. </br>

            Can you verify your email address again?
            <br>Thank you.</br>

            Best, 
            Catherine Ng  
            
            <br>Thank you!</br>
 
            <br>Best,</br>
            Catherine Ng</br>
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

#If User does not create one form or the other: 
""" for message in messages:
    if message.UnRead == True and message.SentOn.strftime("%y-%m-%d")==y and ("Request to Copy Project" in message.subject or "Request to Create New Project" in message.subject) and "Request for New or Additional REDCap Database for Research (development)" in message.subject:
        subject = "API Rights"
        recipient = "can2010@med.cornell.edu"
        sender = "can2010@med.cornell.edu"
        #sender = message.Sender.Address
        sendername= message.Sender
        #cc = "redcap@med.cornell.edu"
        password = base64.b64decode(b'UG9vcHNAIzQ=')
        password = password.decode("utf-8")

        message = MIMEMultipart("alternative")
        message["From"] = sender 
        message["To"] = recipient
        message["Subject"] = subject
        #message["Cc"] = cc
        html = """\
        <html>
            <body>
                <p style= "color:#1E497D; font-family: Calibri, sans-serif;">Dear {sendername}, <br>
            <br>Just a reminder, you must have approval to collect data on mobiles on your IRB. 
            
            <br>Thank you!</br>
 
            <br>Best,</br>
            Catherine Ng</br>
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
        s.quit() """

#forms development services
""" It looks like you have requested for REDCap form development (fee-based service). 

If you have any forms to send us so we can start working on form development, we can possibly estimate how long it will take for us to complete the project. Usually, projects take around ~1-3 hours to complete. Please note that this fee-based service is $100 per hour. 

If you would like to discuss about this in person, we can discuss in person or in a web meeting instead. 

Please let me know!

Thank you, 
Catherine Ng """

#If the user asks about the payment fee for non-research, respond 
""" Dear Ryan, 

Yes, but if it is practice for research project, you can submit your initiation form as a research project (provide all the necessary information, i.e. PI name, Research name, etc.), and then after you are done familiarizing with REDCap, you can change your project as being associated with research on REDCap.
If you simply want to practice on REDCap, you can also login to REDCap with username: demo_user, password: DemoUser1. Otherwise, there will be a fee when submitting the initiation form as non-research.

Best,
Catherine Ng  """

