# REDCap User Accounts Creation

<img alt="https://img.shields.io/pypi/pyversions/requests.svg" src="https://img.shields.io/pypi/pyversions/requests.svg">

This project was created with the intention that others can use it to create a user account, assuming there is an existng REDCap project with a survey that allows users to sign up for a user account. The survey contains the user's *first name, last name, institutional email address*, and the *sponsor* that is responsible for the account (suspended, expired, etc.) One can then go to the Control Center to add the User Accounts, and use <img src = 'greenthing.gif'>**__Create users (bulk upload)__** to upload a csv file of all the user's names, email address, and sponsor. I use [Selenium](https://selenium-python.readthedocs.io/) to automate the process of clicking through webpages to create the User Accounts.I also use [MIMEText](https://docs.python.org/2/library/email.mime.html) to send emails to the users once the accounts have been created.</p> 


```diff
+ this was created using Python3.7 on Windows
```

## Setup Python

1. Install Python 2 or Python 3 from https://www.python.org/downloads/
2. Verify your installation by running `python` and `pip` from your terminal.
3. Install the Python `redcap` library:

    ```
    pip install redcap
    ```

# 1) Successful REDCap user account creation
![](Enable_Survey_3.gif)
made with the help of ScreenToGif


# 2) REDCap user account was already created:
![](Enable_Survey.gif)


# 3)Other issues:
<ul>
  <li>When the user is part of the main institution (have their own username and associated password and can automatically log into REDCap)</li>
  <IMG SRC="Enable_Survey_2.gif">
    
  <li> When the user provides a personal email </li>
  <IMG SRC="personalemail.png">
  
  <li> When the sponsor is not from the main institution or affiliated institutions </li>
  <IMG SRC="sponsemail.png">
</ul>
    
```python
    name = user[0]['firstname'] + ' ' + user[0]['lastname'] #get the users' first and last name
    subject = "User Account(s) Requests"
    recipient = "[the email address of receiver]"
    sender = "[your email address]"
    cc = "[additional email addresses]"
    password = base64.b64decode(b'[password]')
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
                I noticed that you have an email address that is not affiliated with [institution] or are partners with them.
                You must fill in a sponsor that is a member of [institution] partner institutions (e.g. ..., etc). 
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
```   

