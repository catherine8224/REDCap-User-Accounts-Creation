# REDCap User Accounts Creation
```html
<p style="color:red"> This project was created with the intention that others can use it to create a user account, assuming there is an existng REDCap project with a survey that allows users to sign up for a user account. The survey contains the user's first name, last name, institutional email address, and the sponsor that is responsible for the account (suspended, expired, etc.) One can then go to the Control Center to add the User Accounts, and use 'Bulk Upload' to upload a csv file of all the user's names, email address, and sponsor. I use <a href = 'https://selenium-python.readthedocs.io/'>Selenium</a> to automate the process of clicking through webpages to create the User Accounts.I also use MIMEText to send emails to the users once the accounts have been created.</p> 
```

```diff
+ this text is highlighted in green
- this text is highlighted in red
```

1) When the REDCap user account creation is successful:
![](Enable_Survey_3.gif)

2) When the REDCap user account was already created:
![](Enable_Survey.gif)


3)Other issues:
<ul>
  <li>When the user is part of the main institution (have their own username and associated password and can automatically log into REDCap)</li>
  <IMG SRC="Enable_Survey_2.gif">
  <li> When the user provides a personal email </li>
  <IMG SRC="personalemail.png">
  <li> When the sponsor is not from the main institution or affiliated institutions </li>
  <IMG SRC="sponsemail.png">
</ul>
