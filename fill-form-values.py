from __future__ import print_function
import argparse
import mechanicalsoup
from getpass import getpass

parser = argparse.ArgumentParser(description="Login to GitHub.")
parser.add_argument("username")
args = parser.parse_args()

args.password = getpass("Please enter your GitHub password: ")

browser = mechanicalsoup.StatefulBrowser(
    soup_config={'features': 'lxml'},
    raise_on_404=True,
    user_agent='MyBot/0.1: mysite.example.com/bot_info',
)
# Uncomment for a more verbose output:
# browser.set_verbose(2)

browser.open("https://github.com/catherine8224/hello-world")
browser.follow_link("login")
browser.select_form('#login form')
browser["login"] = args.username
browser["password"] = args.password
resp = browser.submit_selected()

# Uncomment to launch a web browser on the current page:
# browser.launch_browser()

# verify we are now logged in
page = browser.get_current_page()
messages = page.find("div", class_="flash-messages")
if messages:
    print(messages.text)
assert page.select(".logout-form")

print(page.title.text)

# verify we remain logged in (thanks to cookies) as we browse the rest of
# the site
page3 = browser.open("https://github.com/MechanicalSoup/MechanicalSoup")
assert page3.soup.select(".logout-form")
'''
login_form = mechanicalsoup.Form(login_page.soup.select_one('#login form'))
login_form.input({"login":args.username, "password": args.password})
page2 = browser.submit(login_form, login_page.url)
messages = page2.soup.find("div", class_="flash-messages")
if messages:
    print(messages.text)
assert page2.soup.select(".logout-form")

print(page2.soup.title.text)

page3 = browser.get("https://github.com/MechanicalSoup/MechanicalSoup")
assert page3.soup.select(".logout-form")
'''

'''
import argparse 

parser = argparse.ArgumentParser(description = 'Process some integers.')
parser.add_argument('integers', type=int, nargs='+', help='an integer for the accumulator')
parser.add_argument('--sum', dest='accumulate', action='store_const', const=sum, default=max)
args = parser.parse_args()
print(args.accumulate(args.integers))
'''

# Fill-in the search form
'''browser.select_form('#search_form_homepage')
browser["q"] = "MechanicalSoup"
browser.submit_selected()

# Display the results
for link in browser.get_current_page().select('a.result__a'):
    print(link.text, '->', link.attrs['href'])'''