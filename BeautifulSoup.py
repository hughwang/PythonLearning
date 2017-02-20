from bs4 import BeautifulSoup
import csv

soup = BeautifulSoup(open("Congressional43.html"))

#print(soup.prettify())

final_link = soup.p.a
final_link.decompose()

links = soup.find_all('a')

f = csv.writer(open("43rd_Congress.csv", "w"))
f.writerow(["Name", "Years", "Position", "Party", "State", "Congress", "Link"]) # Write column headers as the first line
trs = soup.find_all('tr')
for tr in trs:
    for link in tr.find_all('a'):
        fulllink = link.get ('href')
        print fulllink #print in terminal to verify results
    tds = tr.find_all("td")
    try: #we are using "try" because the table is not well formatted. This allows the program to continue after encountering an error.
        names = str(tds[0].get_text()) # This structure isolate the item by its column in the table and converts it into a string.
        years = str(tds[1].get_text())
        positions = str(tds[2].get_text())
        parties = str(tds[3].get_text())
        states = str(tds[4].get_text())
        congress = tds[5].get_text()

    except:
        print "bad tr string"
        continue #This tells the computer to move on to the next item after it encounters an error

    print names, years, positions, parties, states, congress
    f.writerow([names, years, positions, parties, states, congress, fulllink])

====================
h2 = soup.h2
h2.decompose()
print(soup.prettify())
===============

links = soup.find_all('a')
for link in links:
    names = link.contents[0]
    fullLink = link.get('href')
    print names
    print fullLink
    f.writerow([names, fullLink])

==================

letters = soup.find_all("div",class_ = "ec-statement")
