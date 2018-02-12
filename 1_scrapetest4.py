from urllib.request import urlopen
from bs4 import BeautifulSoup
import re


# html = urlopen("http://www.owgr.com/ranking")
html = open('docs/Official World Golf Ranking - Ranking.html')
bsObj = BeautifulSoup(html, "html.parser")
for item in bsObj.body.find('div', class_='table_container').table.tbody.findAll('tr'):
    # print(item.encode('utf8')) # gets rid of ascii encoding error
    # find age group string inside a <ul><li><span> ... </span></li></ul>
    rank = item.findNext('td').contents[1]
    # find gender in the next item in the <ul> ... </ul>
    name = item.findNext('td', {'class': 'name'}).a.string
    print(rank)
    print(name)
