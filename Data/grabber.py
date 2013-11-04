
import urllib2
import re
from bs4 import BeautifulSoup
import xlwt

BASE = 'file:///C:/Projects/Sandbox/down/2013/msdn.microsoft.com/en-us/library/office/'
# BASE = 'file:///C:/Projects/Sandbox/down/2003/msdn.microsoft.com/en-us/library/'

def getCellInfo(sheet, row, tag):
    m = re.search("(.*)\s+Cell\s+\((.*)\s+Section", tag.text);
    if (m):
        sheet.write(row, 1, m.group(1))
        sheet.write(row, 2, m.group(2))
    
    print "row: {0} href: {1}".format(row, tag["href"])

    resp = urllib2.urlopen(BASE + tag["href"])
    bs = BeautifulSoup(resp)

    p = bs.find("p", text=re.compile(".*Cell name:.*"))
    if (p):
        sheet.write(row, 3, p.findNext("td").text.strip())

    p = bs.find("p", text=re.compile(".*Section index:.*"))
    if (p):
        sheet.write(row, 4, p.findNext("td").text.strip())

    p = bs.find("p", text=re.compile(".*Row index:.*"))
    if (p):
        sheet.write(row, 5, p.findNext("td").text.strip())

    p = bs.find("p", text=re.compile(".*Cell index:.*"))
    if (p):
        sheet.write(row, 6, p.findNext("td").text.strip())

    # p = bs.find("th", text="Value")
    p = bs.find("p", text="Value")
    if (p):
        values = []
        tr_no = 0
        for tr in p.findParent("table").findChildren("tr"):
            tr_no = tr_no + 1
            if tr_no == 1:
                continue;

            s = ''
            td_col  = 0
            for td in tr.findChildren("td"):
                td_col = td_col + 1
                if (td_col  == 1):
                    s = s + td.text.strip()
                if (td_col == 3):
                    s = s + "=" + td.text.strip()

            values.append(s)

        sheet.write(row, 7, ";".join(values))


URL = BASE + 'jj945060.aspx'
# URL = BASE+ 'aa246643(v=office.11).aspx'

response = urllib2.urlopen(URL)
bs = BeautifulSoup(response)

book = xlwt.Workbook()
sheet = book.add_sheet('sheet')

section = bs.find('div', {'class' : 'sectionblock'})
# section = bs.find('div', {'id': 'content'})

row = 1
for tag in section.findAll('a'):
    getCellInfo(sheet, row, tag)
    row = row + 1

book.save('C:\\Users\\Nikolay\\Documents\\book.xls')
