# Library for opening url and creating 
# requests
import urllib.request

# pretty-print python data structures
from pprint import pprint

# for parsing all the tables present 
# on the website
from html_table_parser.parser import HTMLTableParser

# for converting the parsed data in a
# pandas dataframe
import pandas as pd

# Opens a website and read its
# binary contents (HTTP Response Body)
def url_get_contents(url):

    """ Opens a website and read its binary contents (HTTP Response Body) """
    req = urllib.request.Request(url=url)
    f = urllib.request.urlopen(req)
    return f.read()

url = 'https://my.sunysullivan.edu/ICS/CRM_Faculty/Faculty_Page.jnz?portlet=Facility_Schedules'
xhtml = url_get_contents(url).decode('utf-8')

p = HTMLTableParser()
p.feed(xhtml)
pprint(p.tables)