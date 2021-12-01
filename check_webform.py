from bs4 import BeautifulSoup
#from requests_html import HTMLSession
#from urllib.parse import urljoin
import pprint

exampleFile = open(r'C:\Users\docha\OneDrive\Leka\PPA LTR\E-taotluskeskkond01.html')
soup = BeautifulSoup(exampleFile, 'html.parser')
#class_='form-control'
pprint.pprint(soup.find_all(class_='ng-pristine'))

list_classes = soup.find_all(class_='ng-pristine')
for cl in list_classes:
    print(cl.get('name'))
    #pprint.pprint(cl.attrs)

#attr_dict = soup.find('input').attrs
#pprint.pprint(attr_dict)


list_of_selects = soup.find('select').find_all('option')
for ls in list_of_selects:
    print(ls.get('value'), ls.get_text())

