import time

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
import pandas as pd
import requests
from openpyxl import load_workbook



name=input('Enter ZIP you want to search for: ')
time.sleep(10)
driver = webdriver.Chrome(r"C:\Users\LENOVO\Desktop\BeautifulSoup_webExtract\chromedriver-win64\chromedriver.exe")

# Navigate to the URL
url = "https://vfm-makler.de/wendlingen/suche/"
driver.get(url)

cookie_button = driver.find_element(By.XPATH, '//*[@id="klaro"]/div/div/div[2]/div[3]/div/button[2]')  # You should inspect the website's HTML to find the actual text or attributes
if cookie_button:
    cookie_button.click()

# search_box = driver.find_element("tx_vfmmakler_maklerlisting[search][subject]")  # Replace "q" with the actual name of the search box element

search_box = driver.find_element(By.XPATH, '//*[@id="c7924"]/div/form/div[3]/div/input')

# Enter the search query in the search box
search_query = name
search_box.send_keys(search_query)

search_box.send_keys(Keys.RETURN)



driver.implicitly_wait(20)  # Wait for up to 10 seconds
d={'Agency Name':[],
   'Street':[],
   'Zip Code':[],
   'City':[],
   'Telephone':[],
   'Telefax':[],
   'Mail':[],
   'Website':[]}


links={
    'agency':[],
    'link':[],
    'website':[]
}

# Now you can interact with the page after the search results load
# soup = BeautifulSoup(driver, 'html.parser')
agency_name = driver.find_elements(By.CLASS_NAME, 'title')
# print(some_element[1].text)
if agency_name:
    for i in range(len(agency_name)):
        a_name=agency_name[i].text
        # print(a_name)
        d['Agency Name'].append(a_name)
        links['agency'].append(a_name)

page_source = driver.page_source
soup = BeautifulSoup(page_source, 'html.parser')


div_elements = soup.find_all('div', class_='alert alert-searchresult')
for div in div_elements:
    a=div.find_next('a').get('href')
    agency_website=f'https://vfm-makler.de{a}'
    team_link=f'https://vfm-makler.de{a}ueber-uns/#team'
    links['website'].append(agency_website)
    links['link'].append(team_link)
    p=div.find_next('p').get_text()
    p=p.split('\n')
    street=p[1].strip()
    d['Street'].append((street))
    city=p[2].strip()
    city=city.split(' ')
    d['Zip Code'].append(city[0])
    d['City'].append(city[1])
    # print(city)
    telephone_element = div.find('span', class_='type1')

    if (telephone_element):
        telephone = telephone_element.find_next('span')
        data = telephone.get_text().strip()
        d['Telephone'].append(data)
    else:
        d['Telephone'].append('NA')

    telefax_element = div.find('span', class_='type2')
    if(telefax_element):
        telefax = telefax_element.find_next('span')
        data = telefax.get_text().strip()
        d['Telefax'].append(data)
    else:
        d['Telefax'].append('NA')

    email_element = div.find('span', class_='type4')
    if(email_element):
        email = email_element.find_next('a').get_text('href').strip()
        d['Mail'].append(email)
    else:
        d['Mail'].append('NA')


    website_element = div.find('span', class_='type5')
    if(website_element):
        web = website_element.find_next('a').get_text('href').strip()
        d['Website'].append(web)
    else:
        d['Website'].append('NA')


df_main=pd.DataFrame(d)
file_name_main=f'agency_{search_query}.xlsx'
df_main.to_excel(file_name_main,index=False)
# print(df)
book = load_workbook(file_name_main)
writer = pd.ExcelWriter(file_name_main, engine='openpyxl', mode='a')
writer.book = book

c=1
# print(links)
for i in range(len(links['agency'])):
    team_d={
        'Name':[],
        'Qualification':[],
        'Telephone':[],
        'Telefax':[],
        'Mail':[],
        'Website':[]
    }
    agency_=links['agency'][i]
    l=links['link'][i]
    website_ag=links['website'][i]
    response = requests.get(l)
    soup=BeautifulSoup(response.text,'html.parser')
    div_elements = soup.find_all('div', class_='member')
    for div in div_elements:
        team_d['Website'].append(website_ag)
        name=div.find('div',class_='name').get_text()
        team_d['Name'].append(name)

        qualification=div.find_next('p') #.get_text()
        formatted_text = ';'.join(qualification.stripped_strings)

        print(formatted_text)
        if(qualification):
            team_d['Qualification'].append(formatted_text)

        telephone_element = div.find('span', class_='phone d-block')

        if (telephone_element):
            telephone = telephone_element.find_next('span',class_='communication-val')
            data = telephone.get_text()
            team_d['Telephone'].append(data)
        else:
            d['Telephone'].append('NA')

        telefax_element = div.find('span', class_='fax d-block')
        if (telefax_element):
            telefax = telefax_element.find_next('span', class_='communication-val')
            data = telefax.get_text()
            team_d['Telefax'].append(data)
        else:
            team_d['Telefax'].append('NA')

        email_element = div.find('span', class_='email d-block')
        if (email_element):
            email = email_element.find_next('span', class_='communication-val')
            data = email.get_text()
            team_d['Mail'].append(data)
        else:
            d['Mail'].append('NA')

    team_df=pd.DataFrame(team_d)

    fname=f'team{c}_{agency_}.xlsx'
    team_df.to_excel(writer, sheet_name=fname, index=False)
    writer.save()
    # print(team_d)
    c+=1
    # team_df.to_excel(fname, index=False)

print('Successful Extraction ........')
# Close the browser window when done
driver.quit()


