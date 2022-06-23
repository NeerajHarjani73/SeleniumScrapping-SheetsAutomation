import time
import pandas as pd
import os
import numpy as np
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from datetime import datetime, timedelta
yesterday = datetime.now() - timedelta(3)
yesterday = datetime.strftime(yesterday, '%m/%d/%y')
if os.path.exists("/Users/neerajharjani/Downloads/data.csv"):
  os.remove("/Users/neerajharjani/Downloads/data.csv")
else:
  print("The file does not exist")
#from selenium.webdriver.support.ui import WebDriverWait
#from selenium.webdriver.support import expected_conditions as EC
#from selenium.webdriver.common.by import By
#from selenium.webdriver.support.ui import Select
#import path
#from selenium.webdriver.chrome.options import Options

op = webdriver.ChromeOptions()
op.add_argument('headless')
driver = webdriver.Chrome("/Users/neerajharjani/Downloads/chromedriver 3")
driver.get("URL")

time.sleep(1)

email = driver.find_element_by_xpath('XPATH.../div[1]/div[1]/div/div/input')
email.clear()
email.send_keys("Email")

password = driver.find_element_by_xpath('XPATH.../div[1]/div[2]/div/div/input')
password.clear()
password.send_keys("Password")

time.sleep(1)

driver.find_element_by_xpath('XPATH.../div[2]/div/div/button').click()

time.sleep(5)

driver.find_element_by_xpath('XPATH...[3]/a').click()

time.sleep(1)

driver.find_element_by_xpath('XPATH...[1]/div[2]/button').click()

time.sleep(1)

driver.find_element_by_xpath('XPATH..custom-table-view-selector"]/div[3]/div[2]').click()

time.sleep(1)

driver.find_element_by_xpath('XPATH...[2]/div/div[3]/div').click()

#driver.find_element_by_xpath('XPATH..This Year"]').click()

driver.find_element_by_xpath('XPATH..custom-date-range"]').click()
driver.find_element_by_xpath('XPATH..start-date-input"]').send_keys(yesterday)
driver.find_element_by_xpath('XPATH..end-date-input"]').send_keys(yesterday)

driver.find_element_by_xpath('XPATH..end-date-input"]').send_keys(Keys.RETURN)
driver.find_element_by_xpath('XPATH..apply-date-range-button"]').click()

time.sleep(5)

driver.find_element_by_xpath('XPATH...[3]/div/div[1]/div[2]/div/div/button[2]').click()

time.sleep(5)

driver.quit()

df = pd.read_csv("/Users/neerajharjani/Downloads/data.csv")

df = df[['Product Name','Column Name','Column Name','Column Name','Column Name','Column Name']]
df['Column Name'] = df['Column Name'].replace(np.nan,0)
df = df.sort_values(by='Product Name', ascending = False)
df = df.reset_index(drop=True)
df = pd.DataFrame(np.vstack([df.columns, df]))
l1 = df.values.tolist()

import pygsheets 
gc = pygsheets.authorize(service_file='name.json')
sh = gc.open('NAME')
sh.add_worksheet(yesterday)
wks = sh.worksheet_by_title(yesterday)
for data in l1:
    wks.append_table(data, start='A1', end=None, dimension='ROWS',overwrite=False)

wk1 = sh.worksheets()[-1]
temp_today = pd.DataFrame(wk1.get_all_records())

wk1 = sh.worksheets()[-2]
temp_yesterday = pd.DataFrame(wk1.get_all_records())

temp11 = temp_today[temp_today[['Column Name','Column Name']]['Column Name']<90][['Column Name','Column Name']]

wk1 = sh.worksheets()[0]
wk1.clear('A2:Q50')
wk1.set_dataframe(temp11,'A2')

temp_units = pd.DataFrame()
temp_units = temp_yesterday[['Column Name','Column Name']]
temp_units = temp_units.rename(columns={'Column Name': 'Column Name Yesterday'})
temp_units['Column Name Today'] = temp_today['Column Name']
temp_units['Difference'] = temp_units['Column Name Today'] - temp_units['Column Name Yesterday']
temp_units[['Column Name Yesterday', 'Column Name Today', 'Difference']] = temp_units[['Column Name Yesterday', 'Column Name Today', 'Difference']].astype(float)
temp_units.loc['Total']= temp_units[['Column Name Yesterday', 'Column Name Today', 'Difference']].sum()

wk1.set_dataframe(temp_units,'D2')

temp_Column Name = pd.DataFrame()
temp_Column Name = temp_yesterday[['Column Name','Column Name']]
temp_Column Name = temp_Column Name.rename(columns={'Column Name': 'Column Name Yesterday'})
temp_Column Name['Column Name Today'] = temp_today['Column Name']
temp_Column Name['Difference'] = temp_Column Name['Column Name Today'] - temp_Column Name['Column Name Yesterday']
temp_Column Name[['Column Name Yesterday', 'Column Name Today', 'Difference']] = temp_Column Name[['Column Name Yesterday', 'Column Name Today', 'Difference']].astype(float)
temp_Column Name.loc['Total']= temp_Column Name[['Column Name Yesterday', 'Column Name Today', 'Difference']].sum()
wk1.set_dataframe(temp_Column Name,'I2')

temp_Column Name1 = pd.DataFrame()
temp_Column Name1 = temp_yesterday[['Column Name','Column Name']]
temp_Column Name1 = temp_Column Name1.rename(columns={'Column Name': 'Column Name Yesterday'})
temp_Column Name1['Column Name Today'] = temp_today['Column Name']
temp_Column Name1 = temp_Column Name1[temp_Column Name1["Column Name"].str.contains("S-FLOAT-001-FBM|GY-8C7I-7C6D-FBA|A-SPIKEBRITE-001|A-SPIKEBRITE-001-USFBM|S-ROOK-001-FBM|S-PRO-001-US-FBM|S-PRO-002|S-SS-001-US-FBM|S-CM-002-AMZUS-FBM")==False]
temp_Column Name1['Difference'] = temp_Column Name1['Column Name Today'] - temp_Column Name1['Column Name Yesterday']
temp_Column Name1[['Column Name Yesterday', 'Column Name Today', 'Difference']] = temp_Column Name1[['Column Name Yesterday', 'Column Name Today', 'Difference']].astype(float)
temp_Column Name1.loc['Total']= temp_Column Name1[['Column Name Yesterday', 'Column Name Today', 'Difference']].sum()
wk1.set_dataframe(temp_Column Name1,'N2')