# -*- coding: utf-8 -*-
"""
Created on Thu Jun 18 04:55:47 2020

@author: User
"""

import selenium.webdriver
import pandas as pd

#url='https://www.w3schools.com/html/html_tables.asp'
url='https://intranet-grid.hms.se/intra/scala/pl/pl_pending.php?PL01001=6679&scco=HS'

driver = selenium.webdriver.Chrome()
driver.get(url)
all_tables=pd.read_html(driver.page_source, attrs={'class': 'sellist'})
df = all_tables[0]
print(df)
row=len(df)


driver.close()
#get all href in a webpage
"""
elems = driver.find_elements_by_xpath("//a[@href]")
for elem in elems:
    print(elem.get_attribute("href"))
"""
all_Tables=[]
Df=[]
for x in range(row-1):
    all_Tables.append(x)
    Df.append(x)
    xpth='//*[@id="row_'
    xpth+=str(x)
    xpth+='"]/td[1]'
    driver = selenium.webdriver.Chrome()
    driver.get(url)
    driver.find_element_by_xpath(xpth).click()
    all_Tables[x]=pd.read_html(driver.page_source, attrs={'class': 'sellist'})
    Df[x] = all_Tables[x]
    print(Df[x])
    driver.close()

m=0
for x in range(row-1):
    df1=Df[x]
    y=df1[0]
    skp=len(y)
    rp=0
    m=m+1
    for z in range(1,skp+1):
        line = pd.DataFrame({}, index=[z])
        df = pd.concat([df.iloc[:z+m], line, df.iloc[z+m:]]).reset_index(drop=True)
    m=m+skp
    


writer = pd.ExcelWriter(r'Data.xlsx', engine='xlsxwriter')
rp=0
for x in range(row-1):
    df1=Df[x]
    y=df1[0]
    if x==0:
        rp=1
    else:
        rp=rp+skp+1
    # Position the dataframes in the worksheet.
    y.to_excel(writer, sheet_name='Sheet1', startrow=rp, startcol=7,header=False, index=False)  # Default position, cell A1.    
    skp=len(y)

df.to_excel(writer, sheet_name='Sheet1', header=False, index=False)
writer.save()

