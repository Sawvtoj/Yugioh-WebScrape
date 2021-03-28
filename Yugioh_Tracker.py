import requests
from selenium import webdriver
from bs4 import BeautifulSoup
import pandas as pd
import time

start_time = time.time()

df = pd.read_excel('C:\\Users\\Xp\\Documents\\GitHub\\Yugioh-WebScrape\\Yugioh.xlsx')

cdName = df["Card Name"]
cdNumber = df["Card Number"]
cdRarity = df["Rarity"] 

arrayName = cdName.to_numpy()
arrayNumber = cdNumber.to_numpy()
arrayRarity = cdRarity.to_numpy()
arrayPrice = []

driver = webdriver.Chrome()

driver.get('https://shop.tcgplayer.com/yugioh')

driver.implicitly_wait(10)
#len(df)
#i = 781 
for i in range(len(df)):
  
    searchName = driver.find_element_by_id('ProductName')
    searchName.send_keys(arrayName[i])

    searchNumber = driver.find_element_by_id('Number')
    searchNumber.send_keys(arrayNumber[i])
    
    if(arrayRarity[i] == 'Common'):
        searchRarity = driver.find_element_by_xpath("//input[@value='Common / Short Print']").click()
    if(arrayRarity[i] == 'Super'):    
        searchRarity = driver.find_element_by_xpath("//input[@value='Super']").click()
    if(arrayRarity[i] == 'Ultra'):    
        searchRarity = driver.find_element_by_xpath("//input[@value='Ultra']").click()
    if(arrayRarity[i] == 'Secret'):    
        searchRarity = driver.find_element_by_xpath("//input[@value='Secret']").click()
    if(arrayRarity[i] == 'Ultimate'):    
        searchRarity = driver.find_element_by_xpath("//input[@value='Ultimate']").click()
    if(arrayRarity[i] == 'Ghost'):    
        searchRarity = driver.find_element_by_xpath("//input[@value='Ghost']").click()
    if(arrayRarity[i] == 'Rare'):    
        searchRarity = driver.find_element_by_xpath("//input[@value='Rare']").click()
    if(arrayRarity[i] == 'Prismatic Secret Rare'):    
        searchRarity = driver.find_element_by_xpath("//input[@value='Prismatic Secret Rare']").click()
    if(arrayRarity[i] == 'Starlight Rare'):    
        searchRarity = driver.find_element_by_xpath("//input[@value='Starlight Rare']").click()
    
    search = driver.find_element_by_xpath("//input[@value='Search']").click()
    
    productCardURL = driver.current_url

    r = requests.get(productCardURL)

    soup = BeautifulSoup(r.text, "html.parser")
    
    info = soup.find(["strong", "dd"])
    
    if info == soup.find("strong"):
        for m in info:
            if m == "Oh no! Nothing was found!":
                er = "Error"
                arrayPrice.append(er)
                break
            
    if info == soup.find("dd"):
        for n in info:
            
            if n == "Unavailable":
                arrayPrice.append(n.string)
                break
                
            price = n.string
            arrayPrice.append(float(price[1:]))
   
    driver.back()    
    
    clearAll = driver.find_element_by_xpath("//input[@value='Clear All']").click()
    
    #i = i+1
    
    #if i == 784:
    #    break 
    
driver.close()

print(arrayPrice)

cdPrices = pd.DataFrame({'Prices': arrayPrice}) 

df.update(cdPrices)

writer = pd.ExcelWriter(r'C:\Users\Xp\Documents\GitHub\Yugioh-WebScrape\\Yugioh2.xlsx', engine = 'xlsxwriter')

df.to_excel(writer, index = False)

writer.save()

duration = time.time() - start_time

print(f'Time done: {duration}')

#https://pbpython.com/improve-pandas-excel-output.html

