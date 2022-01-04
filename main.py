#price of best brands of watches in amazon on an excel sheet fully automated
#modules used:xlwt and selenium

#code:
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from time import sleep
import xlwt
from xlwt import Workbook

options = Options()
options.add_experimental_option('excludeSwitches', ['enable-logging'])

s=Service("C:\Program Files (x86)\chromedriver.exe")

driver=webdriver.Chrome(service=s,options=options)
driver.get("https://amazon.com")

watches=["citizen","fitbit","honorband","amazfit","Titan Watch",
"Rolex Watch","Omega watch","Casio"]

wb=xlwt.Workbook(encoding="utf-8")
sheet1=wb.add_sheet('Sheet 1',cell_overwrite_ok=True)

i=1
wb=Workbook()
sheet1=wb.add_sheet("Sheet 1", cell_overwrite_ok=True)
sheet1.write(0,0,"Watch products")
sheet1.write(0,1,"Price")
try:
      for watch in watches:
          #i=0
          search=driver.find_element(By.ID,"twotabsearchtextbox")
          search.send_keys(watch)

          search.send_keys(Keys.RETURN)
    
          click=driver.find_element(By.CLASS_NAME,"s-image")
          click.click()
       

          price=driver.find_element(By.ID,"price_inside_buybox")
          cost=price.text

          #print(cost)
          driver.find_element(By.ID,"twotabsearchtextbox").clear()
          #i+=1
          sheet1.write(i,0,watch)
          sheet1.write(i,1,cost)
          i+=1
          #print(i)
          sleep(2)
except Exception as e:
       print(Exception)
       print("Sorry the price of this product of this object is not available")

sleep(5.7)
driver.quit()


wb.save('watche_prices.xls')

#final output:
#the output has been converted to html to ensure that the users who dont have office can view at as well
#https://drive.google.com/drive/u/0/folders/1ZO4qMz4-GB3iTboO-w38Z-C5ujdwdrze
#go to the link above 
#download the following file as .html or .html
#and open the file

#if you want to view it in excel run the folowing code in pycharm,atom,vscode or etc.

