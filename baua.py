from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
import os
import xlsxwriter
import time
import sys

URL="https://www.baua.de/DE/Biozid-Meldeverordnung/Offen/offen.html"

PRODUCTSITE_URLS = []
PAGE_URLS = []


try:
  workbook = xlsxwriter.Workbook(sys.argv[2])
except:
  print("File not created: ", sys.argv[2])
  try:
    workbook = xlsxwriter.Workbook("results.xlsx")
  except:
    print("Provide a name for results.xlsx")
    sys.exit()

# create excel file

worksheet = workbook.add_worksheet()
row = 0
col = 0

# prepare chrome 
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--window-size=1024x1400")
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--window-size=1024x1400")
chrome_driver = os.path.join(os.getcwd(), "chromedriver")
driver = webdriver.Chrome(options=chrome_options, executable_path=chrome_driver)
driver2 = webdriver.Chrome(options=chrome_options, executable_path=chrome_driver)



# create excel headers
worksheet.write(row,col, "Handelname" )
worksheet.write(row,col+1, "RegNr" )
worksheet.write(row,col+2, "MaldeDatum" )
worksheet.write(row,col+3, "Wirktstoff" )
worksheet.write(row,col+4, "CasNr" )
worksheet.write(row,col+5, "EcNr" )
worksheet.write(row,col+6, "PT" )
worksheet.write(row,col+7, "FirmName" )
worksheet.write(row,col+8, "FirmAddr" )
worksheet.write(row,col+9, "FirmLand" )
row += 1



# # open URL (site = 1)
# driver.get(URL)

# # breaks when cannot click on "next" button
# click=3128
# while True:
#   # get all URLs of products on page (driver #1)
#   try:
#     entrys = driver.find_elements_by_xpath("//*[@id='produkteDatatable']/tbody/tr[*]/td[5]/a")
#   except:
#     print("one missed")
#   for e in entrys:
#     try:
#       PRODUCTSITE_URLS.append(e.get_attribute('href'))
#     except:
#       print("no such elem")
#   x=0
#   # click "next" button
#   while x < 10:
#     try:
#       driver.find_element_by_id("produkteDatatable_next").click()
#       print("Click2go: ", click)
#       x=100
#       time.sleep(1)
#       click-=1
#       break
#     except:
#       print("!")
#       x+=2
#       time.sleep(1)
#   if x == 10 or click < 1:
#     break

# remove URLs with duplicate products
#PRODUCTSITE_URLS = list(dict.fromkeys(PRODUCTSITE_URLS))


# f = open("uniq-product-URLs.txt", "a")
# for url in PRODUCTSITE_URLS:
#   f.write(url)

# f.close()

try:
  f = open(sys.argv[1], "r")
except:
  print("File not found: ", sys.argv[1])
  try:
    f = open("URLS.txt")
  except:
    print("Provide a file name as argument or place URLS.txt ")
    sys.exit()

LINES = f.read().split("\n")
for l in LINES:
  PRODUCTSITE_URLS.append(l)

print("List has: ", len(PRODUCTSITE_URLS), " uniqs")
# get product page (driver #2)
for u in PRODUCTSITE_URLS:
  # get product page
  c=5
  # retry geting page until success
  while c > 0:
    try:
      #print("working on ", u, 'c=',c)
      c-=1
      driver2.get(u)
    except Exception as e:
      print("Retry " , c, " on: ", u)
      #print("message:", str(e))
  try:
    # extract product
    Handelname = driver2.find_element_by_xpath("//*[@id='content']/div/div/div/div[1]/table/tbody/tr[1]/td").text
    RegNr = driver2.find_element_by_xpath("//*[@id='content']/div/div/div/div[1]/table/tbody/tr[2]/td").text
    MaldeDatum = driver2.find_element_by_xpath("//*[@id='content']/div/div/div/div[1]/table/tbody/tr[3]/td").text
    wCount = 0
    # //*[@id="content"]/div/div/div/div[2]/table/tbody/tr[1]/td
    Wirktstoff = driver2.find_element_by_xpath(str('//*[@id="content"]/div/div/div/div[2]/table/tbody/tr[1]/td')).text
    # //*[@id="content"]/div/div/div/div[2]/table/tbody/tr[2]/td
    CasNr = driver2.find_element_by_xpath(str('//*[@id="content"]/div/div/div/div[2]/table/tbody/tr[2]/td')).text
    # //*[@id="content"]/div/div/div/div[2]/table/tbody/tr[3]/td
    EcNr = driver2.find_element_by_xpath(str('//*[@id="content"]/div/div/div/div[2]/table/tbody/tr[3]/td')).text
    # //*[@id="content"]/div/div/div/div[2]/table/tbody/tr[4]/td
    PT = driver2.find_element_by_xpath(str('//*[@id="content"]/div/div/div/div[2]/table/tbody/tr[4]/td')).text      
    if len(driver2.find_elements_by_class_name('nextEntry')) > 0:
      NextWirkts = len(driver2.find_elements_by_class_name('nextEntry'))
      limit = NextWirkts / 2
      i=1
      #Wirktstoff = ""
      while i <= limit:
        xp = str('//*[@id="content"]/div/div/div/div[2]/table/tbody/tr[XXX]/td').replace('XXX', str(i  * 5 + 1))
        Wirktstoff += "\n" + driver2.find_element_by_xpath(xp).text
        xp = str('//*[@id="content"]/div/div/div/div[2]/table/tbody/tr[XXX]/td').replace('XXX', str(i  * 5 + 2))
        CasNr += "\n" + driver2.find_element_by_xpath(xp).text
        xp = str('//*[@id="content"]/div/div/div/div[2]/table/tbody/tr[XXX]/td').replace('XXX', str(i  * 5 + 3))
        EcNr += "\n" + driver2.find_element_by_xpath(xp).text
        xp = str('//*[@id="content"]/div/div/div/div[2]/table/tbody/tr[XXX]/td').replace('XXX', str(i  * 5 + 4))
        PT += "\n" + driver2.find_element_by_xpath(xp).text
        i+=1 

    FirmName = driver2.find_element_by_xpath('//*[@id="content"]/div/div/div/div[3]/table/tbody/tr[1]/td').text
    FirmAddr = driver2.find_element_by_xpath('//*[@id="content"]/div/div/div/div[3]/table/tbody/tr[2]/td').text
    FirmLand = driver2.find_element_by_xpath('//*[@id="content"]/div/div/div/div[3]/table/tbody/tr[3]/td').text
  except:
    print('Missing section')

  #add product to excel
  worksheet.write(row,col, Handelname )
  worksheet.write(row,col+1, RegNr )
  worksheet.write(row,col+2, MaldeDatum )
  worksheet.write(row,col+3, Wirktstoff )
  worksheet.write(row,col+4, CasNr )
  worksheet.write(row,col+5, EcNr )
  worksheet.write(row,col+6, PT )
  worksheet.write(row,col+7, FirmName )
  worksheet.write(row,col+8, FirmAddr )
  worksheet.write(row,col+9, FirmLand )
  row += 1
  print("Added excel row: ", row)




# save all results at once

driver.close()
driver2.close()
workbook.close()