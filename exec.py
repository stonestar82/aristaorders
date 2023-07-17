# -*- coding: utf-8 -*-
#!/usr/bin/env python
from tkinter import *
from tkinter import ttk, messagebox
import platform

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.support import expected_conditions as EC
import pause, os, shutil, sys
from operator import eq, ne
# from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime
import config as mc
from time import sleep
from openpyxl import load_workbook

def clickMe():
  v1 = id.get()
  v2 = pw.get()
  # v1 = mc.email
  # v2 = mc.password
  if eq(platform.system().lower(), "windows"):
    path = "."
  else:
    path = os.path.sep.join(sys.argv[0].split(os.path.sep)[:-1])
  orderUrlList = []
  
  data = []
  
  if (v1 == "" or v2 == ""):
    messagebox.showinfo("Button Clicked", "ID or PW 입력해주세요.")
    idTextbox.focus()
  else:
    
    service = ChromeService(executable_path=f"{path}/chromedriver.exe")
    
    options = ChromeOptions()
    options.page_load_strategy = 'normal'
    
    driver = webdriver.Chrome(service=service, options=options)

    wait = WebDriverWait(driver, 10)

    driver.get(mc.loginUrl)

    # print("페이지 열림")
    id_elem = wait.until(EC.element_to_be_clickable((By.ID, "email")))
    pass_elem = driver.find_element(By.ID, "password")

    id_elem.send_keys(v1)
    pass_elem.send_keys(v2)

    # print("id/pw")
    # login_elem = driver.find_element(By.CLASS_NAME, "w-100.btn.btn-lg.btn-primary")
    login_elem = driver.find_element(By.ID, "login-submit")

    login_elem.click()

    # print("로그인 처리")


    # 주문 페이지 이동
    driver.get(mc.ordersUrl)


    nextBtn = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "navig-next")))
    
    
    nextBtn = driver.find_element(By.CLASS_NAME, "navig-next")
    
    iframe = driver.find_element(By.ID, "list")
    orderUrlList.append(iframe.get_attribute("src"))
    
    
    # table
    table = driver.find_element(By.ID, "div__bodytab")
    

    # tbody
    tbody = table.find_element(By.TAG_NAME, "tbody")
    for tr in tbody.find_elements(By.TAG_NAME, "tr")[1:-1]:
        tmpData = []
        for idx, td in enumerate(tr.find_elements(By.TAG_NAME, "td")[1:]):
            tmpData.append(td.get_attribute("innerText"))
            
            if idx == 1:
              f = td.find_element(By.TAG_NAME, "a").get_attribute("href")
              tmpData.append(f)
              
            
        data.append(tmpData)
    
    
    while True:
      try:
        
        next = driver.find_element(By.XPATH, '//*[@id="segment_fs"]/span[2]/a[2]/span/span')
        t = next.get_attribute("innerText")
        
        # print("span text", t)
        
        if (t.strip() != ""):
          nextBtn.click()
        
          sleep(3)
        
          table = driver.find_element(By.ID, "div__bodytab")
          tbody = table.find_element(By.TAG_NAME, "tbody")
          for tr in tbody.find_elements(By.TAG_NAME, "tr")[1:-1]:
              tmpData = []
              for idx, td in enumerate(tr.find_elements(By.TAG_NAME, "td")[1:]):
                tmpData.append(td.get_attribute("innerText"))
              
                if idx == 1:
                  f = td.find_element(By.TAG_NAME, "a").get_attribute("href")
                  tmpData.append(f)
              
              data.append(tmpData)
        else:
          break
      except:
        break
    
    ## 상세페이지 정보 확인
    orderItemData = []
    for d in data:
      url = d[2]
      
      # print("url", url)

      if ne("", url):
        driver.get(url)
        
        po = driver.find_element(By.XPATH, '//*[@id="detail_table_lay"]/tbody/tr/td[2]/table/tbody/tr[4]/td/div/span[2]').get_attribute("innerText")
        orderNo = driver.find_element(By.XPATH, '//*[@id="detail_table_lay"]/tbody/tr/td[1]/table/tbody/tr[6]/td/div/span[2]').get_attribute("innerText")
        orderDate = driver.find_element(By.XPATH, '//*[@id="detail_table_lay"]/tbody/tr/td[1]/table/tbody/tr[5]/td/div/span[2]').get_attribute("innerText")
        saleRep = driver.find_element(By.XPATH, '//*[@id="detail_table_lay"]/tbody/tr/td[2]/table/tbody/tr[2]/td/div[1]/span[2]/span').get_attribute("innerText")
        terms = driver.find_element(By.XPATH, '//*[@id="detail_table_lay"]/tbody/tr/td[2]/table/tbody/tr[1]/td/div/span[2]/span').get_attribute("innerText")
        billToTmp = driver.find_element(By.XPATH, '//*[@id="address_form"]/table/tbody/tr/td/table/tbody/tr/td[1]/table/tbody/tr/td/div/span[2]').get_attribute("innerText")
        shipToTmp = driver.find_element(By.XPATH, '//*[@id="address_form"]/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr[1]/td/div/span[2]').get_attribute("innerText")
                                                  
        if ne("", billToTmp):
          billTo = billToTmp.split("/")[0].split("\n")[0].strip().replace("\"", "")
          
        if ne("", shipToTmp):
          shipTo = shipToTmp.split("/")[0].split("\n")[0].strip().replace("\"", "")
          
          
        # printBtn = nextBtn = wait.until(EC.element_to_be_clickable((By.ID, "tdbody_print")))
        
        ## End of Group 
        
        table = driver.find_element(By.ID, "item_splits")
        tbody = table.find_element(By.TAG_NAME, "tbody")
        for tr in tbody.find_elements(By.TAG_NAME, "tr")[1:]:
          tmpData = []
          tmpData.append(po)
          tmpData.append(orderNo)
          for idx, td in enumerate(tr.find_elements(By.TAG_NAME, "td")):
            tmpData.append(td.get_attribute("innerText").strip())
          
          tmpData.append(orderDate)
          tmpData.append(terms)
          tmpData.append(saleRep)
          tmpData.append(billTo)
          tmpData.append(shipTo)
          
          # print("order item data ", tmpData)
            
          orderItemData.append(tmpData)

    driver.close()
    ## 오늘 날짜
    today = datetime.today().strftime("%Y%m%d")
    
    
    ## 파일 체크
    orderFile = f"{path}/arista_order_" + today + ".xlsx"
    idx = 0
    while True:
      if os.path.isfile(orderFile):
        idx += 1
        orderFile = f"{path}/arista_order_" + today + "_" + str(idx) + ".xlsx"
      else:
        break
      
    ## sample 파일 복사
    shutil.copy(f"{path}/sample.xlsx", orderFile)
    
    workbook = load_workbook(filename=orderFile, read_only=False, data_only=True)
    sheet = workbook["ORDERS"]
  


    for idx, d in enumerate(data):
      # print(d)
      sheet.append([idx+1, d[0], d[1], d[3], d[4], d[5], d[6], d[7], d[8], d[9], d[10], d[11], d[12], d[13], d[14], d[15]])
    
    sheet = workbook["ITEMS"]
    for idx, d in enumerate(orderItemData):
      sheet.append([d[0], d[1], d[2], d[3], d[4], d[5], d[6], d[7], d[8], d[9], d[10], d[11]])
      
    workbook.save(orderFile)
    workbook.close()
    
    
  
root = Tk()
root.title("i-Cloud - Arista Orders")
if eq(platform.system().lower(), "windows"):
  root.geometry("200x150+500+300") ## w, h, x, y
else:
  root.geometry("230x150+500+300") ## w, h, x, y
root.resizable(False, False)


label = ttk.Label(root, text="ID")
label.grid(column=0, row=0)

id = StringVar()
idTextbox = ttk.Entry(root, width=20, textvariable=id)
idTextbox.grid(column = 1 , row = 0)
idTextbox.focus()

label = ttk.Label(root, text="PW")
label.grid(column=0, row=1)

pw = StringVar()
pwTextbox = ttk.Entry(root, width=20, textvariable=pw, show="*")
pwTextbox.grid(column = 1 , row = 1)

btn = Button(root, text="시작", padx=10, pady=5, command=clickMe)
btn.grid(column=1, row=2)
# btn = Button(root, text="버튼", width=10, height=3)

# btn.pack()
root.mainloop()