# -*- coding: utf-8 -*-
#!/usr/bin/env python
from tkinter import *
from tkinter import messagebox
import platform, threading
import ttkbootstrap as ttk

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

def threadCall(func, **kwargs):
		threading.Thread(target=func, kwargs=kwargs).start()

def resource_path(relative_path):
		try:
		# PyInstaller에 의해 임시폴더에서 실행될 경우 임시폴더로 접근하는 함수
			base_path = sys._MEIPASS
		except Exception:
			base_path = os.path.abspath(".")
		return os.path.join(base_path, relative_path)

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
		
		btn.configure(state="disabled")
		btn.configure(text="진행중")
				
		if eq(platform.system().lower(), "windows"):
			service = ChromeService(executable_path="./chromedriver.exe")
		else:
			service = ChromeService(executable_path=f"{path}/chromedriver")
		
		#service = ChromeService()
		
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
				status = driver.find_element(By.XPATH, '//*[@id="main_form"]/table/tbody/tr[1]/td/div/div[4]/div[3]').get_attribute("innerText")
																									
				if ne("", billToTmp):
					billTo = billToTmp.split("/")[0].split("\n")[0].strip().replace("\"", "")
					
				if ne("", shipToTmp):
					shipTo = shipToTmp.split("/")[0].split("\n")[0].strip().replace("\"", "")
					
				
					
				if eq(d[4].strip(), ""):
					d[4] = status
					print("status = ", status)
				else:
					d[4] = d[4].upper()
					
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
		
		compayName = "globaltelecom" if "globaltelecom" in v1 else "i_cloud"

		## 파일 체크
		orderFile = f"{path}/arista_order_{compayName}_" + today + ".xlsx"
		idx = 0
		while True:
			if os.path.isfile(orderFile):
				idx += 1
				orderFile = f"{path}/arista_order_{compayName}_" + today + "_" + str(idx) + ".xlsx"
			else:
				break
			
		## sample 파일 복사
		shutil.copy(f"{path}/sample.xlsx", orderFile)
		
		workbook = load_workbook(filename=orderFile, read_only=False, data_only=True)
		etcData = []
		
		sheet = workbook["ORDERS"]
		for idx, d in enumerate(data):
			# print(d)
			
			## CURRENT SHIP DATE, RECOMMIT DATE, EXPEDITE DATE 값이 다 없는경우 다른 시트에 작성
			if eq("", d[5].strip()) and eq("", d[6].strip()) and eq("", d[7].strip()):
				etcData.append(d)
			else:
				sheet.append([idx+1, d[0], d[1], d[3], d[4], d[5], d[6], d[7], d[8], d[9], d[10], d[11], d[12], d[13], d[14], d[15]])
			
		sheet = workbook["ITEMS"]
		for idx, d in enumerate(orderItemData):
			sheet.append([d[0], d[1], d[2], d[3], d[4], d[5], d[6], d[7], d[8], d[9], d[10], d[11]])
			
			
		sheet = workbook["ETC ORDERS"]
		for idx, d in enumerate(etcData):
			sheet.append([idx+1, d[0], d[1], d[3], d[4], d[5], d[6], d[7], d[8], d[9], d[10], d[11], d[12], d[13], d[14], d[15]])
				
		workbook.save(orderFile)
		workbook.close()

		btn.configure(state="normal")
		btn.configure(text="시작")
		
		
root = ttk.Window(themename="litera")
root.title("i-Cloud - Arista Orders")
if eq(platform.system().lower(), "windows"):
	root.geometry("400x150+500+300") ## w, h, x, y
	root.iconbitmap(resource_path("icloud.ico"))
else:
	root.geometry("230x150+500+300") ## w, h, x, y
root.resizable(False, False)

label = ttk.Label(root, text="ID :")
label.grid(column=0, row=0, padx=(8, 8), pady=(5, 10))

id = StringVar()
idTextbox = ttk.Entry(root, width=42, textvariable=id)
idTextbox.grid(column = 1, row=0, padx=(0, 8), pady=(10, 10))
idTextbox.focus()

label = ttk.Label(root, text="PW :")
label.grid(column=0, row=1, padx=(8, 8), pady=(5, 10))

pw = StringVar()
pwTextbox = ttk.Entry(root, width=42, textvariable=pw, show="*")
pwTextbox.grid(column = 1 , row = 1, padx=(0, 8), pady=(10, 10))

btn = ttk.Button(root, text="시작", bootstyle="danger", width=51, command=lambda: threadCall(clickMe))
btn.grid(columnspan=2, row=2, padx=(8, 8), pady=(10, 10))

root.mainloop()