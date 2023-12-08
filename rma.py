# -*- coding: utf-8 -*-
#!/usr/bin/env python
from tkinter import *
from tkinter import ttk, messagebox
import platform
import ttkbootstrap as ttk

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select

import pause, os, shutil, sys
from operator import eq, ne
# from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime
import configRma as mc
from time import sleep
from openpyxl import load_workbook
import requests
import threading

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
	#v1 = mc.email
	#v2 = mc.password
	if eq(platform.system().lower(), "windows"):
		path = "."
	else:
		path = os.path.sep.join(sys.argv[0].split(os.path.sep)[:-1])
	
	data = []
	rmaList = []
	
	if (v1 == "" or v2 == ""):
		messagebox.showinfo("Button Clicked", "ID or PW 입력해주세요.")
		idTextbox.focus()
	else:

		btn.configure(state="disabled")
		btn.configure(text="진행중")
		
		if eq(platform.system().lower(), "windows"):
			service = ChromeService()
		else:
			service = ChromeService(executable_path=f"{path}/chromedriver")
		
		options = ChromeOptions()
		options.page_load_strategy = 'normal'
		
		driver = webdriver.Chrome(service=service, options=options)

		wait = WebDriverWait(driver, 10)

		driver.get(mc.loginUrl)

		print("페이지 열림")

		continue_btn = wait.until(EC.element_to_be_clickable((By.ID, "btnLoginOrigin")))

		id_elem = wait.until(EC.element_to_be_clickable((By.ID, "username")))

		id_elem.send_keys(v1)

		continue_btn.click()

		sleep(5)

		pass_elem = driver.find_element(By.ID, "password")

		pass_elem.send_keys(v2)

		login_elem = driver.find_element(By.ID, "btnLogin")

		login_elem.click()


		print("Portal 페이지 이동")
		driver.get(mc.suportUrl)

		sleep(10)


		print("RMA 클릭")

		rmaBtn = driver.execute_script("return document.querySelector('[data-label=rmadetail]')")

		rmaBtn.click()


		sleep(5)

		print("select box 100 선택")

		## select box 선택
		select = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "rma-details-num")))
		select = Select(driver.find_element(By.CLASS_NAME, 'rma-details-num'))
		select.select_by_value('100')

		sleep(3)

		cookieBtn = driver.find_element(By.ID, 'onetrust-accept-btn-handler')

		print("cookieBtn = ", cookieBtn)

		if (cookieBtn.is_displayed()):
			cookieBtn.click()

		# pagination 
		pagination = driver.execute_script("return document.querySelector('.slds-p-around_small .slds-col').querySelectorAll('span')")

		for k in range(0, len(pagination)):

			if k > 0:
				p = pagination[k].find_element(By.TAG_NAME, "a")
				p.click()

				sleep(3)

			listCount = len(driver.find_element(By.XPATH, '//*[@id="firstTab"]/div/div/div[1]/table').find_element(By.TAG_NAME, "tbody").find_elements(By.TAG_NAME, "tr"))


			for x in range(0, listCount):

					tr = driver.find_element(By.XPATH, '//*[@id="firstTab"]/div/div/div[1]/table').find_element(By.TAG_NAME, "tbody").find_elements(By.TAG_NAME, "tr")[x]
					
					rma = tr.find_elements(By.TAG_NAME, "th")[0].find_element(By.TAG_NAME, "div").find_element(By.TAG_NAME, "a")

					rmaNo = rma.get_attribute("innerText")

					print("rmaNo = ", rmaNo)

					tds = tr.find_elements(By.TAG_NAME, "td")
					rmaTitle = tds[0].get_attribute("innerText")
					rmaDate = tds[1].get_attribute("innerText")
					accountName = tds[2].get_attribute("innerText")
					endCustomer = tds[3].get_attribute("innerText")


					rmaList.append({"RMA No": rmaNo, "RMA CASE TITLE": rmaTitle, "RMA DATE": rmaDate, "ACCOUNT NAME": accountName, "END CUSTOMER": endCustomer})


					#sleep(2)
					driver.execute_script(f'document.querySelector("a[name={rmaNo}]").focus()')
					rma.click()

					sleep(3)

					trCount = len(driver.find_element(By.XPATH, '//*[@id="modal-content-id-1"]/table').find_element(By.TAG_NAME, "tbody").find_elements(By.TAG_NAME, "tr"))


					for i in range(0, trCount):
							
						rmaTable = driver.find_element(By.XPATH, '//*[@id="modal-content-id-1"]/table')

						# rma tbody
						rmaTbody = rmaTable.find_element(By.TAG_NAME, "tbody")
						rmaTr = rmaTbody.find_elements(By.TAG_NAME, "tr")[i]

						more = rmaTr.find_elements(By.TAG_NAME, "td")[2].find_element(By.TAG_NAME, "a")
						#more.
						more.click()

						d = {"RMA CASE TITLE": rmaTitle, "RMA DATE": rmaDate, "ACCOUNT NAME": accountName, "END CUSTOMER": endCustomer, "Switch Serial Number": "", "SR Case": "", "Defective Part Number": "", "Required Delivery Method": "", "Defective Serial Number": "", "Onsite FE Required": "", "Replacement Part Number": "", "Delivery Status": "", "Replacement Serial Number": "", "Carrier": "", "Ship to Company Name": "", "Carrier Method": "", "Ship to Contact Name": "", "Carrier Tracking Number": "", "Ship to Contact Number": "", "Ship Date": "", "Ship to Address": "", "ETA": "", "Ship to City": "", "Delivery Date/Time": "", "Ship to State/County": "", "Proof of Delivery Name": "", "Ship to Country": "", "Ship to Zip Code": ""}

						#sleep(3)
						
						close = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="firstTab"]/section/div/footer/button')))

						dataTable = driver.find_element(By.XPATH, '//*[@id="modal-content-id-1"]/table')

						# data tbody
						dataTbody = dataTable.find_element(By.TAG_NAME, "tbody")
						for dataTr in dataTbody.find_elements(By.TAG_NAME, "tr"):
							tds = dataTr.find_elements(By.TAG_NAME, "td")
							#print(tds[0].get_attribute("innerText").strip(), " = ", tds[1].get_attribute("innerText").strip())
							d[tds[0].get_attribute("innerText").strip()] = tds[1].get_attribute("innerText").strip()

							if len(tds) == 4:
								#print(tds[2].get_attribute("innerText").strip(), " = ", tds[3].get_attribute("innerText").strip())
								d[tds[2].get_attribute("innerText").strip()] = tds[3].get_attribute("innerText").strip()

						data.append(d)

						close.click()

					## 데이터가 없는경우
					if (trCount == 0):
						d = {"RMA No": rmaNo, "RMA CASE TITLE": rmaTitle, "RMA DATE": rmaDate, "ACCOUNT NAME": accountName, "END CUSTOMER": endCustomer, "Switch Serial Number": "", "SR Case": "", "Defective Part Number": "", "Required Delivery Method": "", "Defective Serial Number": "", "Onsite FE Required": "", "Replacement Part Number": "", "Delivery Status": "", "Replacement Serial Number": "", "Carrier": "", "Ship to Company Name": "", "Carrier Method": "", "Ship to Contact Name": "", "Carrier Tracking Number": "", "Ship to Contact Number": "", "Ship Date": "", "Ship to Address": "", "ETA": "", "Ship to City": "", "Delivery Date/Time": "", "Ship to State/County": "", "Proof of Delivery Name": "", "Ship to Country": "", "Ship to Zip Code": ""}
						data.append(d)

					close = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="firstTab"]/section/div/footer/button')))

					close.click()

					#sleep(2)


		driver.close()
		## 오늘 날짜
		today = datetime.today().strftime("%Y%m%d")
		
		compayName = "globaltelecom" if "globaltelecom" in v1 else "i_cloud"
		
		## 파일 체크
		orderFile = f"{path}/arista_rma_{compayName}_" + today + ".xlsx"
		idx = 0
		while True:
			if os.path.isfile(orderFile):
				idx += 1
				orderFile = f"{path}/arista_rma_{compayName}_" + today + "_" + str(idx) + ".xlsx"
			else:
				break
			
		## sample 파일 복사
		shutil.copy(f"{path}/sample_rma.xlsx", orderFile)
		
		workbook = load_workbook(filename=orderFile, read_only=False, data_only=True)

		sheet = workbook["RMA"]
		for idx, d in enumerate(rmaList):
			print(d)

			sheet.append([idx+1, d["RMA No"], d["RMA CASE TITLE"], d["RMA DATE"], d["ACCOUNT NAME"], d["END CUSTOMER"]])


		
		sheet = workbook["RMA Item"]
		for idx, d in enumerate(data):
			print(d)

			sheet.append([idx+1, d["RMA No"], d["RMA CASE TITLE"], d["RMA DATE"], d["ACCOUNT NAME"], d["END CUSTOMER"], d["Switch Serial Number"], d["SR Case"], d["Defective Part Number"], d["Required Delivery Method"], d["Defective Serial Number"], d["Onsite FE Required"], d["Replacement Part Number"], d["Delivery Status"], d["Replacement Serial Number"], d["Carrier"], d["Ship to Company Name"], d["Carrier Method"], d["Ship to Contact Name"], d["Carrier Tracking Number"], d["Ship to Contact Number"], d["Ship Date"], d["Ship to Address"], d["ETA"], d["Ship to City"], d["Delivery Date/Time"], d["Ship to State/County"], d["Proof of Delivery Name"], d["Ship to Country"], d["Ship to Zip Code"]])
			
		workbook.save(orderFile)
		workbook.close()

		btn.configure(state="normal")
		btn.configure(text="시작")
	
root = ttk.Window(themename="litera")
root.title("i-Cloud - Arista RMA")
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