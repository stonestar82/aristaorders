import os, shutil
from datetime import datetime

## 파일 체크
today = datetime.today().strftime("%Y%m%d")
orderFile = "./arista_order_" + today + ".xlsx"
while True:
	idx = 0
	if os.path.isfile(orderFile):
		idx += 1
		orderFile = "./arista_order_" + today + "_" + str(idx) + ".xlsx"
	else:
		break
	
## sample 파일 복사
shutil.copy("./sample.xlsx", orderFile)