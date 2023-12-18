#-*- coding:utf-8 -*-
import shutil, os, platform, sys
from operator import eq
from datetime import datetime, timedelta
import subprocess, py_compile


# kkk = "83BC0105E48BF7878DE20DA02A5AF235"
kkk = "83BC0105E48BF787"
      
if eq(platform.system().lower(), "windows"):
  path = "."
  icon = "icloud.ico"
  cmd = f'pyinstaller -F --icon=./{icon} --add-data="{icon};." --key={kkk} --distpath="release" -w -n=pmAutomation exec.py'
else:
  ## mac에서는 .ico는 인식안됨. -w 옵션이 들어가야 실행파일에 아이콘이 들어감
  path = os.path.sep.join(sys.argv[0].split(os.path.sep)[:-1])
  icon = "icloud.icns"
  cmd = f'pyinstaller -F --icon={icon} --add-data="{icon}:." --add-data="icloud.png:." --key={kkk} --distpath="release" -n=pmAutomation exec.py'
  
releaseFolder = f"{path}/release"
sampleXlsx = f"{path}/sample.xlsx"
toSampleXlsx = f"{releaseFolder}/sample.xlsx"



###### release폴더 삭제/생성 및 소스 복사, exec.py 생성
if os.path.exists(f"{releaseFolder}"):
  shutil.rmtree(f"{releaseFolder}")  ##### 하위 폴터파일 전부 삭제

os.mkdir(f"{releaseFolder}")


  
  
###### dist 폴더에 db.xlsx, config.js2, inventory.xlsx, topology.j2, defaults.j2 복사
shutil.copy(sampleXlsx, toSampleXlsx)


print(cmd)
os.system(cmd)


###### exec 생성
# filepath = f"python release/createExec.py"
# subprocess.check_output(filepath, shell=True)

# os.remove(f"{src_dir}/createExec.py")

print("complete~!")