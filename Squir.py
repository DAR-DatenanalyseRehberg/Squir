import pyautogui
import subprocess
import time
from datetime import datetime
import pandas as pd
SquirJob= pd.read_excel(r'Squir.xls')
Rownr=0
SquirPathApp=SquirJob.iat[Rownr,0]
SquirTable= SquirJob.iat[Rownr,2]
SquirVariant=SquirJob.iat[Rownr,3]
SquirFile=SquirJob.iat[Rownr,4]
SquirPath=SquirJob.iat[Rownr,5]
print("SQUIR by DAR-Analytics.com")
sap_gui = subprocess.Popen(SquirPathApp) #"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
time.sleep(1)
start = datetime.now()
pos = None
time.sleep(1)
while (datetime.now() - start).total_seconds() < 20 and not pos:
    pos = pyautogui.locateCenterOnScreen('SAPProd.png',confidence=0.8)
    print(pos)
if not pos: raise Exception('SAP login image was not found within 20 seconds!')
time.sleep(1)
pyautogui.doubleClick(*pos)
time.sleep(1)
start = datetime.now()
button = None
time.sleep(1)
while (datetime.now() - start).total_seconds() < 20 and not button:
    button = pyautogui.locateCenterOnScreen('SE16N.png',confidence=0.9)
    print(button)
if not button: raise Exception('SE16N screenshot was not found within 20 seconds!')
pyautogui.doubleClick(*button)
time.sleep(1)
print("SQUIR started..")
while Rownr < 4:
    SquirTable = SquirJob.iat[Rownr, 2]
    SquirVariant = SquirJob.iat[Rownr, 3]
    SquirFile = SquirJob.iat[Rownr, 4]
    SquirPath = SquirJob.iat[Rownr, 5]
    time.sleep(3)
    pyautogui.typewrite(SquirTable, interval=0.001)
    pyautogui.press('enter')
    pyautogui.keyDown('fn') # holding FN down so F8 can run on my marc
    pyautogui.press('f8') # works only if you also press FN
    time.sleep(1)
    pyautogui.press('f6') # works only if you also press FN, getting variant
    time.sleep(1)
    pyautogui.typewrite(SquirVariant, interval=0.001)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.keyDown('fn')
    pyautogui.press('f8')
    time.sleep(1)
    start = datetime.now()
    time.sleep(1)
    elande = None
    time.sleep(1)
    while (datetime.now() - start).total_seconds() < 300 and not elande:
        elande = pyautogui.locateCenterOnScreen('export.png',confidence=0.8)
    if not elande: raise Exception('export image was not found within 1000 seconds!')
    time.sleep(1)
    pyautogui.click(*elande)
    time.sleep(1)
    pyautogui.move(0, 140)
    time.sleep(1)
    pyautogui.click()
    start = datetime.now()
    shiva = None
    time.sleep(1)
    while (datetime.now() - start).total_seconds() < 300 and not shiva:
        shiva = pyautogui.locateCenterOnScreen('SaveAs.png',confidence=0.8)
    if not shiva: raise Exception('SaveAs image was not found within 1000 seconds!')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.click()
    time.sleep(1)
    start = datetime.now()
    salut = None
    time.sleep(1)
    while (datetime.now() - start).total_seconds() < 300 and not salut:
        salut = pyautogui.locateCenterOnScreen('TxtNameEntry.png',confidence=0.8)
    if not salut: raise Exception('TxtNameEntry image was not found within 1000 seconds!')
    time.sleep(1)
    pyautogui.typewrite(SquirFile, interval=0.001)
    time.sleep(1)
    pyautogui.keyDown('tab')
    time.sleep(1)
    pyautogui.keyDown('tab')
    time.sleep(1)
    pyautogui.keyDown('tab')
    time.sleep(1)
    pyautogui.keyDown('tab')
    time.sleep(1)
    pyautogui.keyDown('tab')
    time.sleep(1)
    pyautogui.keyDown('tab')
    time.sleep(1)
    pyautogui.typewrite(SquirPath, interval=0.001)
    pyautogui.press('enter')
    time.sleep(1)
    start = datetime.now()
    kos = None
    time.sleep(1)
    while (datetime.now() - start).total_seconds() < 300 and not kos:
        kos = pyautogui.locateCenterOnScreen('BackResultAfterSaving.png',confidence=0.8)
        print(kos)
    if not kos: raise Exception('BackResultAfterSaving image was not found within 1000 seconds!')
    time.sleep(1)
    pyautogui.press('f3')
    time.sleep(1)
    pyautogui.press('f3')
    time.sleep(1)
    Rownr=Rownr+1
    time.sleep(1)
sap_gui.kill()