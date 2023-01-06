import pyautogui
import win32com.client
import subprocess
import time
import os

def __init__():
        application = win32com.client.GetObject("SAPGUI").GetScriptingEngine
        connection = application.Children
        session = connection

def saplogin():

        path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
        subprocess.Popen(path)
        time.sleep(2)
        pyautogui.hotkey('enter')
        time.sleep(1)
        time.sleep(1)
        pyautogui.hotkey('tab')
        time.sleep(1)
        pyautogui.typewrite("Colombia2023")
        time.sleep(1)
        pyautogui.hotkey('enter')
        time.sleep(2)

def reprocessar_monitor() :
        time.sleep(2)
        txt = open(r"C:\Users\abimaelsoares\AppData\Roaming\SAP\SAP GUI\Scripts\script.bat", 'r')
        processo = txt.read().strip()
        subprocess.Popen(processo)



os.system('wmic process where name="saplogon.EXE" delete')
time.sleep(1)
saplogin()
time.sleep(2)
reprocessar_monitor()


