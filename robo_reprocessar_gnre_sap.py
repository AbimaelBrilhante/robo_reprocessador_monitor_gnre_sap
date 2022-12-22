import pyautogui
import win32com.client
import sys
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
        pyautogui.typewrite("*********")
        time.sleep(1)
        pyautogui.hotkey('tab')
        time.sleep(1)
        pyautogui.typewrite("*********")
        time.sleep(1)
        pyautogui.hotkey('enter')
        time.sleep(2)

def reprocessar_monitor() :
        time.sleep(1)
        pyautogui.typewrite("zgnre_monitor")
        time.sleep(1)
        pyautogui.hotkey('enter')
        time.sleep(1)
        pyautogui.hotkey('shift', 'f5')
        time.sleep(1)
        pyautogui.hotkey('down')
        time.sleep(1)
        pyautogui.hotkey('down')
        time.sleep(1)
        pyautogui.hotkey('down')
        time.sleep(1)
        pyautogui.hotkey('down')
        time.sleep(1)
        pyautogui.hotkey('down')
        time.sleep(1)
        pyautogui.hotkey('f2')
        time.sleep(1)
        pyautogui.hotkey('f8')
        time.sleep(15)
        pyautogui.click(1065, 543)
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'a')
        #pyautogui.click(16, 171)
        time.sleep(1)
        pyautogui.hotkey('f8')
        time.sleep(1)
        pyautogui.hotkey('enter')


try:
        os.system('wmic process where name="saplogon.EXE" delete')
        time.sleep(1)
        saplogin()
        time.sleep(2)
        reprocessar_monitor()
except:
        pass

