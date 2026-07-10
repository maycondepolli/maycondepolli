import pyautogui
import time
link="https://globo.com" 
link2="https://r7.com"
link3="https://web.whatsapp.com/"
pyautogui.press("win")
pyautogui.write("chrome")
pyautogui.press("enter")
time.sleep(3)
pyautogui.click(x=175, y=64)
pyautogui.write(link)
pyautogui.press("enter")
time.sleep(3)
pyautogui.click(x=289, y=15)
pyautogui.write(link2)
pyautogui.press("enter")
time.sleep(3)
pyautogui.click(x=523, y=21)
pyautogui.write(link3)
pyautogui.press("enter")