import pyautogui as task
import time

while True:
    screenWidth, screenHeight = task.size()
    currentMouseX, currentMouseY = task.position()
    time.sleep(2)
    task.moveTo(100, 1060)
    task.click()
    time.sleep(2)
    task.write("Andro")
    time.sleep(2)
    task.press('enter')
    time.sleep(20)
    #task.moveTo(1110, 664)
    task.moveTo(1529, 43)
    time.sleep(2)
    task.click()
    time.sleep(1)
    task.moveTo(0, 0)