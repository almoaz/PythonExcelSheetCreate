import pyautogui as task
import time




while True:

    time.sleep(4)
    task.moveTo(473, 117)
    task.click()
    task.write('<!', interval=0.25)
    task.press('enter')
    time.sleep(1)
    task.press('enter')
    task.write('</html>', interval=0.25)
    task.press('enter')
    time.sleep(2)
    task.moveTo(1750, 118)
    task.click()
    task.press('enter')
    time.sleep(2)
    x, y = task.position()
    print(x,",",y)


    '''time.sleep(10)
    x, y = task.position()
    print(x, ",", y)
    time.sleep(4)'''

