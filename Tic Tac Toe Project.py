import pyautogui as task
import time

while True:
    time.sleep(3)

    # For Project  Folder Create
    task.moveTo(412, 441)
    task.click()
    time.sleep(3)

    task.moveTo(533, 52)
    task.click()
    time.sleep(3)

    task.moveTo(139, 83)
    task.click()
    time.sleep(1)
    task.write('Tic Tac Toe', interval=0.25)
    time.sleep(3)

    task.moveTo(431, 313)
    task.click()
    time.sleep(3)

    task.moveTo(269, 169)
    task.click()
    time.sleep(3)

    task.moveTo(1016, 681)
    task.click()
    time.sleep(3)

    # For Folder Open

    task.moveTo(757, 624)
    task.click()
    time.sleep(3)

    task.moveTo(849, 678)
    task.click()
    time.sleep(3)

    task.moveTo(1041, 735)
    task.click()
    time.sleep(1)

    # For HTML File Create
    task.moveTo(270, 90)
    time.sleep(1)
    task.click()
    time.sleep(3)


    task.moveTo(111, 114)
    time.sleep(1)
    task.write('index.html', interval=0.25)
    time.sleep(3)

    task.moveTo(414, 274)
    task.click()
    time.sleep(3)

    # For CSS File Create
    task.moveTo(270, 90)
    time.sleep(1)
    task.click()
    time.sleep(3)

    task.moveTo(111, 114)
    time.sleep(1)
    task.write('style.css', interval=0.25)
    time.sleep(3)

    task.moveTo(414, 274)
    task.click()
    time.sleep(3)

    # For JavaScript File Create
    task.moveTo(270, 90)
    time.sleep(1)
    task.click()
    time.sleep(3)

    task.moveTo(111, 114)
    time.sleep(1)
    task.write('app.js', interval=0.25)
    time.sleep(3)

    task.moveTo(414, 274)
    task.click()
    time.sleep(3)



    task.moveTo(0, 0)


    """ # For Live Server
    time.sleep(3)
    task.moveTo(188, 135)
    task.rightClick()
    time.sleep(3)

    task.moveTo(293, 150)
    time.sleep(1)
    task.click()
    time.sleep(3)
    """