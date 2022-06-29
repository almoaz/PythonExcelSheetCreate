import pyautogui as task
import time




while True:
    screenWidth, screenHeight = task.size()
    task.sleep(7)
    X, Y = task.position()
    file = open("C:\\Users\Mahfuz Salehin Moaz\Desktop\\file1.txt", "a")
    print(X,",",Y)
    x = X
    y = Y

    a = "task.moveTo("
    b = str(x)+", "+str(y)+")"
    c = a+""+b

    file.write(c+"\n")

    a = ""
    b = ""
    c = ""




