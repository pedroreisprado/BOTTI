import pyautogui

while True:
    try:
        x, y = pyautogui.position()
        print(f'X: {x}, Y: {y}', end='\r')
        
    except KeyboardInterrupt:
        break