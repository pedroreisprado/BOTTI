import subprocess
import time
import win32gui
import win32con
import pygetwindow as gw
import pyautogui

from config import LINX_USER, LINX_PWD, LINX_PATH
from functions.logger import logger
from functions.helpers import screenShotForText, pathImg

def open(PATH) -> bool:
    processo = subprocess.Popen(PATH)
    time.sleep(10)

    try:
        hwnd = win32gui.GetForegroundWindow()
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
        return True
        
    except Exception as e:
        print(f"Erro ao tentar trazer a janela para o primeiro plano: {e}")
        return False

def auth():
    time.sleep(10)
    pyautogui.write(LINX_USER)
    time.sleep(1)
    pyautogui.press('tab')
    time.sleep(1)
    pyautogui.write(LINX_PWD)
    time.sleep(3)
    pyautogui.press('enter')
    time.sleep(10)
    pyautogui.press('esc')
    time.sleep(2)

    #img = pathImg("../assets" , "favorite.png")

    favorite = pyautogui.locateOnScreen("C:/2RFP_1/assets/favorite.png",confidence=0.3)

    if(favorite):
        x, y, width, height = favorite
        x_centro = x + width / 2
        y_centro = y + height / 2
        logger(f"LINX Login efetuado com sucesso")
        return x_centro, y_centro
    
    return False

def findStock(sheetClients):
    productsInStock = []
    oldProd = 0
    try:
        for item in sheetClients:
               
            #--- ENCONTRAR 34.5 TELA
            if item['codNewProd'] != oldProd:
                if oldProd != 0:
                    for z in range(2):
                        pyautogui.click(25,133)

                oldProd = item['codNewProd']
                print(item['codNewProd'])
                time.sleep(5)
                pyautogui.write(item['codNewProd'])
                pyautogui.press('enter')
                time.sleep(10) 

                #--- AJUSTAR TELA
                print('AdjustScreen')
                adjustScreen(180)

                # img = pathImg("assets/ "345.png")
                print('Pegando Coordinates')
                coordinates = pyautogui.locateOnScreen("C:/2RFP_1/assets/345.png",confidence=0.2)
                print(coordinates)
                time.sleep(2)

                x, y, width, height = coordinates
                x_centro = x + width / 2
                y_centro = y + height / 2

                xi = x_centro - 100
                xf = xi + 60
                yi = y_centro + 37
                yf = yi + 13

                calcado = 34.0
                grids = {calcado: [xi, yi, xf, yf]}    

                for i in range(14):
                    calcado += 0.5
                    if i == 3 or i == 11:
                        xi += 70
                        xf = xi+72
                    else:
                        xi += 60
                        xf = xi+66

                    grids[calcado] = [xi, yi, xf, yf]

            for grid in grids:

                if float(item['grid']) == grid:
                    print(grid)
                    x1, y1, x2, y2 = grids[grid]

                    stock = str(screenShotForText(x1, y1, x2, y2))
                    print(f'Print : {x1} {y1} {x2} {y2}')
                    print(f'STOCK: {stock}')
                    print("--- ENCONTRADO: ", stock.strip())

                    #-- TRATAR POSSIVEIS LETRAS
                    if(stock.strip() == "Z"):
                        stock = "2"

                    #if stock.strip() != "0" and stock.strip() != "Oo" and stock.strip() != 0:
                    if stock.strip() != "teste":
                        for i in range(20):
                            if stock.strip() == str(i):
                                row_data = {
                                    "id": item['id'],
                                    "codClient": item['codClient'],
                                    "nameClient": item['nameClient'],
                                    "codNewProd": item['codNewProd'],
                                    "newProduct": item['newProduct'],
                                    "grid": item['grid'],    
                                    "seller": item['seller'],
                                    "stock": stock.strip()
                                }
                                productsInStock.append(row_data)
                                break            
                                     

    except Exception as e:
        logger(f"!! O script encontrou um erro: {e}")
    
    logger(f"Estoque de produtos encontrados: {str(productsInStock)}")
    time.sleep(100000)  
    return productsInStock
        
def adjustScreen(number):
    width, height = pyautogui.size()
    center_x = width // 2
    center_y = height // 2
    pyautogui.moveTo(center_x, center_y)
    pyautogui.mouseDown()
    pyautogui.move(number, 0, duration=1)
    pyautogui.mouseUp() 

    return True       

def close() -> bool:
    time.sleep(1)
    pyautogui.click(25, 45)
    time.sleep(2)
    pyautogui.press('enter')
    time.sleep(5)

    return True
