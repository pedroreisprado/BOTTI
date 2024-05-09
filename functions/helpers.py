from PIL import ImageGrab, ImageEnhance, ImageFilter
import pytesseract
import os
from datetime import datetime
from functions.logger import createFolder, logger
import random
import string
import time
import sys


def screenShotForText(x1, y1, x2, y2, save = False):
    captura = ImageGrab.grab(bbox=(x1, y1, x2, y2))

    if save:
        thisFolder = createFolder()
        named = generateHash()
        named = os.path.join(thisFolder, f"{named}.png")
        captura.save(named)
 
    # Pré-processamento adicional
    captura = captura.resize((captura.width * 2, captura.height * 2))  # Aumentar o tamanho
    captura = captura.convert("L")  # Converter para escala de cinza
    captura = captura.point(lambda x: 0 if x < 128 else 255)  # Binarização
    
    # Configurações do Tesseract
    captura = captura.filter(ImageFilter.SHARPEN)
    captura = ImageEnhance.Contrast(captura).enhance(2.0)

    config = '--psm 6'  
    text = pytesseract.image_to_string(captura, config=config)
    time.sleep(2)

    if not text:
        text = pytesseract.image_to_string(captura, config=config)

    logger(f"Screenshot registrado com sucesso, string encontrada: {str(text)}")

    return text

def generateHash(length=8):
    characters = string.ascii_letters + string.digits
    random_hash = ''.join(random.choice(characters) for _ in range(length))
    return random_hash

def pathImg(folder, image_name):
    script_dir = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    image_path = os.path.join(script_dir, folder, image_name)

    if not os.path.exists(image_path) and getattr(sys, 'frozen', False):
        temp_dir = sys._MEIPASS
        image_path = os.path.join(temp_dir, folder, image_name)
    
    return str(image_path)