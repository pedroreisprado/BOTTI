import os
import requests
import json

from config import WHATS_NOAH_KEY, WHATS_NOAH_TOKEN
from functions.logger import logger

def noahSendMsgFile(number, message, imagePath, imageName):
    url = "https://2rfpapi.bragimulticanal.com.br/v1/api/external/" + str(WHATS_NOAH_KEY)

    payload = {
        'number': str(number),
        'externalKey': WHATS_NOAH_KEY,
        'body': str(message)
    }

    image_path = os.path.join('/', imagePath)
    files = {'media': (imageName, open(image_path, 'rb'), 'image/jpeg')}

    headers = {
        'Authorization': 'Bearer ' + str(WHATS_NOAH_TOKEN)
    }

    response = requests.post(url, headers=headers, data=payload, files=files)
    logger(f"MENSAGEM ENVIADA: "+ str(response.text) + str(number))

    return response.text

def noahSendMsgText(number, message):
    url = "https://2rfpapi.bragimulticanal.com.br/v1/api/external/" + str(WHATS_NOAH_KEY)

    payload = json.dumps({
        "body": str(message),
        "number": str(number),
        "externalKey": WHATS_NOAH_KEY
     })
    
    headers = {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + str(WHATS_NOAH_TOKEN)
    }

    response = requests.request("POST", url, headers=headers, data=payload)
    logger(f"MENSAGEM ENVIADA: "+ str(response.text) + str(number))

    return response.text
