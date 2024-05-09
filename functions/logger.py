import os
from datetime import datetime
from config import LOG_PATH

def createFolder():
    myDate = datetime.now()
    month = myDate.month
    year = myDate.year
    
    myDir = os.path.join(LOG_PATH, str(year))

    if not os.path.exists(myDir):
        os.makedirs(myDir)

    myDir = os.path.join(myDir, str(month))
    
    if not os.path.exists(myDir):
        os.makedirs(myDir)

    return myDir


def logger(message, message_label=None):
    thisFolder = createFolder()
    myDate = datetime.now()
    today = myDate.day

    fullPath = os.path.join(thisFolder, f"{today}.txt")
    
    mydatetime = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
    message = f"[{mydatetime}] - {message}"
    
    with open(fullPath, "a") as log_file:
        log_file.write(f"{message}\n")
    
    print(message)
        

