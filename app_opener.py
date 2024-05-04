# app_opener.py

import subprocess
import webbrowser
import pyautogui, time
import os
from datetime import datetime


DOWNLOADS_FOLDER = os.path.join(os.path.expanduser("~"), "Downloads")

def open_application(app_name):
    app_name = app_name.lower()
    if app_name == "whatsapp":
        webbrowser.open("https://web.whatsapp.com/")
        return "Opening WhatsApp"

    elif app_name == "spotify":
        subprocess.Popen(["spotify"])
        return "Opening Spotify"  
    elif app_name == "chrome":
        webbrowser.open("https://www.google.com")
        return "Opening Chrome"
    elif app_name == "excel":
        os.startfile("excel.exe")
        return "Opening Excel"
    
    else:
        return "Unsupported application"
    
    
def take_screenshot():
    try:
        screenshot = pyautogui.screenshot()
        timestamp = time.strftime("%Y%m%d-%H%M%S")  # Generate a unique timestamp
        filename = os.path.join(DOWNLOADS_FOLDER, f"screenshot_{timestamp}.png")
        screenshot.save(filename)
        return "Screenshot captured and saved to Downloads folder."
    except Exception as e:
        print(f"Error taking screenshot: {e}")
        return "Failed to take screenshot."

