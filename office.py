import win32com.client
from datetime import datetime
import os

def open_word_document():
    word_app = win32com.client.Dispatch("Word.Application")
    word_app.Visible = True
    doc = word_app.Documents.Add()
    return word_app, doc

def close_word_document(word_app, doc):
    doc.Close(False)
    word_app.Quit()

def save_document(doc):
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
    filename = os.path.join(downloads_folder, f"document_{timestamp}.docx")
    doc.SaveAs(filename)
    return filename



