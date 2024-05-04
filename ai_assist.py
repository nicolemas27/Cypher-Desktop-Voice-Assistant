import sys
import os
import google.generativeai as genai
import speech_recognition as sr
import pyttsx3
import webbrowser, docx
import subprocess
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, \
    QHBoxLayout, QPushButton, QLineEdit, QTextEdit, QLabel, QRadioButton
from PyQt5.QtGui import QFont, QIcon
from office import open_word_document, close_word_document, save_document
from datetime import datetime

from app_opener import open_application, take_screenshot
from dotenv import load_dotenv
import webbrowser
from PyQt5.QtGui import QPixmap, QMovie
from PyQt5.QtCore import Qt, QSize


DOWNLOADS_FOLDER = os.path.join(os.path.expanduser("~"), "Downloads")

class ChatWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.init_ui()
        self.word_app = None
        self.doc = None
        self.listen_for_text = False
        self.continued_input = False
        self.last_response = ""
        self.voice_gender = "male"
        self.manual_path = "commands_manual.pdf"

    def init_ui(self):
        load_dotenv()
        google_api_key = os.getenv("GOOGLE_API_KEY")
        genai.configure(api_key=os.environ['GOOGLE_API_KEY'])
        self.model = genai.GenerativeModel('gemini-pro')

        self.setWindowTitle("Cypher Assistant")  # Set window title
        self.setGeometry(100, 100, 600, 400)

        self.setWindowIcon(QIcon('va.ico'))

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout()
        central_widget.setLayout(layout)

        text_label = QLabel("How can I help?")
        text_label.setStyleSheet("color: #66FCF1; font-size: 24px;")
        text_label.setAlignment(Qt.AlignCenter)  # Align text to the center
        layout.addWidget(text_label)

        
        # Input box
        input_layout = QHBoxLayout()
        layout.addLayout(input_layout)

        input_label = QLabel("Input:")
        input_label.setStyleSheet("color: white;")
        input_layout.addWidget(input_label)

        self.user_input = QLineEdit()
        self.user_input.setStyleSheet("background-color: #022140; color: white;")  #input box
        input_layout.addWidget(self.user_input)

        
        self.chat_history_display = QTextEdit()
        self.chat_history_display.setStyleSheet("background-color: #022140; color: white;") 
        self.chat_history_display.setReadOnly(True)  
        layout.addWidget(self.chat_history_display)

        
        self.chat_history_display.append("Hello, I'm Cypher. How can I assist you today?")

        # GIF image
        movie = QMovie("mic.gif")  # Replace "wave.gif" with your GIF path
        movie.setScaledSize(QSize(600, 200))  # Adjust the size of the GIF
        movie.start()
        gif_label = QLabel()
        gif_label.setMovie(movie)
        layout.addWidget(gif_label)

        # Send button
        send_button = QPushButton("Send")
        send_button.setStyleSheet("background-color: #4681f4; color: white;")  
        layout.addWidget(send_button)
        send_button.clicked.connect(self.send_message)

        # Voice input button
        voice_button = QPushButton("Voice Input")
        voice_button.setStyleSheet("background-color: #5783db; color: white !important;")  
        layout.addWidget(voice_button)
        voice_button.clicked.connect(self.listen_voice_input)

        read_aloud_button = QPushButton("Read Aloud")
        read_aloud_button.setStyleSheet("background-color: #55c2da; color: white;")   
        layout.addWidget(read_aloud_button)
        read_aloud_button.clicked.connect(self.read_document_aloud)

        self.listening_label = QLabel()
        self.listening_label.setStyleSheet("color: #9AC8CD;")
        layout.addWidget(self.listening_label)

        gender_layout = QHBoxLayout()
        male_radio = QRadioButton("Male")
        male_radio.setChecked(True)  # Default to male
        male_radio.setStyleSheet("color: #A3D8FF;")
        male_radio.toggled.connect(lambda: self.set_voice_gender("male"))
        female_radio = QRadioButton("Female")
        female_radio.setStyleSheet("color: #A3D8FF;")
        female_radio.toggled.connect(lambda: self.set_voice_gender("female"))
        gender_layout.addWidget(male_radio)
        gender_layout.addWidget(female_radio)
        layout.addLayout(gender_layout)

        self.setStyleSheet("background-color: #1F2833;")  # Set main window background color to white
        

        self.show()

        
        self.tts_engine = pyttsx3.init()

    def send_message(self):
        query = self.user_input.text().strip().lower()  # Convert the query to lowercase and strip whitespaces
        self.user_input.clear()
        self.update_chat_history(f"User: {query}\n")

        if query == "quit":
            self.close()
            return
        
        if query == "user manual":
            print("Opening user manual")
            self.display_help()
            return
        
        if query == "open word document":
            self.open_word_document()
            return
        
        if query == "start typing":
            self.listen_for_text = True
            self.read_out_loud("Please start speaking the text you want to add to the document.")
            return
        
        if query == "stop typing":
            self.listen_for_text = False  # New line to stop listening for text input
            if self.doc is not None:
                self.save_document()  # Save the document when user stops typing
                self.read_out_loud("Document saved.")
            return
        
        if query == "close word document":
            if self.doc is not None:
                self.close_word_document()
            else:
                self.update_chat_history("No document opened. Please open a Word document first.\n")
            return
        
        if query == "save document":
            if self.doc is not None:
                self.save_document()
            else:
                self.update_chat_history("No document opened. Please open a Word document first.\n")
            return
        
        if query == "read aloud":
            if self.doc is not None:
                self.read_document_aloud()
            else:
                self.update_chat_history("No document opened. Please open a Word document first.\n")
            return
        
        if self.listen_for_text:
            if query:  # Only proceed if there's actual text
                if self.doc is not None:
                    if "\n" in query:
                        query = query.replace("\n", "")
                        self.doc.Content.InsertAfter("\n")
                    self.update_chat_history(f"User said: {query}\n")
                    self.doc.Content.InsertAfter(query + "\n")
                    self.read_out_loud("Text added. Keep typing or say 'stop typing' to finish.")
                else:
                    self.update_chat_history("No document opened. Please open a Word document first.\n")
            return

    

        if query == "take screenshot":
            response = take_screenshot()
            self.update_chat_history(response)
            if "captured" in response:  
                pass
            return
            

        if query.startswith("search"):
            search_query = query[len("search"):].strip()
            self.update_chat_history(f"Searching for: {search_query}")
            webbrowser.open(f"https://www.google.com/search?q={search_query}")
            return

        if query.startswith("open"):
            app_name = query[len("open"):].strip().lower()
            if app_name == "camera":
                try:
                    subprocess.Popen(["start", "cmd", "/k", "start", "microsoft.windows.camera:"], shell=True)
                    self.camera_opened = True
                    self.update_chat_history(f"Opening {app_name}")
                except Exception as e:
                    self.update_chat_history(f"Error: {e}")
            elif app_name == "google":
                self.update_chat_history("Opening Google")
                webbrowser.open("https://www.google.com")
            elif app_name == "youtube":
                self.update_chat_history("Opening YouTube")
                webbrowser.open("https://www.youtube.com")
            elif app_name == "notepad":
                subprocess.Popen(["notepad.exe"])
                self.update_chat_history(f"Opening {app_name}")
            elif app_name == "calculator":
                subprocess.Popen(["calc.exe"])
                self.update_chat_history(f"Opening {app_name}")
            elif app_name == "whatsapp":
                response = open_application("whatsapp")
                self.update_chat_history(response)
                return
            elif app_name == "spotify":
                response = open_application("spotify")
                self.update_chat_history(response)
                return
            elif app_name == "excel":
                os.startfile("excel.exe")
                self.update_chat_history(f"Opening {app_name}")
            elif app_name == "chrome":
                os.startfile("chrome.exe")
                webbrowser.get("chrome").open(f"https://www.google.com/search?q={search_query}")
                self.update_chat_history(f"Searching in Chrome for: {search_query}")

            else:
                self.update_chat_history("Unsupported application")
            return
        
        if query.startswith("close"):
            close_app_name = query[len("close"):].strip().lower()
            if close_app_name == "camera":
                subprocess.Popen(["taskkill", "/F", "/IM", "WindowsCamera.exe"], shell=True)
                self.update_chat_history(f"closing {close_app_name}")
                return
            
            elif close_app_name == "google" or close_app_name == "youtube" or close_app_name =="whatsapp":
                subprocess.Popen(["taskkill", "/F", "/IM", "chrome.exe"], shell=True)
                self.update_chat_history("closing Google")
                return
            
            elif close_app_name == "notepad":
                subprocess.Popen(["taskkill", "/F", "/IM", "notepad.exe"], shell=True)
                self.update_chat_history("closing notepad")
                return
            
            elif close_app_name == "calculator":
                subprocess.Popen(["taskkill", "/F", "/IM", "CalculatorApp.exe"], shell=True)
                self.update_chat_history("closing calculator")
                return
            
            elif close_app_name == "excel":
                subprocess.Popen(["taskkill", "/F", "/IM", "excel.exe"], shell=True)
                self.update_chat_history(f"closing {close_app_name}")
                return
            
            elif close_app_name == "spotify":
                subprocess.Popen(["taskkill", "/F", "/IM", "spotify"], shell=True)
                self.update_chat_history(f"closing {close_app_name}")
                return
        

        if query.startswith("youtube search"):
            search_query = query[len("youtube search"):].strip()
            self.update_chat_history(f"Searching YouTube for: {search_query}")
            webbrowser.open(f"https://www.youtube.com/results?search_query={search_query}")
            return
        
        if "time" in query:
            current_time = datetime.now().strftime("%I:%M %p")
            self.update_chat_history(f"The current time is {current_time}.\n")
            self.read_out_loud(f"The current time is {current_time}.")
            return
        
        if "date" in query:
            current_date = datetime.now().strftime("%A, %B %d, %Y")
            self.update_chat_history(f"Today's date is {current_date}.\n")
            self.read_out_loud(f"Today's date is {current_date}.")
            return

        if query.startswith("what day is it"):
            current_day = datetime.now().strftime("%A")
            self.update_chat_history(f"Today is {current_day}.\n")
            self.read_out_loud(f"Today is {current_day}.")
            return

        response = self.model.generate_content(query)
        response_text = response.text.replace('*', '')  
        self.update_chat_history(f"Assistant: {response_text}\n")
        self.read_out_loud(response_text)
        self.last_response = response_text

  

    def listen_voice_input(self):
        self.read_out_loud("Listening...")
        self.listening_label.setText("Listening...")
        recognizer = sr.Recognizer()
        with sr.Microphone() as source:
            print("Listening...")
            recognizer.adjust_for_ambient_noise(source)
            audio = recognizer.listen(source)

        try:
            query = recognizer.recognize_google(audio).lower()
            self.user_input.setText(query)

            # Check if the user asked to repeat
            if "repeat" in query or "what did you say" in query:
                if self.last_response:
                    self.update_chat_history(f"User: {query}\n")
                    self.update_chat_history(f"Assistant: {self.last_response}\n")
                    self.read_out_loud(self.last_response)
                    return
                else:
                    self.update_chat_history("Assistant: There is nothing to repeat.\n")
                    self.read_out_loud("There is nothing to repeat.")
                    return
            # Check for other commands
            elif query:
                self.send_message()


        
        except sr.UnknownValueError:
            print("Could not understand audio")
            self.update_chat_history("Could not understand audio. Please try again.\n")
            self.read_out_loud("Could not understand audio. Please try again.")    
            self.listening_label.setText("Could not understand audio")
        except sr.RequestError as e:
            print(f"Could not request results; {e}")
    
    

    def listen_for_repeat(self):
        recognizer = sr.Recognizer()
        with sr.Microphone() as source:
            recognizer.adjust_for_ambient_noise(source)
            audio = recognizer.listen(source)

        

    def update_chat_history(self, text):
        self.chat_history_display.append(text)  # Append the text to the response box
        self.chat_history_display.verticalScrollBar().setValue(self.chat_history_display.verticalScrollBar().maximum())

    def read_out_loud(self, text):
        if self.voice_gender == "male":
            self.tts_engine.setProperty('voice', 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\TTS_MS_EN-US_DAVID_11.0')
        elif self.voice_gender == "female":
            self.tts_engine.setProperty('voice', 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\TTS_MS_EN-US_ZIRA_11.0')
        self.tts_engine.say(text)
        self.tts_engine.runAndWait()
    
    def set_voice_gender(self, gender):
        if gender == "male":
            self.voice_gender = "male"
        elif gender == "female":
            self.voice_gender = "female"    

    def open_word_document(self):
        self.word_app, self.doc = open_word_document()
        self.update_chat_history("Word document opened.\n")
        self.read_out_loud("Word document opened.")

    def add_text_to_document(self):
        while self.listen_for_text:
            text = self.listen_voice_input()
            if text:
                if "stop typing" in text:
                    self.listen_for_text = False  # New line to stop listening for text input
                    self.update_chat_history("Document saved.\n")
                    self.save_document()
                    break
                
                if "\n" in text:
                    text = text.replace("\n", "")
                    self.doc.Content.InsertAfter("\n")
                
                self.update_chat_history(f"User said: {text}\n")
                self.doc.Content.InsertAfter(text + "\n")
                
                self.read_out_loud("Text added. Keep typing or say 'stop typing' to finish.")
            else:
                self.read_out_loud("No text detected. Please try again.")
    
    def save_document(self):
        filename = save_document(self.doc)
        self.update_chat_history(f"Document saved as '{filename}'.\n")
    
    def close_word_document(self):
        close_word_document(self.word_app, self.doc)
        self.update_chat_history("Word document closed.\n")
        self.read_out_loud("Word document closed.")

    def read_document_aloud(self):
        if self.doc:
            content = self.doc.Content.Text
            if content:
                self.read_out_loud(content)
                self.update_chat_history("Document read aloud.\n")
            else:
                self.update_chat_history("No content to read.\n")
        else:
            self.update_chat_history("No document opened. Please open a Word document first.\n")
    
    def display_help(self):
        if os.path.exists(self.manual_path):
            try:
                webbrowser.open(self.manual_path)
            except Exception as e:
                self.update_chat_history(f"Error opening user manual: {e}\n")
        else:
            self.update_chat_history("User manual not found.\n")
    
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ChatWindow()
    sys.exit(app.exec_())
