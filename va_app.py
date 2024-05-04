from flask import Flask, render_template, send_from_directory
from PyQt5.QtWidgets import QApplication
from ai_assist import ChatWindow
import os

app = Flask(__name__)

# Get the current directory
current_directory = os.path.dirname(os.path.abspath(__file__))
# Define the template folder path
template_folder = os.path.join(current_directory, 'templates')
static_folder = os.path.join(current_directory, 'static')
dist_folder = os.path.join(current_directory, 'dist')
# Set the template folder for Flask
app.template_folder = template_folder
app.static_folder = static_folder

# Route for serving static files
@app.route('/static/<path:filename>')
def serve_static(filename):
    return send_from_directory(app.static_folder, filename)

# Route for serving .exe file from the dist folder
@app.route('/dist/<path:filename>')
def serve_exe(filename):
    return send_from_directory(dist_folder, filename)

@app.route('/')
def index():
    return render_template('index.html')  # This will serve the HTML template with your PyQt5 application

@app.route('/voice_assistant')
def voice_assistant():
    from ai_assist import ChatWindow  # Import your ChatWindow class
    app = QApplication([])
    window = ChatWindow()
    window.show()  # Show the ChatWindow
    app.exec_()  # Start the application event loop
    return "Voice assistant started"

if __name__ == '__main__':
    app.run(debug=True)  # Run the Flask application
