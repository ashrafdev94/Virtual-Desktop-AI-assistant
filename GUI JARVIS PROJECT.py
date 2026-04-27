from PyQt5 import QtWidgets, QtGui, QtCore
from PyQt5.QtGui import QMovie
import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.uic import loadUiType
import pyttsx3
import speech_recognition as sr
import os
import time
import webbrowser
import datetime
import pyautogui
import pywhatkit
import pyjokes
import wikipedia
from PIL import Image, ImageDraw, ImageFont
flags = QtCore.Qt.WindowFlags(QtCore.Qt.FramelessWindowHint)

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice',voices[0].id)
engine.setProperty('rate',180)



def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def wish():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour <12:
        speak("Good morning Sir. How may I Assist you Today")
    elif hour>=12 and hour<18:
        speak("Good Afternoon Sir. How may I Assist you Today")
    else:
        speak("Good night Sir. How may I Assist you Today")

        
class mainT(QThread):
    def __init__(self):
        super(mainT,self).__init__()
    
    def run(self):
        self.JARVIS()
    
    def STT(self):
        try:
            R = sr.Recognizer()
            with sr.Microphone() as source:
                print("Microphone detected. Adjusting for ambient noise...")
                R.adjust_for_ambient_noise(source, duration=1)
                R.energy_threshold = 100  # Lower threshold for better sensitivity
                print(f"Energy threshold set to: {R.energy_threshold}")
                print("Listening...........")
                audio = R.listen(source, timeout=30, phrase_time_limit=15)
            try:
                print("Recognizing......")
                text = R.recognize_google(audio, language='en-in')
                print(">> ", text)
            except sr.UnknownValueError:
                speak("Sorry, I didn't catch that. Please speak again.")
                return "None"
            except sr.RequestError as e:
                speak("Sorry, there was an error with the speech recognition service.")
                print(f"RequestError: {e}")
                return "None"
            text = text.lower()
            return text
        except sr.WaitTimeoutError:
            speak("Listening timed out. Please try again.")
            return "None"
        except Exception as e:
            speak("Microphone error or unexpected issue. Please check your microphone.")
            print(f"Microphone/Exception error: {e}")
            return "None"
    

    def JARVIS(self):
        wish()
        while True:
            self.query = self.STT()
            if 'good bye' in self.query:
                sys.exit()
            elif 'open google' in self.query:
                webbrowser.open('www.google.co.in')
                speak("opening google")
            elif 'open the firefox' in self.query:
                speak("Opening Firefox Application sir...")
                try:
                    os.startfile("C:\\Program Files\\Mozilla Firefox\\firefox.exe")
                except FileNotFoundError:
                    speak("Firefox is not installed or the path is incorrect.")
            elif 'joke' in self.query:
                jarvisJoke = pyjokes.get_joke()
                print(jarvisJoke)
                speak(jarvisJoke)
            elif 'play' in self.query:
               playQuery=self.query.replace('play','')
               speak("Playing " + playQuery)
               pywhatkit.playonyt(playQuery)
            elif "pause" in self.query:
                pyautogui.press("k")
                speak("video paused")
            elif "now start" in self.query:
                pyautogui.press("k")
                speak("video started")
            elif "mute" in self.query:
                pyautogui.press("m")
                speak("video muted")
            elif 'open youtube' in self.query:
                webbrowser.open("www.youtube.com")
                speak("Opening YouTube")
            elif 'introduce yourself' in self.query or 'who are you' in self.query:
             speak("Hello Sir! I am Jarvis, a Virtual Artificial Intelligence Assistant. "
          "I am designed to assist you with various tasks such as answering queries, "
          "opening applications, controlling media playback, telling jokes, and much more. "
          "How can I assist you today?")
            elif 'what is artificial intelligence' in self.query or 'define artificial intelligence' in self.query:
              speak("Artificial Intelligence, or AI, refers to the simulation of human intelligence in machines. "
          "It enables computers to perform tasks that typically require human intelligence, such as learning, "
          "problem-solving, decision-making, and understanding natural language. AI is used in various fields, "
          "including robotics, healthcare, finance, and more. Would you like to know anything else?")
            elif "jarvis i am tired" in self.query:
                speak("Playing your favourite song sir")
                webbrowser.open("https://www.youtube.com/watch?v=-r687V8yqKY")


            elif "click my photo" in self.query:
                    pyautogui.press("super")
                    pyautogui.typewrite("camera")
                    pyautogui.press("enter")
                    pyautogui.sleep(2)
                    speak("SMILE PLEASE")
                    pyautogui.press("enter")
            elif 'date' in self.query or 'time' in self.query:
             speak(datetime.datetime.now().strftime("Today's date is %A, %d %B %Y, and the time is %I:%M %p."))
            elif 'calculate' in self.query:
                try:
                    expression = self.query.replace('calculate', '').strip()
                    result = eval(expression)
                    speak(f"The result is {result}")
                except:
                    speak("Sorry, I couldn't calculate that.")
            elif 'wikipedia' in self.query:  # Wikipedia query
                try:
                    speak("Searching Wikipedia...")
                    query = self.query.replace('wikipedia', '').strip()
                    result = wikipedia.summary(query, sentences=2)
                    speak("According to Wikipedia")
                    speak(result)
                except Exception as e:
                    speak("Sorry, I couldn't find any results on Wikipedia.")
            elif 'microphone off' in self.query or 'stop listening' in self.query:
                speak("Microphone is now off. Call me when you need me.")
                break
            elif 'microphone on' in self.query or 'start listening' in self.query:
                speak("Microphone is now on. How can I assist you?")
                self.JARVIS()
            elif 'open notepad' in self.query:
                speak("Opening Notepad")
                os.startfile("C:\\Windows\\System32\\notepad.exe")
            elif 'open command prompt' in self.query:
                speak("Opening Command Prompt")
                os.startfile("C:\\Windows\\System32\\cmd.exe")
            elif 'open calculator' in self.query:   
                speak("Opening Calculator")
                os.startfile("C:\\Windows\\System32\\calc.exe")
            elif 'open camera' in self.query:
                speak("Opening Camera")
                os.startfile("C:\\Windows\\System32\\Camera.exe")
            elif 'open paint' in self.query:
                speak("Opening Paint")
                os.startfile("C:\\Windows\\System32\\mspaint.exe")
            elif 'open word' in self.query:
                speak("Opening Microsoft Word")
                os.startfile("C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE")
            elif 'open excel' in self.query:
                speak("Opening Microsoft Excel")
                os.startfile("C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE")
            elif 'open power point' in self.query:
                speak("Opening Microsoft PowerPoint")
                os.startfile("C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE")    
            elif 'write what i say' in self.query:
                speak("I am ready to write. Please start speaking.")
                while True:
                    text_to_write = self.STT()
                    if 'stop writing' in text_to_write:
                        speak("Stopped writing.")
                        break
                    else:
                        with open("output.txt", "a") as file:
                            file.write(text_to_write + "\n")
                        speak("Written successfully.")
            elif 'open file' in self.query:
                speak("Please tell me the file name")
                file_name = self.STT()
                try:
                    os.startfile(file_name)
                    speak(f"Opening {file_name}")
                except FileNotFoundError:
                    speak(f"Sorry, I couldn't find the file named {file_name}.")
            elif 'close notepad' in self.query:
                speak("Closing Notepad")
                os.system("taskkill /f /im notepad.exe")
            elif 'close command prompt' in self.query:
                speak("Closing Command Prompt")
                os.system("taskkill /f /im cmd.exe")    
            elif 'close calculator' in self.query:
                speak("Closing Calculator")
                os.system("taskkill /f /im calc.exe")
            elif 'close paint' in self.query:
                speak("Closing Paint")
                os.system("taskkill /f /im mspaint.exe")
            elif 'close word' in self.query:
                speak("Closing Microsoft Word")
                os.system("taskkill /f /im WINWORD.EXE")
            elif 'close excel' in self.query:
                speak("Closing Microsoft Excel")
                os.system("taskkill /f /im EXCEL.EXE")
            elif 'close power point' in self.query:
                speak("Closing Microsoft PowerPoint")
                os.system("taskkill /f /im POWERPNT.EXE")
            elif 'close file' in self.query:
                speak("Closing the file")
                os.system("taskkill /f /im " + self.query.replace('close file', '').strip())
            elif 'close all' in self.query:
                speak("Closing all applications")
                os.system("taskkill /f /im notepad.exe")
                os.system("taskkill /f /im cmd.exe")
                os.system("taskkill /f /im calc.exe")
                os.system("taskkill /f /im Camera.exe")
                os.system("taskkill /f /im mspaint.exe")
                os.system("taskkill /f /im WINWORD.EXE")
                os.system("taskkill /f /im EXCEL.EXE")
                os.system("taskkill /f /im POWERPNT.EXE")
                os.system("taskkill /f /im firefox.exe")
                os.system("taskkill /f /im chrome.exe")
            elif 'close firefox' in self.query:
                speak("Closing Firefox")
                os.system("taskkill /f /im firefox.exe")
            elif 'close all tabs' in self.query:
                speak("Closing all tabs in Firefox")
                pyautogui.hotkey('ctrl', 'w')
            elif 'search google for' in self.query:
                search_query = self.query.replace('search google for', '').strip()
                if search_query:
                    url = f"https://www.google.com/search?q={search_query.replace(' ', '+')}"
                    webbrowser.open(url)
                    speak(f"Searching Google for {search_query}")
                else:
                    speak("Please specify what to search for.")
                
            elif 'take a screenshot' in self.query:
                speak("Taking a screenshot")
                screenshot = pyautogui.screenshot()
                screenshot.save("screenshot.png")
                speak("Screenshot saved as screenshot.png")
            elif 'open file explorer' in self.query:
                speak("Opening File Explorer")
                os.startfile("explorer.exe")
            elif 'open settings' in self.query:
                speak("Opening Settings")
                os.startfile("ms-settings:")
            elif 'shut down the system' in self.query or 'shutdown the system' in self.query:
                speak("Shutting down the system. Goodbye!")
                os.system("shutdown /s /t 1")
            elif 'restart the system' in self.query:
                speak("Restarting the system. Goodbye!")
                os.system("shutdown /r /t 1")                




FROM_MAIN,_ = loadUiType(os.path.abspath(os.path.join(os.path.dirname(__file__),"./scifi.ui")))


class Main(QMainWindow,FROM_MAIN):
    def __init__(self,parent=None):
        super(Main,self).__init__(parent)
        self.setupUi(self)
        self.setFixedSize(1920,1080)
        self.exitB.setStyleSheet("background-image:url(" + os.path.abspath(os.path.join(os.path.dirname(__file__), "lib", "exit - Copy.png")).replace("\\", "/") + ");\n"
        "border:none;")
        self.exitB.clicked.connect(self.close)
        self.setWindowFlags(flags)
        Dspeak = mainT()
        self.label_7 = QMovie(os.path.abspath(os.path.join(os.path.dirname(__file__), "lib", "gifloader.gif")), QByteArray(), self)
        self.label_7.setCacheMode(QMovie.CacheAll)
        self.label_4.setMovie(self.label_7)
        self.label_7.start()

        self.ts = time.strftime("%A, %d %B")

        Dspeak.start()
        self.label.setPixmap(QPixmap(os.path.abspath(os.path.join(os.path.dirname(__file__), "lib", "tuse.png"))))
        self.label_5.setText("<font size=8 color='white'>"+self.ts+"</font>")
        self.label_5.setFont(QFont(QFont('Acens',8)))




app = QtWidgets.QApplication(sys.argv)
main = Main()
main.show()
exit(app.exec_())
