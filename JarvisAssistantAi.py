import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import smtplib
import random
import PyPDF2
import requests
from dotenv import load_dotenv
import time
import sys
import pywhatkit
import re
import cv2
import pyautogui
import psutil
from plyer import notification
import tkinter as tk
from tkinter import messagebox
import pyperclip
import pygetwindow as gw
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
import glob
import shutil
import schedule
import threading
import logging
import pyttsx3

# === Setup Logging ===
logging.basicConfig(level=logging.INFO, filename="jarvis.log", format="%(asctime)s - %(levelname)s - %(message)s")

# === Load Secrets ===
load_dotenv()
EMAIL = os.getenv("EMAIL_ADDRESS")
PASSWORD = os.getenv("EMAIL_PASSWORD")

# === Context Tracking ===
context = {"last_file": None, "last_app": None, "last_query": None}

# === Synonym Mapping for NLP ===
synonyms = {
    "open": ["start", "launch", "run"],
    "close": ["quit", "exit", "kill", "stop"],
    "find": ["search", "look for"],
    "move": ["transfer", "shift"],
    "type": ["write", "input"],
    "maximize": ["enlarge", "fullscreen"],
    "minimize": ["shrink", "hide"],
    "copy": ["copy", "duplicate"]
}

# === App Mapping ===
app_map = {
    "browser": "chrome",
    "code": "code",  # Fixed for Visual Studio Code
    "notepad": "notepad",
    "music": "spotify",
    "editor": "notepad"
}

# === Response Variants for Personality ===
response_variants = {
    "open": ["Launching {app}, sir!", "{app} is ready!", "Here comes {app}!"],
    "close": ["Shutting down {app}.", "{app} is closed, sir.", "Goodbye, {app}!"],
    "find": ["Found {result}!", "Here’s {result}, sir.", "Located {result}."],
    "error": ["Oops, something went wrong.", "My apologies, sir, that didn’t work.", "Hmm, let’s try that again."]
}

# === Speak Function ===
def speak(audio):
    print(f"JARVIS: {audio}")
    try:
        engine = pyttsx3.init('sapi5')
        voices = engine.getProperty('voices')
        for voice in voices:
            if "male" in voice.name.lower():
                engine.setProperty('voice', voice.id)
                break
        else:
            engine.setProperty('voice', voices[0].id)
        engine.say(audio)
        engine.runAndWait()
    except Exception as e:
        logging.error(f"pyttsx3 error: {e}")
        print(f"Speech error: {e}")

# === Notification ===
def notify(title, message):
    try:
        notification.notify(title=title, message=message, timeout=5)
    except Exception as e:
        logging.error(f"Notification error: {e}")

# === Voice Command with Noise Handling ===
def takeCommand(timeout=5):
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.energy_threshold = 1000
        r.adjust_for_ambient_noise(source, duration=1)
        r.pause_threshold = 0.8
        try:
            audio = r.listen(source, timeout=timeout, phrase_time_limit=5)
            query = r.recognize_google(audio, language='en-in').lower()
            print(f"User: {query}")
            context["last_query"] = query
            return query
        except sr.UnknownValueError:
            return "None"
        except sr.RequestError:
            speak("Check your internet connection.")
            return "None"
        except sr.WaitTimeoutError:
            return "None"
        except Exception as e:
            logging.error(f"Microphone error: {e}")
            return "None"

# === Wake Word Listener ===
def waitForWakeWord():
    print("Waiting for wake word: Say 'Jarvis'")
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.adjust_for_ambient_noise(source, duration=0.5)
        while True:
            print("Listening for wake word...")
            try:
                audio = r.listen(source, timeout=3, phrase_time_limit=3)
                query = r.recognize_google(audio, language='en-in').lower()
                print(f"User: {query}")
                if 'jarvis' in query:
                    speak("I am listening, sir.")
                    return
            except sr.WaitTimeoutError:
                continue
            except sr.UnknownValueError:
                continue
            except sr.RequestError:
                speak("Check your internet connection.")
                return

# === Greet ===
def wishMe():
    hour = int(datetime.datetime.now().hour)
    greeting = "Good Morning Rumi!" if 0 <= hour < 12 else "Good Afternoon Rumi!" if 12 <= hour < 18 else "Good Evening Rumi!"
    speak(f"{greeting} I am JARVIS, your personal assistant. How may I assist you?")
    notify("JARVIS Activated", "Your assistant is ready!")

# === Match Synonyms ===
def match_synonym(query, action):
    return any(word in query.lower() for word in synonyms.get(action, [action]))

# === PC Control Functions ===
def openAnything(query):
    app = query
    for phrase in synonyms["open"]:
        app = app.replace(phrase, '').strip()
    app = app_map.get(app.lower(), app)
    try:
        os.system(f"start {app}")
        speak_response("open", app)
        context["last_app"] = app
    except Exception as e:
        speak_response("error")
        logging.error(f"Open app error: {e}")

def closeApp(query):
    app = query
    for phrase in synonyms["close"]:
        app = app.replace(phrase, '').strip()
    app = app_map.get(app.lower(), app)
    for proc in psutil.process_iter(['name']):
        try:
            if app.lower() in proc.name().lower():
                proc.kill()
                speak_response("close", app)
                return
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    speak(f"Could not find {app} running.")

def manage_window(query):
    app = (context["last_app"] or context["last_query"]).strip().lower()
    windows = [win for win in gw.getAllTitles() if app in win.lower()]
    if not windows:
        speak(f"No window found for {app}")
        return
    if match_synonym(query, 'maximize'):
        for win in gw.getWindowsWithTitle(windows[0]):
            try:
                win.maximize()
                speak(f"Maximized {app}")
                return
            except:
                speak(f"Could not maximize {app}")
    elif match_synonym(query, 'minimize'):
        for win in gw.getWindowsWithTitle(windows[0]):
            try:
                win.minimize()
                speak(f"Minimized {app}")
                return
            except:
                speak(f"Could not minimize {app}")
    elif 'switch to' in query:
        pyautogui.hotkey('alt', 'tab')
        speak(f"Switched to {app}")

def type_text(query):
    text = query
    for phrase in synonyms["type"]:
        text = text.replace(phrase, '').strip()
    pyautogui.write(text.strip())
    speak(f"Typed {text}")

def control_volume(query):
    if 'mute' in query:
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            volume = session._ctl.QueryInterface(ISimpleAudioVolume)
            volume.SetMute(1, None)
        speak("Muted")
    elif 'volume up' in query:
        pyautogui.press('volumeup')
        speak("Volume increased")
    elif 'volume down' in query:
        pyautogui.press('volumedown')
        speak("Volume decreased")

def clipboard_ops(query):
    if match_synonym(query, 'copy'):
        text = query
        for phrase in synonyms["copy"]:
            text = text.replace(phrase, '').strip()
        pyperclip.copy(text.strip())
        speak("Text copied to clipboard")
    elif 'paste' in query:
        pyautogui.hotkey('ctrl', 'v')
        speak("Pasted from clipboard")

# === File Operations ===
def file_ops(query):
    if match_synonym(query, 'find'):
        term = query
        for phrase in synonyms["find"]:
            term = term.replace(phrase, '').strip()
        speak("Searching files, please wait...")
        search_path = f"C:\\Users\\{os.getlogin()}\\Desktop\\*{term}*"  # Limited to Desktop
        files = glob.glob(search_path, recursive=True)
        if files:
            context["last_file"] = files[0]
            speak_response("find", os.path.basename(files[0]))
            return files[0]
        else:
            speak("No files found")
    elif match_synonym(query, 'move') and context["last_file"]:
        dest = query
        for phrase in synonyms["move"]:
            dest = dest.replace(phrase, '').strip()
        try:
            dest_path = os.path.join(f"C:\\Users\\{os.getlogin()}\\Desktop", dest)
            if not os.path.exists(os.path.dirname(dest_path)):
                os.makedirs(os.path.dirname(dest_path))
            shutil.move(context["last_file"], dest_path)
            speak(f"Moved to {dest}")
        except Exception as e:
            speak_response("error")
            logging.error(f"File move error: {e}")
    elif 'delete' in query and context["last_file"]:
        speak("Are you sure you want to delete this file?")
        confirm = takeCommand()
        if 'yes' in confirm:
            os.remove(context["last_file"])
            speak("File deleted")
            context["last_file"] = None
        else:
            speak("Deletion canceled")

# === Extra Features ===
def takePhoto():
    speak("Taking your photo. Get ready.")
    cam = cv2.VideoCapture(0)
    result, image = cam.read()
    if result:
        filename = f"photo_{int(time.time())}.png"
        cv2.imwrite(filename, image)
        speak(f"Photo saved as {filename}")
        notify("Photo Captured", filename)
    else:
        speak("Failed to access webcam")
    cam.release()

def takeScreenshot():
    speak("Taking screenshot.")
    screenshot = pyautogui.screenshot()
    filename = f"screenshot_{int(time.time())}.png"
    screenshot.save(filename)
    speak(f"Screenshot saved as {filename}")
    notify("Screenshot Taken", filename)

def stopMusic():
    speak("Stopping music players.")
    media_apps = ["vlc", "spotify", "wmplayer"]
    for proc in psutil.process_iter(['name']):
        if any(app in proc.info['name'].lower() for app in media_apps):
            try:
                proc.kill()
                speak(f"Closed {proc.info['name']}")
            except:
                pass
    speak("Music stopped")

def readPDF():
    file_path = context["last_file"] if context["last_file"] and context["last_file"].endswith('.pdf') else None
    if not file_path:
        speak("Please specify a PDF file or find one first. What is the file name?")
        file_name = takeCommand()
        if file_name != "None":
            file_path = f"C:\\Users\\{os.getlogin()}\\Desktop\\{file_name}.pdf"
    try:
        speak("Reading the PDF now.")
        with open(file_path, 'rb') as f:
            reader = PyPDF2.PdfReader(f)
            for page in reader.pages:
                text = page.extract_text()
                if text:
                    speak(text[:500])
                    time.sleep(1)
    except Exception as e:
        speak("Sorry, I couldn't read the PDF. Ensure the file exists.")
        logging.error(f"PDF read error: {e}")

def playYouTube(query):
    search = query.replace("play", "").replace("on youtube", "").strip()
    speak(f"Playing {search} on YouTube")
    pywhatkit.playonyt(search)

def askAI(prompt):  # Placeholder for future API
    speak("AI integration is not active. Please provide an API key or use another command.")
    logging.info("Attempted AI query without active API.")

def sendEmail(to, content):
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(EMAIL, PASSWORD)
        server.sendmail(EMAIL, to, content)
        server.quit()
        speak("Email has been sent.")
    except Exception as e:
        speak("Failed to send the email. Check your email settings.")
        logging.error(f"Email error: {e}")

def tellJoke():
    jokes = [
        "Why don’t scientists trust atoms? They might be up to something!",
        "My computer said it needed a break. I guess it’s overclocked!",
        "Why did the developer go broke? He used up all his cache!"
    ]
    speak(random.choice(jokes))

def searchGoogle(query):
    search_term = query.replace("search for", "").strip()
    speak(f"Searching Google for {search_term}")
    webbrowser.open(f"https://www.google.com/search?q={search_term}")

def getWeather(city="Lahore"):
    api_url = "https://wttr.in/" + city + "?format=3"
    try:
        res = requests.get(api_url)
        if res.status_code == 200:
            speak(f"The weather in {city} is {res.text}")
        else:
            speak("Couldn't fetch weather data")
    except:
        speak("Check your internet connection")

def takeNote():
    speak("What should I write?")
    note = takeCommand()
    if note != "None":
        with open("notes.txt", "a") as f:
            f.write(f"{datetime.datetime.now()} - {note}\n")
        speak("Note saved")
    else:
        speak("No note recorded")

def readNotes():
    if os.path.exists("notes.txt"):
        with open("notes.txt", "r") as f:
            notes = f.read()
            speak("Reading your notes.")
            print(notes)
    else:
        speak("No notes found")

def setAlarm(alarm_time):
    try:
        # Validate time format
        if not re.match(r'^\d{1,2}:\d{2}$', alarm_time):
            speak("Please use HH:MM format for the alarm")
            return
        speak(f"Setting alarm for {alarm_time}")
        schedule.every().day.at(alarm_time).do(lambda: speak("Wake up! This is your alarm."))
        speak("Alarm scheduled")
    except Exception as e:
        speak("Failed to set alarm")
        logging.error(f"Alarm error: {e}")

def sendWhatsappMessage():
    speak("To whom should I send the WhatsApp message?")
    name = takeCommand()
    contact_list = {
        "hamza": "+92xxxxxxxxx",
        "zain": "+92xxxxxxxxx",
        "tahir ": "+92xxxxxxxxx",
        "rehan": "+92xxxxxxxxx"
    }
    phone = contact_list.get(name.lower())
    if not phone:
        speak("Sorry, contact not found")
        return
    speak("What should I say?")
    message = takeCommand()
    now = datetime.datetime.now()
    try:
        pywhatkit.sendwhatmsg(phone, message, now.hour, now.minute + 1)
        speak("WhatsApp message scheduled")
    except Exception as e:
        speak("Failed to send WhatsApp message")
        logging.error(f"WhatsApp error: {e}")

# === Power Controls ===
def shutdown():
    speak("Are you sure you want to shutdown?")
    confirm = takeCommand()
    if 'yes' in confirm:
        os.system("shutdown /s /t 1")
        notify("System", "Shutting down now")
    else:
        speak("Shutdown canceled")

def restart():
    speak("Are you sure you want to restart?")
    confirm = takeCommand()
    if 'yes' in confirm:
        os.system("shutdown /r /t 1")
        notify("System", "Restarting now")
    else:
        speak("Restart canceled")

def sleep():
    speak("Putting system to sleep.")
    os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")

def hibernate():
    speak("Hibernating the system.")
    os.system("shutdown /h")

# === Scheduler Thread ===
def run_scheduler():
    while True:
        schedule.run_pending()
        time.sleep(1)
threading.Thread(target=run_scheduler, daemon=True).start()

# === GUI Interface ===
def launch_gui():
    root = tk.Tk()
    root.title("JARVIS Assistant")
    root.geometry("300x200")
    label = tk.Label(root, text="JARVIS Assistant", font=("Helvetica", 16))
    label.pack(pady=10)

    def run_jarvis():
        root.destroy()
        wishMe()
        while True:
            waitForWakeWord()
            main_loop()

    start_btn = tk.Button(root, text="Start JARVIS", command=run_jarvis, font=("Helvetica", 12))
    start_btn.pack(pady=20)
    root.mainloop()

# === Speak Response ===
def speak_response(action, response_data=None):
    response = random.choice(response_variants.get(action, ["Done"]))
    if response_data:
        response = response.format(app=response_data, result=response_data)
    speak(response)

# === Main Loop ===
def main_loop():
    query = takeCommand()
    if query == "None":
        speak("Sorry, I didn’t catch that. Try again, sir.")
        return

    # Intent Matching
    if match_synonym(query, 'open'):
        openAnything(query)
    elif match_synonym(query, 'close'):
        closeApp(query)
    elif match_synonym(query, 'maximize') or match_synonym(query, 'minimize') or 'switch to' in query:
        manage_window(query)
    elif match_synonym(query, 'type'):
        type_text(query)
    elif 'mute' in query or 'volume' in query:
        control_volume(query)
    elif match_synonym(query, 'copy') or 'paste' in query:
        clipboard_ops(query)
    elif match_synonym(query, 'find') or match_synonym(query, 'move') or 'delete' in query:
        file_ops(query)
    elif 'take note' in query:
        takeNote()
    elif 'read notes' in query:
        readNotes()
    elif 'schedule' in query:
        match = re.search(r'schedule (.+) at (\d{1,2}:\d{2})', query)
        if match:
            task, time = match.groups()
            schedule.every().day.at(time).do(lambda: os.system(task))
            speak(f"Scheduled {task} at {time}")
    elif 'shutdown' in query:
        shutdown()
    elif 'restart' in query:
        restart()
    elif 'sleep' in query:
        sleep()
    elif 'hibernate' in query:
        hibernate()
    elif 'take photo' in query:
        takePhoto()
    elif 'take screenshot' in query:
        takeScreenshot()
    elif 'stop music' in query:
        stopMusic()
    elif 'play' in query and 'youtube' in query:
        playYouTube(query)
    elif 'ask ai' in query or 'ask chatgpt' in query:
        speak("What would you like to ask?")
        prompt = takeCommand()
        askAI(prompt)
    elif 'email to' in query:
        match = re.search(r'email to (\w+)', query)
        if match:
            recipient = match.group(1).lower()
            contacts = {"harry": "harryyourEmail@gmail.com"}
            to = contacts.get(recipient)
            if to:
                speak("What should I say?")
                content = takeCommand()
                sendEmail(to, content)
            else:
                speak("Recipient not found.")
    elif 'tell joke' in query:
        tellJoke()
    elif 'search for' in query:
        searchGoogle(query)
    elif 'weather' in query:
        match = re.search(r'weather in ([a-zA-Z ]+)', query)
        city = match.group(1) if match else "Lahore"
        getWeather(city)
    elif 'read pdf' in query:
        readPDF()
    elif 'set alarm' in query:
        match = re.search(r'set alarm for (\d{1,2}:\d{2})', query)
        if match:
            setAlarm(match.group(1))
        else:
            speak("Please say time in format HH:MM")
    elif 'send whatsapp' in query:
        sendWhatsappMessage()
    elif 'notify me' in query:
        speak("What message should I notify you?")
        msg = takeCommand()
        notify("JARVIS Notification", msg)
    elif 'time' in query:
        strTime = datetime.datetime.now().strftime("%H:%M:%S")
        speak(f"Sir, the time is {strTime}")
    elif 'wikipedia' in query:
        speak("Searching Wikipedia...")
        query = query.replace("wikipedia", "").strip()
        try:
            results = wikipedia.summary(query, sentences=2)
            speak("According to Wikipedia")
            speak(results)
        except:
            speak("No results found on Wikipedia")
    elif 'exit' in query or 'stop' in query or 'quit' in query:
        speak("Goodbye, sir. Until next time!")
        sys.exit()
    else:
        speak("Sorry, I didn’t understand that. Could you rephrase, please?")

# === Start ===
if __name__ == '__main__':
    launch_gui()      
    
