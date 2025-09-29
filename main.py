import speech_recognition as sr
import os
import pyttsx3
import webbrowser
import datetime
import subprocess
import sys

moviePath = r"C:\Users\KIIT\Downloads\Ramaiya_Vastavaiya_(2013)_480p_BluRay.mp4"

WEBSITES = {
    "youtube": "https://www.youtube.com",
    "wikipedia": "https://www.wikipedia.org",
    "google": "https://www.google.com",
    "whatsapp_web": "https://web.whatsapp.com",
}

VSCODE_PATHS = [
    os.path.join(os.environ.get("LOCALAPPDATA", ""), r"Programs\Microsoft VS Code\Code.exe"),
    r"C:\Program Files\Microsoft VS Code\Code.exe",
    r"C:\Program Files (x86)\Microsoft VS Code\Code.exe",
]

WHATSAPP_PATHS = [
    os.path.join(os.environ.get("LOCALAPPDATA", ""), r"WhatsApp\WhatsApp.exe"),
    os.path.join(os.environ.get("LOCALAPPDATA", ""), r"Programs\WhatsApp\WhatsApp.exe"),
    r"C:\Program Files\WindowsApps\WhatsApp.exe",  # often inaccessible, kept for completeness
]

engine = pyttsx3.init()

def say(text):
    engine.say(text)
    engine.runAndWait()

def listen(timeout=6, phrase_time_limit=8):
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.adjust_for_ambient_noise(source, duration=0.5)
        try:
            audio = r.listen(source, timeout=timeout, phrase_time_limit=phrase_time_limit)
        except sr.WaitTimeoutError:
            return ""
    try:
        return r.recognize_google(audio, language="en-in")
    except sr.UnknownValueError:
        return ""
    except sr.RequestError as e:
        print("Speech recognition request error:", e)
        return ""
    except Exception as e:
        print("Recognition exception:", e)
        return ""

def open_website_by_name(name):
    url = WEBSITES.get(name)
    if url:
        webbrowser.open(url, new=2)  # open in new tab
        return True
    return False

def try_start_executable(path):
    try:
        if os.path.exists(path):
            os.startfile(path)
            return True
    except Exception as e:
        print("os.startfile failed:", e)
    try:
        subprocess.Popen([path])
        return True
    except Exception as e:
        print("subprocess start failed:", e)
        return False

def open_vscode():
    try:
        subprocess.Popen(["code"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return True
    except Exception:
        for p in VSCODE_PATHS:
            if try_start_executable(p):
                return True
    return False

def open_whatsapp():
    for p in WHATSAPP_PATHS:
        if try_start_executable(p):
            return True
    try:
        subprocess.Popen(["explorer", "whatsapp:"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return True
    except Exception as e:
        print("whatsapp protocol/open via explorer failed:", e)
    webbrowser.open(WEBSITES["whatsapp_web"], new=2)
    return False

def play_movie():
    if os.path.exists(moviePath):
        try:
            os.startfile(moviePath)
            return True
        except Exception as e:
            print("Could not start movie with os.startfile:", e)
            try:
                subprocess.Popen([moviePath])
                return True
            except Exception as e2:
                print("subprocess failed:", e2)
                return False
    else:
        print("Movie path doesn't exist:", moviePath)
        return False

if __name__ == "__main__":
    say("Hello. I am Jarvis AI.")
    while True:
        print("Listening...")
        query = listen()
        if not query:
            continue
        q = query.lower().strip()
        print("Heard:", q)

        if any(w in q for w in ("exit", "quit", "stop", "bye")):
            say("Goodbye.")
            break

        if "time" in q:
            now = datetime.datetime.now()
            say(f"The time is {now.strftime('%I:%M %p')}")
            continue

        if "youtube" in q and any(x in q for x in ("open", "launch", "play")):
            say("Opening YouTube.")
            open_website_by_name("youtube")
            continue

        if "wikipedia" in q and any(x in q for x in ("open", "search", "launch")):
            say("Opening Wikipedia.")
            open_website_by_name("wikipedia")
            continue

        if "google" in q and any(x in q for x in ("open", "search", "launch")):
            say("Opening Google.")
            open_website_by_name("google")
            continue

        if "whatsapp" in q and any(x in q for x in ("open", "launch", "start")):
            say("Opening WhatsApp.")
            opened = open_whatsapp()
            if not opened:
                say("Couldn't open the WhatsApp app. Opening WhatsApp Web instead.")
            continue

        if any(phrase in q for phrase in ("vs code", "visual studio code", "vscode")) and any(x in q for x in ("open", "launch", "start")):
            say("Opening Visual Studio Code.")
            if not open_vscode():
                say("VS Code could not be launched. Please ensure it is installed or 'code' is in PATH.")
            continue

        if any(x in q for x in ("open movie", "play movie", "play film")):
            if play_movie():
                say("Playing movie.")
            else:
                say("Could not find or open the movie file. Please check the movie path in the script.")
            continue

        if "open" in q:
            tokens = q.split()
            for token in tokens[::-1]:
                if token in WEBSITES:
                    webbrowser.open(WEBSITES[token], new=2)
                    say(f"Opening {token}.")
                    break
            else:
                say("I couldn't map that command to an app or website. Try 'open YouTube' or 'open VS Code'.")
            continue

        if "what can you do" in q or "help" in q:
            say("I can open YouTube, Google, Wikipedia, WhatsApp app or web, launch Visual Studio Code, tell the time, and play a saved movie.")
            continue

        say("Sorry, I didn't understand. Try saying 'open YouTube' or 'play movie'.")
#say(query)

# import win32com.client
# speaker=win32com.client.Dispatch("SAPI.SpVoice")
# while 1:
#     print("Enter the word you want to speak it out by computer")
#     s=input()
#     speaker.Speak(s)

