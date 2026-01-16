import time
import tempfile
import os
import datetime
import webbrowser
import wikipedia
import re
import subprocess
import sys

import sounddevice as sd
import soundfile as sf
import speech_recognition as sr

# ---------------------------
# Configuration
# ---------------------------
CHROME_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"  
VSCODE_PATH = r"C:\Users\{USERNAME}\AppData\Local\Programs\Microsoft VS Code\Code.exe" 
DEFAULT_RECORD_SECONDS = 4


_launched = {}  

# ---------------------------
# TTS: prefer Windows SAPI, fallback to pyttsx3
# ---------------------------
def _sapi_speak(text):
    import win32com.client
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    speaker.Speak(text)

def _pyttsx3_speak(text, rate=160):
    import pyttsx3
    engine = pyttsx3.init()
    try:
        engine.setProperty('rate', rate)
    except Exception:
        pass
    try:
        voices = engine.getProperty('voices')
        for v in voices:
            if 'male' in v.name.lower() or 'male' in v.id.lower():
                engine.setProperty('voice', v.id)
                break
    except Exception:
        pass
    engine.say(text)
    engine.runAndWait()
    try:
        engine.stop()
    except Exception:
        pass

def speak(text):
   
    print(f"Assistant (speaking): {text}")
    try:
        _sapi_speak(text)
        return
    except Exception:
      
        try:
            _pyttsx3_speak(text)
            return
        except Exception as e:
            print("TTS failed:", e)
            return

# ---------------------------
# Utilities: clean & extract topic
# ---------------------------
def clean_topic(topic):
    if not topic:
        return ""
    topic = topic.strip().lower()
   
    topic = re.sub(r"^(what is|what's|what are|who is|tell me about|define|explain)\s+", "", topic)
   
    topic = re.sub(r"^(a|an|the)\s+", "", topic)
    
    topic = topic.strip(" ?!.")
    return topic

def wikipedia_summary_for(topic):
   
    if not topic:
        return None
    try:
        
        results = wikipedia.search(topic, results=5)
        if not results:
            return None
        page_title = results[0]
       
        summary = wikipedia.summary(page_title, sentences=2)
        return summary
    except wikipedia.exceptions.DisambiguationError as e:
        
        for opt in e.options:
            try:
                s = wikipedia.summary(opt, sentences=2)
                return s
            except Exception:
                continue
        return None
    except wikipedia.exceptions.PageError:
        return None
    except Exception:
        return None

# ---------------------------
# Application management
# ---------------------------
def open_app(app_name):
    key = app_name.lower().strip()
   
    p = _launched.get(key)
    if p and p.poll() is None:
        return f"{app_name} is already open."

 
    if key == "chrome":
        if os.path.exists(CHROME_PATH):
            proc = subprocess.Popen([CHROME_PATH])
            _launched[key] = proc
            return "Opening Chrome..."
        else:
            webbrowser.open("https://www.google.com")
            return "Opening web browser (Chrome not found)."
    if key == "vscode" or key == "code":
        if os.path.exists(VSCODE_PATH):
            proc = subprocess.Popen([VSCODE_PATH])
            _launched[key] = proc
            return "Opening VS Code..."
        else:
            return "VS Code path not found; please update VSCODE_PATH in the script."

   
    if key == "youtube":
       
        if os.path.exists(CHROME_PATH):
            proc = subprocess.Popen([CHROME_PATH, "https://www.youtube.com"])
            _launched[key] = proc
            return "Opening YouTube..."
        else:
            webbrowser.open("https://www.youtube.com")
            
            _launched[key] = None
            return "Opened YouTube in your default browser (cannot  close this)."

 
    if key.startswith("http"):
        webbrowser.open(app_name)
        _launched[key] = None
        return f"Opened {app_name} in browser."
    return f"Application or shortcut '{app_name}' not recognised."

def close_app(app_name):
    key = app_name.lower().strip()
    p = _launched.get(key)
    if p is None:
        return f"No program was launched by this assistant for '{app_name}', or it was opened via default browser (can't close)."
   
    try:
        if p.poll() is None:
            p.terminate()
            time.sleep(0.3)
            if p.poll() is None:
                p.kill()
        del _launched[key]
        return f"Closed {app_name}."
    except Exception as e:
        return f"Failed to close {app_name}: {e}"

# ---------------------------
# Command functions
# ---------------------------
def tell_time():
    now = datetime.datetime.now()
    return f"It’s {now.strftime('%I:%M %p')}."

def tell_date():
    today = datetime.datetime.now()
    return f"Today’s date is {today.strftime('%B %d, %Y')}."

def search_web(query):
    if not query:
        return "No search query provided."
    webbrowser.open(f"https://www.google.com/search?q={query}")
    return f"Here’s what I found for '{query}'."

def fetch_wikipedia(topic):
    topic = clean_topic(topic)
    if not topic:
        return None
    summary = wikipedia_summary_for(topic)
    return summary

def process_command(command):
    command_lower = (command or "").lower().strip()

   
    if command_lower.startswith("close ") or command_lower.startswith("shutdown "):
        target = re.sub(r"^(close|shutdown)\s+", "", command_lower)
        return close_app(target)

   
    if command_lower.startswith("open ") or command_lower.startswith("launch "):
        target = re.sub(r"^(open|launch)\s+", "", command_lower)
        return open_app(target)

  
    if any(kw in command_lower for kw in ["time", "what time", "current time"]):
        return tell_time()
    if any(kw in command_lower for kw in ["date", "today's date", "what is the date"]):
        return tell_date()
    if command_lower.startswith("search "):
        q = command_lower.replace("search", "", 1).strip()
        return search_web(q)

    
    if re.search(r"\b(what is|what's|what are|who is|tell me about|define|explain)\b", command_lower):
        raw_topic = command_lower
        topic = clean_topic(raw_topic)
       
        if not topic:
            return "What would you like me to search for?"
        summary = fetch_wikipedia(topic)
        if summary:
            return summary
        else:
           
            webbrowser.open(f"https://www.google.com/search?q={topic}")
            return f"Couldn't find a good Wikipedia summary for '{topic}'. I opened a web search instead."

   
    if any(g in command_lower for g in ["hello", "hi", "hey"]):
        return "Hello! How can I help you today?"
    if "how are you" in command_lower:
        return "I’m doing great! I can assist you with tasks and information."
    if command_lower in ["stop", "quit", "exit"]:
        return "stop"

   
    return None  

# ---------------------------
# Continuous listening loop
# ---------------------------
def listen_loop(interval=DEFAULT_RECORD_SECONDS, device=None, enable_text_fallback=True):
   
    recognizer = sr.Recognizer()
    speak("Assistant is now listening...")
    print(">> Starting listen loop. Press Ctrl+C to stop.")

    while True:
        
        time.sleep(0.12)

        with tempfile.NamedTemporaryFile(delete=False, suffix=".wav") as tmpfile:
            tmp_path = tmpfile.name

        command_text = ""  

        try:
            if device:
                sd.default.device = device

            print(f"Listening for {interval} seconds...")
            recording = sd.rec(int(interval * 44100), samplerate=44100, channels=1)
            sd.wait()
            sf.write(tmp_path, recording, 44100)

            try:
                with sr.AudioFile(tmp_path) as source:
                    audio = recognizer.record(source)

                command_text = recognizer.recognize_google(audio)
                print(f"Captured Input: {command_text}")

            except sr.UnknownValueError:
                
                print("Captured Input: Could not understand audio")
                speak("I couldn't understand you. Please say that again.")
               
            except sr.RequestError as e:
                print("Captured Input: Request error:", e)
                speak("There was a problem with the speech recognition service. I'll try again.")
               
            except Exception as e:
                print("Captured Input: Other error:", e)
                speak("An error occurred while capturing audio. I'll keep listening.")
               

        finally:
            try:
                os.remove(tmp_path)
            except Exception:
                pass

       
        if not command_text:
            continue

       
        response = process_command(command_text)

        if response == "stop":
            speak("Goodbye!")
            break

        if response is None:
          
            speak("I didn’t understand that. Could you please repeat?")
           
            continue

       
        speak(response)


if __name__ == "__main__":
    try:
        listen_loop(interval=4)
    except KeyboardInterrupt:
        print("\nStopped by user.")
        speak("Stopped. Goodbye!")
 