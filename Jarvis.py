import win32com.client
import datetime
import speech_recognition as sr
import wikipedia
import webbrowser
import subprocess
import os

speaker = win32com.client.Dispatch("SAPI.SpVoice") 


def wishme():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speaker.Speak("Good Morning!")
    elif hour>=12 and hour<18:
        speaker.Speak("Good Afternoon!")
    else:
        speaker.Speak("Good Evening!")
    speaker.Speak("I am Jarvis Sir. Please tell me how may I help you")
    
def takecommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening....")
        r.adjust_for_ambient_noise(source,duration=1)
        audio = r.listen(source) 
    try:
        print("Recognizing....")
        text = r.recognize_google(audio,language='en-in') 
        print(f"User said: {text}\n") 
    except:
        print("Say that again please....")
        return "None" 
    return text
    
if __name__=="__main__":
    wishme()
    while True:
    #if 1:
        text = takecommand().lower()
        
        #Logic for executing tasks based on query
        if "wikipedia" in text:
            speaker.Speak("Searching Wikipedia....")
            text = text.replace("wikipedia","")
            results = wikipedia.summary(text,sentences=2)
            speaker.Speak("According to Wikipedia")
            print(results)
            speaker.Speak(results)
        elif "open youtube" in text:
            webbrowser.open("youtube.com")
        elif "open google" in text:
            webbrowser.open("google.com")
        elif "open stackoverflow" in text:
            webbrowser.open("stackoverflow.com")    
        elif "play music" in text: 
            os.startfile("Breathless Song.mp3")
        elif "the time" in text:
            strTime = datetime.datetime.now().strftime("%H:%M:%S") 
            speaker.Speak(f"Sir, the time is {strTime}")
        elif "open calculator" in text:
            subprocess.Popen("C:\\Windows\\System32\\calc.exe")
        elif "quit" in text:
            exit()           
            
        
        
