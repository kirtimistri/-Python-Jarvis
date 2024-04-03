import datetime
import os
from tkinter import *
import webbrowser
import win32com.client
import speech_recognition as sr

speaker = win32com.client.Dispatch('SAPI.SpVoice')
def say(text):
    speaker.Speak(text)
def greetme():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<=12:
        say("good Morning")
    elif hour>12 and  hour<=18 :
        say("good Evening")
    else:
        say("good Afternoon")
    say(" how can i help you ")

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 0.6
        r.energy_threshold = 300
        audio = r.listen(source,0,4)
    try:
        print('Recognizing.....')
        query = r.recognize_google(audio, language='en-IN')
        say(f'user said: {query}')
        return query
    except Exception as e:
        print(e)
        return 'Sorry, could you please repeat?'

if __name__ == '__main__':
    print("PyCharm")
    say("Jarvis Activated")
    greetme()

    while True:
        print('Listening . . .')
        query = takeCommand()
        sites = [
            ["YouTube","https://www.youtube.com/"],
            ["Wikipedia", "https://www.wikipedia.com"],
            ["Google", "https://www.google.com"],
            ["Kirti's Instagram id", "https://www.instagram.com/?hl=en"],
            ["Indala college website", "https://icoe.ac.in/"],
            ["SP Instagram id", "https://www.instagram.com/suvesh_05?igsh=MWQ2a3Bycmc2NGdi"]
        ]

        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                webbrowser.open(site[1])
                say(f"Opening {site[0]} sir.....")
        if "play music" in query:
            musicPath = r"C:\Users\asus\Music\One_Love_[_LoFi_+_Slowed_+_Reverb_]_-_Shubh__One_love_shubh_slowed_reverb(256k).mp3"
            os.system(f"start {musicPath}")
            say(f"ok")
        if "play video" in query:
            videoPath = r"C:\Users\asus\Videos\TRANCE___travis_end_part_____EDIT(1080p).mp4"
            os.system(f"start {videoPath}")
            say(f"ok")
        if "open gallery" in query:
            photos = r""
            os.system(f"start {photos}")
            say(f"ok")
        if "what is the time" in query:
            strfTime=datetime.datetime.now().strftime("%H hours %M minute and %S seconds")
            say(f"mam the time is{strfTime}")
        if "open WhatsApp" in query:
           WhatsApp = r"C:\Users\asus\OneDrive\Desktop\WhatsApp.lnk"
           os.system(f"start {WhatsApp}")
           say(f"ok")
        if "open camera" in query:
           cam = r"C:\Users\Public\Desktop\iVCam.lnk"
           os.system(f"start {cam}")
           say(f"ok")
