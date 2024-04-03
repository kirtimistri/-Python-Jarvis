import datetime
import os
import tkinter as tk
import webbrowser
import win32com.client
import speech_recognition as sr

speaker = win32com.client.Dispatch('SAPI.SpVoice')
activated = False  # Flag to indicate whether Jarvis is activated

def say(text):
    speaker.Speak(text)

def greetme():
    hour = int(datetime.datetime.now().hour)
    if 0 <= hour < 12:
        say("Good Morning")
    elif 12 <= hour < 18:
        say("Good Afternoon")
    else:
        say("Good Evening")
    say("How can I help you?")

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 0.6
        r.energy_threshold = 300
        audio = r.listen(source, 0, 4)
    try:
        print('Recognizing.....')
        query = r.recognize_google(audio, language='en-IN')
        say(f'User said: {query}')
        return query
    except Exception as e:
        print(e)
        return 'Sorry, could you please repeat?'

def activateJarvis():
    global activated
    while not activated:
        print("Listening for activation command...")
        query = takeCommand()
        if "activate" in query.lower():  # Change the activation command as needed
            activated = True
            print("Jarvis activated.")
            say("Jarvis activated.")
            greetme()

if __name__ == '__main__':
    activateJarvis()

    while activated:
        print('Listening . . .')
        query = takeCommand()
        sites = [
            ["YouTube", "https://www.youtube.com/"],
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
            say(f"Playing music.")

        if "play video" in query:
            videoPath = r"C:\Users\asus\Videos\TRANCE___travis_end_part_____EDIT(1080p).mp4"
            os.system(f"start {videoPath}")
            say(f"Playing video.")

        if "open gallery" in query:
            photos = r""  # Add your gallery path here
            os.system(f"start {photos}")
            say(f"Opening gallery.")

        if "what is the time" in query:
            strfTime = datetime.datetime.now().strftime("%H hours %M minute and %S seconds")
            say(f"The time is {strfTime}")

        if "open WhatsApp" in query:
            WhatsApp = r"C:\Users\asus\OneDrive\Desktop\WhatsApp.lnk"
            os.system(f"start {WhatsApp}")
            say(f"Opening WhatsApp.")

        if "open camera" in query:
            cam = r"C:\Users\Public\Desktop\iVCam.lnk"
            os.system(f"start {cam}")
            say(f"Opening camera.")

        if "deactivate" in query:
            say("Deactivating Jarvis sir have a good day.")
            print("Jarvis Dactivated.")
            activated = False
            break
