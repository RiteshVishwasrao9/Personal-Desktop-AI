import speech_recognition as sr
import os
import pyttsx3
import win32com.client
import webbrowser
import openai
import datetime
import pygame
from gtts import gTTS
import pyjokes
import requests

#---------------- SPeaker Testing ---------------------------------------------------------------------------------------

# speaker = win32com.client.Dispatch("SAPI.SpVoice")
#
# while 1:
#     print("Enter th word you want to speak it out bu computer")
#     s = input()
#     speaker.Speak(s)

#------------ ALL API's Here -----------------------------------------------------------------------------------------------

newsapi = "<API Key Here>"

#------------- ALL FUNCTIONS -----------------------------------------------------------------------------------------------

def say_old(text):
    engine = pyttsx3.init()
    engine.say(text)
    engine.runAndWait()

# --------------------------------------------------------------------------------------------------------------------------
def say(text):
    tts = gTTS(text)
    tts.save('temp.mp3') 

    # Initialize Pygame mixer
    pygame.mixer.init()

    # Load the MP3 file
    pygame.mixer.music.load('temp.mp3')

    # Play the MP3 file
    pygame.mixer.music.play()

    # Keep the program running until the music stops playing
    while pygame.mixer.music.get_busy():
        pygame.time.Clock().tick(10)
    
    pygame.mixer.music.unload()
    os.remove("temp.mp3") 

# ------------------------------------------------------------------------------------------


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 1
        audio = r.listen(source)
        try:
            print("Recoginizing.....")
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query
        except Exception as e:
            return "Some error Occurred. Sorry From AI"


#------------------ MAIN FUnctON --------------------------------------------------------------------------

if __name__ == '__main__':
    print("Hello")
    say('Hello I am You are Assistant AI')
    while True:
        print("Listenning....")
        query = takeCommand()

#--------------- For Web Surfing -----------------------------------------------------------------------------

        sites = [["youtube","https://www.youtube.com/"],["google","https://www.google.com/"],["instagram","https://www.instagram.com"],
                 ["wikipedia","https://www.wikipedia.org"]]
        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                say(f"Openning {site[0]} Sir..")
                webbrowser.open(site[1])

#-----------------For Desktop Surfing -------------------------------------------------------------------------

        if "play music" in query:
            musicpath = "D:\Work\WORK R\FILM & NATAK WORK\BGM's\Songs\Ya-Re-Ya-Rohan-Pradhan.mp3"
            say("Music Playing Sir")
            os.startfile(musicpath)
            

        elif "time" in query:
            strfTime = datetime.datetime.now().strftime("%H:%M:%S")
            say(f"Sir the Time is {strfTime}")

        elif "open code".lower() in query.lower():
            path =r"C:\Users\rites\AppData\Local\Programs\Microsoft VS Code\Code.exe"
            say("openning Visual Studio Code Sir")
            os.startfile(path)

        elif "open work".lower() in query.lower():
            paths ="D:\Work\WORK R"
            say("openning Work Folder Sir")
            os.startfile(paths)

        # elif "anaconda".lower() in query.lower():
        #     conda=r"C:\Users\rites\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Anaconda3 (64-bit)\Anaconda Prompt.exe"
        #     say("Openning Anaconda Prompt sir..")
        #     os.startfile(conda)
            

        elif "joke".lower() in query.lower():
            joke = pyjokes.get_joke()
            print(joke)
            say(joke)

#---------------- News Section -------------------------------------------------------------------------------------------

        elif "news".lower() in query.lower():
            r = requests.get(f"https://newsapi.org/v2/top-headlines?country=in&apiKey={newsapi}")
            if r.status_code == 200:
                # Parse the JSON response
                data = r.json()
            
                # Extract the articles
                articles = data.get('articles', [])
            
                # Print the headlines
                for article in articles:
                    say(article['title'])

#----------------- AI Name Called Here -----------------------------------------------------------------------------------

        elif "pda".lower() in query.lower():
            say("Yes Sir..")

#----------------- Shut Down Called Here -----------------------------------------------------------------------------------

        elif "shut down".lower() in query.lower():
            say("Shutting Down sir...")
            exit()
