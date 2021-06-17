import subprocess
import gtts
import wolframalpha
import pyttsx3
import tkinter
import json
import random
import operator
import speech_recognition as sr
import datetime
import psutil
import pyautogui
import wikipedia
import webbrowser
import os
import winshell
import feedparser
import smtplib
import ctypes
import time
import requests
import shutil
import pywhatkit
import playsound
from twilio.rest import Client
from clint.textui import progress
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen


# For understanding speech
import speech_recognition as sr

# For fetching the answers
# to computational queries
import wolframalpha

# for fetching wikipedia articles
import wikipedia

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)


def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def voice_change(v):
    x = int(v)
    engine.setProperty('voice', voices[x].id)
    speak("changed the voice sir")


def screenshot():
    img = pyautogui.screenshot()
    img.save("C:\\Users\\PRAVEEN KUMAR A\\PycharmProjects\\pythonProject4\\ss.png")



def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good Morning Sir !")

    elif hour >= 12 and hour < 18:
        speak("Good Afternoon Sir !")

    else:
        speak("Good Evening Sir !")

    assname = ("tesla")
    speak("I am your Assistant")
    speak(assname)


def usrname():
    speak("What should i call you sir")
    uname = takeCommand()
    speak("Welcome Mister")
    speak(uname)
    columns = shutil.get_terminal_size().columns

    #print("#####################".center(columns))
    print("Welcome Mr.", uname)#ter(columns))
    #print("#####################".center(columns))

    speak("How can i Help you, Sir")


def takeCommand():
    r = sr.Recognizer()

    with sr.Microphone() as source:

        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        print(e)
        print("Unable to Recognize your voice.")
        return "None"

    return query




}


if __name__ == '__main__':
    clear = lambda: os.system('cls')

    # This Function will clean any
    # command before execution of this python file
    clear()
    wishMe()
    usrname()

    while True:

        query = takeCommand().lower()

        # All the commands said by user will be
        # stored here in 'query' and will be
        # converted to lower case for easily
        # recognition of command
        if 'wikipedia' in query:
            webbrowser.open("https://wikipedia.org/")


        elif 'open youtube' in query:
            speak("Here you go to Youtube\n")
            webbrowser.open("youtube.com")
            time.sleep(9)

        elif 'open google' in query:
            speak("Here you go to Google\n")
            webbrowser.open("google.com")
            time.sleep(9)

        elif 'open gmail' in query:
            speak("opening gmail\n")
            webbrowser.open("https://mail.google.com/mail/u/0/?tab=rm&ogbl#inbox")
            time.sleep(5)

        elif "open word" in query:
            speak("Opening Microsoft Word")
            os.startfile('"C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE"')
            time.sleep(5)

        elif "open linkedin" in query:
            speak("opening linkedin\n")
            webbrowser.open("https://www.linkedin.com/in/praveenkumar-a-8638531a0")
            time.sleep(5)

        elif "open skillrack" in query:
            speak("opening  your skillrack account\n")
            webbrowser.open("https://www.skillrack.com/faces/ui/profile.xhtml;jsessionid=AA598F06F18FD25CD27EFA289A75D288")
            time.sleep(7)

        elif ("screenshot" in query):
            screenshot()
            speak("screenshot taken sir!")


        elif 'music' in query:
            speak("Here you go with your favourite music")
            # music_dir = "G:\\Song"
            music_dir = "C:\\Users\\PRAVEEN KUMAR A\\Music"
            songs = os.listdir(music_dir)
            print(songs)
            random = os.startfile(os.path.join(music_dir, songs[30]))
            time.sleep(10)

        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            speak(f"Sir, the time is {strTime}")

        elif 'how are you' in query:
            speak("I am fine, Thank you")
            speak("How are you, Sir")

        elif 'i am fine' in query or "i am good" in query:
            speak("It's good to know that your fine")

        elif 'exit' in query or 'bye' in query:
            speak("Thanks for giving me your time")
            exit()

        elif "who made you" in query or "who created you" in query:
            speak("I have been created by PRAVEEN.")

        elif ("voice" in query):
            speak("for female say female and, for male say boy")
            b = takeCommand()
            if ("female" in b):
                voice_change(1)
            elif ("boy" in b):
                voice_change(0)

        elif "calculate" in query:
            speak("what?")
            g=takeCommand().lower().rstrip().strip()
            k=[i for i in g if i!=" "]
            k="".join(k)
            try:
                answer=eval(k)
                print("The answer is " ,answer)
                speak("The answer is " + str(answer))
            except:
                speak("Please give the crt question..... Command invalid")


        elif 'search' in query or 'what is' in query or 'who is' in query:
            person = query.replace('what is', '')
            person = query.replace('who is', '')
            query = query.replace('search', '')
            info = wikipedia.summary(person, 2)
            pywhatkit.search(person)
            pywhatkit.search(query)
            print(info)
            speak(info)

        elif "who am i" in query:
            speak("If you talk then definately your human.")

        elif "why you came to the world" in query:
            speak("Thanks to PRAVEEN. further It's a secret")

        elif 'power point presentation' in query:
            speak("opening Power Point presentation")
            power = r"C:\\Users\\PRAVEENKUMAR\\Desktop\\Minor Project\\Presentation\\Voice Assistant.pptx"
            os.startfile(power)

        elif "who are you" in query:
            speak("I am tesla a virtual assistant created by PRAVEEN")

        elif 'reason for you' in query:
            speak("I was created as a Mini project by Mister PRAVEENKUMAR")


        if 'play' in query:
            song = query.replace('play', '')
            speak('playing ' + song)
            pywhatkit.playonyt(song)


        elif "where is" in query:
            query = query.replace("where is", "")
            location = query
            speak("User asked to Locate")
            speak(location)
            webbrowser.open("https://www.google.com/maps/place/" + location.strip().rstrip())

        elif "camera" in query or "take a photo" in query:
            ec.capture(0,"tesla Camera", "img.jpg")

        elif ("weather" in query or "temperature" in query):
            weather()

       

       




