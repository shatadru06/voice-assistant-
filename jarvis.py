import datetime
import email
from email.message import EmailMessage
from platform import release
import pyttsx3
from requests import request
import speech_recognition as sr
import pywhatkit
from keyboard import press_and_release
from keyboard import press
from keyboard import write
import requests
from bs4 import BeautifulSoup
import pyautogui
import psutil
import webbrowser
import smtplib
import pandas as pd 
from email.message import EmailMessage
           

time = datetime.datetime.now().strftime("%H")
time_mow =datetime.datetime.now().strftime(" %M")
engine = pyttsx3.init('sapi5')
voices=engine.getProperty('voices')
print(voices[1].id)
engine.setProperty('voice',voices[1].id)
def weatherForcast(query):
    city = query.split('weather')
    search = "weather"+city[1]
    url = f"https://www.google.com/search?q={search}"
    r=requests.get(url)
    data = BeautifulSoup(r.text,"html.parser")
    temp= data.find("div",class_="BNeawe iBp4i AP7Wnd").text
    weather= data.find("div",class_="BNeawe tAd8D AP7Wnd").text
    dataa = weather.split('\n')
    time2=dataa[0]
    sky= dataa[1]
    print(sky)
    speak(f"current {search} is {temp}...and the sky description is ...{sky}.... at {time2}")        
def searchEngine(query):
    about=query.split('search')
    webbrowser.open(f"https://www.google.com/search?q={about[1]}")
def sendEmail(to,content):
    server =smtplib.SMTP('smtp.gmail.com',587)
    server.ehlo()
    server.starttls()
    server.login('shatadrusjarvis@gmail.com','kdoqmiusjsxpxsrh')
    email = EmailMessage()
    email['from']='shatadrusjarvis@gmail.com'
    email['to'] = to
    email['subject'] = subject
    email.set_content(content)
    server.send_message(email)
    server.close()
def speak(audio):
    engine.say(audio)
    engine.runAndWait()
def playmusic(song,f):
    speak("playing sir")
    pywhatkit.playonyt(song)
def take_command():
    r=sr.Recognizer()
    with sr.Microphone() as source:
        print('listening....')
        r.pause_threshold = 1
        r.energy_threshold = 300
        audio=r.listen(source)
    try :
        print('recognising...')
        query = r.recognize_google(audio,language='en-in')
        print(f"user said {query}\n")
    except BaseException as e:
        #   ('say that again sir ....')
        return "none"
    return query
   


if __name__ == "__main__": 
    hour =0
    if time >='06' and time <= '12':
        speak('good morning sir ..... welcome back to work')
    elif time >'12' and time < '18':
        speak('good afternoon sir ..... welcome back to work')
    elif time >='17' and time < '24':
        speak('good evening sir ..... welcome back to work')
    elif time >='00' and time <='06':
        speak('good evening sir ..... welcome back to work')

    while True:
        query=take_command().lower()
        if "thank you " in query :
            speak('welcome sir .... i am jarvis ... how can i help you')
        elif " leave " in query :
            speak('bye sir ... have a nice day')
            quit()
        elif "play" in query :
            q=list(query.split('play'))
            song = q[1]
            f=0
            playmusic(song,f)
        elif "pause" in query:
            press_and_release(' ')
        elif "resume" in query:
            press_and_release(' ')
        elif "close " in query:
            press_and_release('ctrl+w')
        elif "i am proud of you" in query:
            speak("thank you sir , its an honour")

        elif "time" in query:
            if time>'12':
                hour= int(time)-12
                str(time)
                str(hour)
                speak(f"the time is {hour} {time_mow} ") 
            else :
                speak(f"the time is {time} {time_mow} sir")
     
        elif 'weather' in query:
            weatherForcast(query) 
        elif "volume up" in query:
            strings=query.split("by")
            for i in range(int(strings[1])):
                pyautogui.press("volumeup")
        elif "volume down" in query:
            strings=query.split("by")
            for i in range(int(strings[1])):
                pyautogui.press("volumedown")
        elif "mute" in query:
            pyautogui.press("volumemute")
        elif "unmute" in query:
            pyautogui.press("volumemute")
        elif "charge" in query:
            battery = psutil.sensors_battery()
            persentage=battery.percent
            speak(f"our system have {persentage} percent battery ")
        elif "search" in query:
            searchEngine(query)
        elif "open " in query:
            about=query.split(" ")
            webbrowser.open(f"https://www.{about[1]}.com/")
        elif "send email" in query:
           # if "joida " in query:
             #   email_name = 'joida'
              #  print(type(email_name))
              #  print(email_name)
            df=pd.read_excel("email.xlsx")
            subject = "hi ! this jarvis shatadru's assistant "
            speak("what to sent")
            content =  take_command()
            for index, item in df.iterrows():
                ind=item['name']
                if ind in query:  
                    to=item['email']
                    sendEmail(to,content)
                    speak("email sent sir")
 