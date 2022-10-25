import subprocess
import wolframalpha
import wikiquote
import pyttsx3
import json
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import winshell
import pyjokes
import feedparser
import smtplib
import ctypes
import time
import requests
import fileinput
import getpass
import wmi
import os
from pathlib import Path
from clint.textui import progress
from selenium import webdriver
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen


engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)

DIRECTORIES = {
    "HTML": [".html5", ".html", ".htm", ".xhtml"],
    "IMAGES": [".jpeg", ".jpg", ".tiff", ".gif", ".bmp", ".png", ".bpg", "svg",
               ".heif", ".psd"],
    "VIDEOS": [".avi", ".flv", ".wmv", ".mov", ".mp4", ".webm", ".vob", ".mng",
               ".qt", ".mpg", ".mpeg", ".3gp", ".mkv"],
    "DOCUMENTS": [".oxps", ".epub", ".pages", ".docx", ".doc", ".fdf", ".ods",
                  ".odt", ".pwi", ".xsn", ".xps", ".dotx", ".docm", ".dox",
                  ".rvg", ".rtf", ".rtfd", ".wpd", ".xls", ".xlsx", ".ppt",
                  "pptx"],
    "ARCHIVES": [".a", ".ar", ".cpio", ".iso", ".tar", ".gz", ".rz", ".7z",
                 ".dmg", ".rar", ".xar", ".zip"],
    "AUDIO": [".aac", ".aa", ".aac", ".dvf", ".m4a", ".m4b", ".m4p", ".mp3",
              ".msv", "ogg", "oga", ".raw", ".vox", ".wav", ".wma"],
    "PLAINTEXT": [".txt", ".in", ".out"],
    "PDF": [".pdf"],
    "PYTHON": [".py",".pyi"],
    "XML": [".xml"],
    "EXE": [".exe"],
    "SHELL": [".sh"]
}
FILE_FORMATS = {file_format: directory
                for directory, file_formats in DIRECTORIES.items()
                for file_format in file_formats}


def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def countdown(n) :
    while n > 0:
        print (n)
        n = n - 1
    if n ==0:
        print('BLAST OFF!')

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning Sir!")

    elif hour>=12 and hour<18:
        speak("Good Afternoon Sir!")   

    else:
        speak("Good Evening Sir!")  

    assname=("Jarvis 1 point o")
    speak("I am your Assistant")
    speak(assname)

def usrname():
    speak("What should i call you sir")
    uname=takeCommandname()
    speak("Welcome Mister")
    speak(uname)
    print("#####################")
    print("Welcome Mr.",uname)
    print("#####################")

def quotaton():
    speak(wikiquote.quote_of_the_day())
    print(wikiquote.quote_of_the_day())

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
        print("Unable to Recognizing your voice.")  
        return "None"
    return query

def takeCommandname():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Username...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Trying to Recognizing Name...")    
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        print(e)    
        print("Unable to Recognizing your name.")
        takeCommandname()  
        return "None"
    return query

def takeCommandmessage():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Enter Your Message")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        query = r.recognize_google(audio, language='en-in')
        print(f'Message to be sent is : {query}\n')

    except Exception as e:
        print (e)
        print("Unable to recognize your message")
        print("Check your Internet Connectivity")
    return query

def takeCommanduser():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Name of User or Group")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        query = r.recognize_google(audio, language='en-in')
        print(f'Client to whome message is to be sent is : {query}\n')

    except Exception as e:
        print (e)
        print("Unable to recognize Client name")
        speak("Unable to recognize Client Name")
        print("Check your Internet Connectivity")
    return query

def takeCommandcontent():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("What Should i say, sir")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        query = r.recognize_google(audio, language='en-in')
        print(f'Message to be sent is: {query}\n')

    except Exception as e:
        print (e)
        print("Unable to recognize")
    return query

def organize():
    for entry in os.scandir():
        if entry.is_dir():
            continue
        file_path = Path(entry.name)
        file_format = file_path.suffix.lower()
        if file_format in FILE_FORMATS:
            directory_path = Path(FILE_FORMATS[file_format])
            directory_path.mkdir(exist_ok=True)
            file_path.rename(directory_path.joinpath(file_path))
    try:
        os.mkdir("OTHER")
    except:
        pass
    for dir in os.scandir():
        try:
            if dir.is_dir():
                os.rmdir(dir)
            else:
                os.rename(os.getcwd() + '/' + str(Path(dir)), os.getcwd() + '/OTHER/' + str(Path(dir)))
        except:
            pass

def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('youremail@email.com', 'your-password')
    server.sendmail('youremail@email.com', to, content)
    server.close()



if __name__ == '__main__':
    clear = lambda: os.system('cls')
    clear()
    wishMe()
    usrname()
    speak("Can i tell you a quote of day")
    useropt=takeCommand().lower()
    if 'yes' in useropt or 'sure' in useropt:
        quotaton()
    else:
        speak("Taking you to command function")

    speak("How can i Help you, Sir")
    while True:
        query = takeCommand().lower()
        assname=("Jarvis 1 point o")
        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=3)
            speak("Answer From Wikipedia")
            print(results)
            speak(results)

        elif "Good Morning" in query:
            speak("A warm" +query)
            speak("How are you Mister")
            speak(assname)

        elif "wikipedia" in query and "hindi" in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            query = query.replace("hindi", "")
            results = wikipedia.summary(query, sentences=3)
            speak("According to Wikipedia")
            r = sr.Recognizer()
            results = r.recognize_google(results, language='hi')
            print(results)
            speak(results)

        elif 'open youtube' in query:
            speak("Taking You To Youtube\n")
            webbrowser.open("youtube.com")

        elif 'open google' in query:
            speak("Taking you to Google\n")
            webbrowser.open("google.com")

        elif "change brightness to " in query:
            query=query.replace("change brightness to","")
            brightness = query 
            c = wmi.WMI(namespace='wmi')
            methods = c.WmiMonitorBrightnessMethods()[0]
            methods.WmiSetBrightness(brightness, 0)

        elif "Organize Files" in query:
            organize()

        elif 'open stackoverflow' in query:
            speak("Here you go to Stack Over flow.Happy coding")
            webbrowser.open("stackoverflow.com")   

        elif "stackoverflow " in query:
            speak("Stackoverflow khola ja rha h")
            webbrowser.open("stackoverflow.com")   

        elif "send a whatsaap message" in query or "send a WhatsApp message" in query:
            driver = webdriver.Chrome('Web Driver Location')
            driver.get('https://web.whatsapp.com/')
            speak("Scan QR code before proceding")
            tim=10
            time.sleep(tim)
            speak("Enter Name of Group or User")
            name = takeCommanduser()
            speak("Enter Your Message")
            msg = takeCommandmessage()
            count = 1
            user = driver.find_element_by_xpath('//span[@title = "{}"]'.format(name))
            user.click()
            msg_box = driver.find_element_by_class_name('_3u328')
            for i in range(count):
                msg_box.send_keys(msg)
                button = driver.find_element_by_class_name('_3M-N-')
                button.click()
                
        elif "stackoverflow " in query:
            speak("Stackoverflow khola ja rha h")
            webbrowser.open("stackoverflow.com")   

        elif 'play music' in query or "play song" in query or "gaana"in query or "song" in query:
            #music_dir = "G:\\Song"
            username = getpass.getuser()
            music_dir = "C:\\Users\\"+username+"\\Music"
            songs = os.listdir(music_dir)
            print(songs)    
            random=os.startfile(os.path.join(music_dir, songs[1]))

        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")    
            speak(f"Sir, the time is {strTime}")

        elif "samay" in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")    
            speak(f"samaye hai {strTime}")

        elif 'open opera' in query:
            codePath = r"C:\\Users\\GAURAV\\AppData\\Local\\Programs\\Opera\\launcher.exe"
            os.startfile(codePath)

        elif 'email to gaurav' in query:
            try:
                content = takeCommandcontent()
                to = "gauravkumarjha27@gmail.com"    
                sendEmail(to, content)
                speak("Email has been sent!")
            except Exception as e:
                print(e)
                speak("I am not able to send this email")

        elif 'how are you' in query:
        	speak("I am fine , Thank you")
        	speak("How are you, Sir")

        elif "change my name to" in query:
            query=query.replace("change my name to","")
            assname=query

        elif "change name" in query:
            speak("What would you like to call me ,Sir ")
            assname = takeCommand()
            speak("Thanks for naming me")

        elif "what's your name" in query or "What is your name" in query:
            speak("My friends call me")
            speak(assname)
            print("My friends call me",assname)

        elif 'exit' in query:
            speak("Thanks for giving me your time")
            exit()

        elif "who made you" in query or "who created you" in query: 
            speak("I have been created by Gaurav.")

        elif 'joke' in query:
            speak(pyjokes.get_joke())

        elif "calculate" in query:
            app_id = "Wolframe Alpha API"
            client = wolframalpha.Client(app_id)
            indx = query.lower().split().index('calculate')
            query = query.split()[indx + 1:]
            res = client.query(' '.join(query))
            answer = next(res.results).text
            print("The answer is " + answer)
            speak("The answer is " + answer)

        elif 'search' in query or 'play' in query: 
            query = query.replace("search", "") 
            query = query.replace("play", "")
            webbrowser.open(query) 

        elif "who i am" in query:
            speak("If you talk then definately your human.")

        elif "why you came to world" in query:
            speak("Thanks to Gaurav. further It's a secret")

        elif 'is love' in query:
            speak("It is 7th sense that destroy all other senses")

        elif "who are you" in query:
            speak("I am your virtual assistant created by Gaurav")

        elif 'reason for you' in query:
            speak("I was created as a Minor project by Mister Gaurav ")

        elif 'change background' in query:
            ctypes.windll.user32.SystemParametersInfoW(20, 0, "C:\\Users\\GAURAV\\OneDrive\\Minor Project\\Voice\\back.jpg" , 0)
            speak("Background changed succesfully")

        elif 'open bluestack' in query:
            appli= r"C:\\ProgramData\\BlueStacks\\Client\\Bluestacks.exe"
            os.startfile(appli)

        elif 'google news' in query:
            try:
                jsonObj = urlopen('''https://newsapi.org/v2/top-headlines?sources=google-news-in&apiKey=Google news API key''')
                data = json.load(jsonObj)
                i = 1
                speak('')
                print('''===============Google News============'''+ '\n')
                for item in data['articles']:
                    print(str(i) + '. ' + item['title'] + '\n')
                    print(item['description'] + '\n')
                    speak(str(i) + '. ' + item['title'] + '\n')
                    i += 1
            except Exception as e:
                    print(str(e))

        elif "bbc news" in query:
            try:
                main_url = " https://newsapi.org/v1/articles?source=bbc-news&sortBy=top&apiKey=BBC News API key"
                open_bbc_page = requests.get(main_url).json() 
                article = open_bbc_page["articles"] 
                results = [] 
                for ar in article: 
                    results.append(ar["title"]) 
                for i in range(len(results)): 
                    print(i + 1, results[i])
            except Exception as e:
                print(str(e))

        elif 'news' in query: #samachar
            try:
                jsonObj = urlopen('''https://newsapi.org/v1/articles?source=the-times-of-india&sortBy=top&apiKey=Time of INDIA API key''')
                data = json.load(jsonObj)
                i = 1
                speak('here are some top news from the times of india')
                print('''===============TIMES OF INDIA============'''+ '\n')
                for item in data['articles']:
                    print(str(i) + '. ' + item['title'] + '\n')
                    print(item['description'] + '\n')
                    speak(str(i) + '. ' + item['title'] + '\n')
                    i += 1
            except Exception as e:
                print(str(e))

        elif 'lock window' in query or "system ko lock Karen" in query:
            speak("locking the device")
            ctypes.windll.user32.LockWorkStation()

        elif 'shutdown system' in query:
            speak("Hold On a Sec! Your system is on its way to shut down")
            subprocess.call('shutdown /p /f')
            
        elif 'empty recycle bin' in query:
            winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
            speak("Recycle Bin Recycled")

        elif "don't listen" in query or "stop listening" in query:
            speak("for how much time you want to stop jarvis from listening commands")
            a=int(takeCommand())
            time.sleep(a)
            print(a)

        elif "where is" in query:
            query=query.replace("where is","")
            location = query
            speak("Locating ")
            speak(location)
            webbrowser.open("https://www.google.nl/maps/place/" + location + "")

        elif "camera" in query or "take a photo" in query:
            ec.capture(0,"Jarvis Camera ","img.jpg")

        elif "restart" in query:
            subprocess.call(["shutdown", "/r"])

        elif "hibernate" in query or "sleep" in query:
            speak("Hibernating")
            subprocess.call("shutdown /i /h")

        elif "log off" in query or "sign out" in query:
            speak("Make sure all the application are closed before sign-out")
            time.sleep(5)
            subprocess.call(["shutdown", "/l"])

        elif "countdown of" in query:
            query = int(query.replace("countdown of ",""))
            countdown(query)

        elif "write a note" in query:
            speak("What should i write , sir")
            note= takeCommand()
            file = open('jarvis.txt','w')
            speak("Sir, Should i include date and time")
            snfm = takeCommand()
            if 'yes' in snfm or 'sure' in snfm:
                strTime = datetime.datetime.now().strftime("%H:%M:%S")
                file.write(strTime)
                file.write(" :- ")
                file.write(note)
            else:
                file.write(note)
        
        elif "show note" in query:
            speak("Showing Notes")
            file = open("jarvis.txt", "r") 
            print(file.read())
            speak(file.read(6))

        elif "update assistant" in query:
            speak("After downloading file please replace this file with the downloaded one")
            url = '#url after uploading file'
            r = requests.get(url, stream=True)
            with open("Voice.py", "wb") as Pypdf:
                total_length = int(r.headers.get('content-length'))
                for ch in progress.bar(r.iter_content(chunk_size = 2391975), expected_size=(total_length/1024) + 1):
                    if ch:
                      Pypdf.write(ch)

        elif "jarvis" in query:
            wishMe()
            speak("Jarvis 1 point o in your service Mister")
            speak(assname)

        elif "weather" in query:
            api_key = "Open weather map API key"

                base_url = "http://api.openweathermap.org/data/2.5/weather?"
            speak(" City name ")
            print("City name : ")
            city_name=takeCommand()
            complete_url = base_url + "appid=" + api_key + "&q=" + city_name
            response = requests.get(complete_url) 
            x = response.json() 
            if x["cod"] != "404": 
                y = x["main"] 
                current_temperature = y["temp"] 
                current_pressure = y["pressure"] 
                current_humidiy = y["humidity"] 
                z = x["weather"] 
                weather_description = z[0]["description"] 
                print(" Temperature (in kelvin unit) = " +str(current_temperature)+"\n atmospheric pressure (in hPa unit) ="+str(current_pressure) +"\n humidity (in percentage) = " +str(current_humidiy) +"\n description = " +str(weather_description)) 
            else: 
                speak(" City Not Found ") 

                elif "wikipedia" in query:
            webbrowser.open("wikipedia.com")

        elif "will you be my gf" in query or "will you be my bf" in query:   #most asked question from google Assistant
            speak("I'm not sure about , may be you should give me some time")

        elif "how are you" in query:
            speak("I'm fine, glad you asked me that")

        elif "i love you" in query:
            speak("It's hard to understand")

        elif "what is" in query or "who is" in query:
            client= wolframalpha.Client("WOlframe Alpha API key")
            res = client.query(query)
            try:
                print(next(res.results).text)
                speak(next(res.results).text)
            except StopIteration:
                print ("No results")

        elif "open Gmail" in query:
            webbrowser.open("https://mail.google.com/mail/u/0/#inbox")

        elif "open yahoo mail" in query:
            webbrowser.open("https://in.mail.yahoo.com")

        elif "Show project Report" in query:
            speak("Opening Minor Project Report")
            projectre= r"C:\\Users\\GAURAV\\Desktop\\Minor Project\\Presentation\\Project Report.docx"
            os.startfile(projectre)
