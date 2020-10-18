'''
Apollo Voice assistant

Rohit Das - https://rohit-das.weebly.com/
10/18/2020
v1.2.2
'''

'''
        Things to try:
        1. "Tell me about ___": Opens up a wikipedia link on the topic
        2. "Search ___": Opens up a google query
        3. "Open ___": Opens up a website or a favorite website
        4. "Launch ___": Launches an application 
        5. "Pokemon: " Can enter battle mode, find weaknesses, open 100iv chart, and some other easter eggs.
        7. "Set brightness to": Sets brightness to percentage
        8. "Tell me a joke": Finds a random dad joke
        9. "News for today": Reads out highlights from today's news
        10. "What's the weather in ___" : Reads out weather 
        11. "What time is it" : Tells the time
        12. "Remind me about ___": Sends a text to phone at reminder time about reminder
        13. "Text/Send ___": Sends a text to a preconfigured contact
        14. "Bubbles": blows bubbles on the screen
        15. "Mystify": Mystifies the screen
        ...etc

Look for #REDACTED comments to see where to edit code. Also, please note that this was designed to work on Windows OS 
so some changes will be needed for MAC and Linux systems.
'''

import speech_recognition as sr
import os
import sys
import re
import webbrowser
import smtplib
import requests
import subprocess
from pyowm import OWM
import youtube_dl
from gtts import gTTS
#import vlc
import urllib
import json
from bs4 import BeautifulSoup as soup
from urllib.request import urlopen
import wikipedia
import random
from playsound import playsound
from time import strftime
import datetime
from win32com.client import Dispatch
import pypokedex
import wmi
from word2number import w2n

def apolloResponse(audio):
    "speaks audio passed as argument"
    print(audio)
    for line in audio.splitlines():
        speak = Dispatch("SAPI.SpVoice")
        speak.rate = 2.5
        speak.Speak(line)
        #tts = gTTS(text=line, lang='en', slow=False)
        #fileName = "response" + datetime.datetime.now().strftime("%d%m%Y%H%M%S") + ".mp3"
        #tts.save(fileName)
        #playsound(fileName)
        #print('')
        #os.remove(fileName)

def myCommand():
    "listens for commands"
    r = sr.Recognizer()
    with sr.Microphone() as source:
        #r.pause_threshold = 1
        r.adjust_for_ambient_noise(source, duration=0.3)
        print("Say something...")
        audio = r.listen(source)
    try:
        command = r.recognize_google(audio).lower()
        print('You said: ' + command + '\n')
    #loop back to continue to listen for commands if unrecognizable speech is received
    except sr.UnknownValueError:
        print('....')
        command = myCommand()
    return command

def text(receiver, content):
    mail = smtplib.SMTP('smtp-mail.outlook.com', 587)
    mail.ehlo()
    mail.starttls()
    mail.ehlo()

    #REDACTED
    mail.login('example@outlook.com', 'example_password')
    mail.sendmail('example@outlook.com', receiver, content)
    mail.close()

#REDACTED
#Fill in apps list with a Dict of form (Name, Location). Location can either be found (.exe file) or simply make a 
#folder of shortcuts (.lnk file)

def initAppsList():
    appsList = []
    appsList.append(makeAppDict('exampleAppName', "C:/Users/Rohit/Example/Location.lnk"))
    appsList.append(makeAppDict('exampleAppName1', "C:/Users/Rohit/Example/Location1.lnk"))
    appsList.append(makeAppDict('exampleAppName2', "C:/Users/Rohit/Example/Location2.lnk"))
    appsList.append(makeAppDict('exampleAppName3', "C:/Users/Rohit/Example/Location3.lnk"))
    appsList.append(makeAppDict('exampleAppName4', "C:/Users/Rohit/Example/Location4.lnk"))
    return appsList
    
def makeAppDict(name, address):
    dict = {
            "name": name,
            "address": address
            }
    return dict

def assistant(command):
    
    try:
        "if statements for executing commands"
            
        #ask me anything
        if 'tell me about' in command:
            reg_ex = re.search('tell me about (.*)', command)
            try:
                if reg_ex:
                    topic = reg_ex.group(1)
                    ny = wikipedia.page(topic)
                    apolloResponse(ny.content[:500].encode('utf-8'))
            except Exception as e:
                    print(e)
                    apolloResponse(e)    
    
        #Google Search
        elif 'search' in command:
            print(command)
            search = command.split()
            print(search)
            found = False
            first = True
            print(search)
            searchString = ""
            for word in search:
                if found and first:
                    searchString += word
                    first = False
                elif found and not first:
                    searchString += '+'
                    searchString += word
                elif word == 'search':
                    found = True
            if searchString == "":
                apolloResponse("Search Failed")
            url = 'https://www.google.com/search?q=' + searchString
            print(url)
            webbrowser.open(url)
            apolloResponse("Your search has been queried")
        
        #open website
        elif 'open' in command:
            reg_ex = re.search('open (.+)', command)
            if reg_ex:
                domain = reg_ex.group(1)
                print(domain)
                if domain == 'calendar' or domain == 'my calendar':
                    url = 'https://calendar.google.com/calendar/r'
                    webbrowser.open(url)
                elif domain == 'Netflix' or domain == 'netflix':
                    url = 'https://www.netflix.com/browse'
                    webbrowser.open(url)
                elif domain == 'Spotify' or domain == 'spotify' or domain == 'music':
                    url = 'https://open.spotify.com/'
                    webbrowser.open(url)
                elif domain == 'mail' or domain == 'email' or domain == 'e-mail':
                    url = 'https://mail.google.com/'
                    webbrowser.open(url)
                else:
                    url = 'https://www.' + domain
                    webbrowser.open(url)
                    apolloResponse('The website you have requested has been opened.')
            else:
                pass
    
        elif 'launch' in command:
            commandArr = command.split()
            for word in commandArr:
                for i in range(len(appsList)):
                    if appsList[i]["name"] == word:
                        launchString = appsList[i]["address"]
                        #launchString += " /start"
                        #print("LaunchString:")
                        #print(launchString)
                        os.startfile(launchString)
                        apolloResponse('Your app has been launched')
                        return
        
        #Pokemon Go cool stuff
        elif "pokemon" in command:
            if "battle mode" in command or "mode" in command:
                webbrowser.open("https://i.ytimg.com/vi/ZN2oc_oYaYs/maxresdefault.jpg")
                apolloResponse("Battle mode activated")
            if "weakness" in command:
                pokemon = command.split()
                found = False
                for word in pokemon:
                    if word == "pokemon":
                        found = True
                    elif found == True:
                        searchString = "search "
                        searchString += word
                        searchString += " weakness"
                        assistant(searchString)
                        apolloResponse("Weakness found")
                        found = False
            if "hundred" in command or "100" in command or "hundo" in command:
                apolloResponse("Opening chart")
                line = command.split()
                found = False
                for word in line:
                    if not (word == 'hundred' or word == '100' or word == 'pokemon' or word == 'hundo') and not found:
                        try:
                            print(word)
                            print(pypokedex.get(name=word))
                            pokemonNum = pypokedex.get(name=word).dex
                            found = True
                        except Exception as e:
                            print(e)
                            apolloResponse("Pokemon could not be found sadly")
                            return
                url = "https://db.pokemongohub.net/pokemon/" + str(pokemonNum)
                webbrowser.open(url)
                
        #Brightness
        elif "brightness" in command:
            commandArr = command.split()
            brightness = 100
            changed = False
            for i in range(len(commandArr), 0, -1):
                if ('%' in commandArr[i - 1]):
                    brightness =  int(commandArr[i - 1][:-1])
                    print(brightness)
                    changed = True
                        
            if changed:
                c = wmi.WMI(namespace='wmi')
                methods = c.WmiMonitorBrightnessMethods()[0]
                methods.WmiSetBrightness(brightness, 0)
                apolloResponse("Done")
            else:
                apolloResponse("Sorry, command could not be interpretted")
                print(commandArr)    
    
        #joke
        elif 'joke' in command:
            res = requests.get(
                    'https://icanhazdadjoke.com/',
                    headers={"Accept":"application/json"})
            if res.status_code == requests.codes.ok:
                apolloResponse(str(res.json()['joke']))
            else:
                apolloResponse('oops!I ran out of jokes')
    
        #top stories from google news
        elif 'news for today' in command:
            try:
                news_url="https://news.google.com/news/rss"
                Client=urlopen(news_url)
                xml_page=Client.read()
                Client.close()
                soup_page=soup(xml_page,"xml")
                news_list=soup_page.findAll("item")
                for news in news_list[:15]:
                    apolloResponse(news.title.text.encode('utf-8'))
            except Exception as e:
                    print(e)
    
        #current weather
        elif 'weather' in command:
            reg_ex = re.search('weather in (.*)', command)
            if reg_ex:
                city = reg_ex.group(1)
                owm = OWM(API_key='ab0d5e80e8dafb2cb81fa9e82431c1fa')
                obs = owm.weather_at_place(city)
                w = obs.get_weather()
                k = w.get_status()
                x = w.get_temperature(unit='fahrenheit')
                apolloResponse('Current weather in %s is %s. The maximum temperature is %0.2f and the minimum temperature is %0.2f degree fahrenheit' % (city, k, x['temp_max'], x['temp_min']))
    
        #time
        elif 'time' in command:
            now = datetime.datetime.now()
            apolloResponse('Current time is %d hours %d minutes' % (now.hour, now.minute))
    
        #send email
        elif 'reminder' in command or 'remind' in command:
            apolloResponse('What should I remind you about?')
            content = "Subject: Reminder\n\n"
            content += myCommand()
            apolloResponse('Got it. Please hold')
            print(content)

            #REDACTED 
            #Fill in number email address (from cell provider) here:
            address = '12345678@cellprovider.com'
            text(address, content)
            apolloResponse("Reminder sent")
        
        elif 'send' in command or 'text' in command:
            address = command
            print(address)
            apolloResponse("What do you want to send?")
            message = "Subject: Apollo\n\n"
            message += myCommand()
            messageSent = False
            #REDACTED
            #Fill in contact names and number email addresses here:
            if 'contact1' in address or 'nickname1' in address:
                text("123456789@cellprovider.com", message)
                messageSent = True
            if 'contact2' in address or 'nickname2' in address:
                text("123456789@cellprovider.com", message)
                messageSent = True
            if 'contact3' in address or 'nickname3' in address:
                text("123456789@cellprovider.com", message)
                messageSent = True
            else:
                apolloResponse("Message failed to send")
        
        elif 'bubbles' in command or 'bubble' in command:
            apolloResponse("Have fun with Bubbles!")
            os.system("C:/Windows/System32/Bubbles.scr /start")
            apolloResponse("Sorry sir, I had to blow out the bubbles")
            
        elif 'mystify' in command:
            apolloResponse("Mystifying")
            os.system("C:/Windows/System32/Mystify.scr /start")
            apolloResponse("Welcome Back")
        
        #shutdown Apollo
        elif 'shutdown' in command or 'shut down' in command or 'bye' in command or 'goodbye' in command or 'stop' in command:
            apolloResponse('Bye bye Sir. Have a nice day')
            os._exit(1)
        
        #greetings
        elif 'hello' in command or 'hi' in command or 'good morning' in command or 'good afternoon' in command or 'good evening' in command:
            day_time = int(strftime('%H'))
            if day_time < 12:
                apolloResponse('Hello Sir. Good morning')
            elif 12 <= day_time < 18:
                apolloResponse('Hello Sir. Good afternoon')
            else:
                apolloResponse('Hello Sir. Good evening')
            
        elif 'are you listening' in command:
            apolloResponse("Yes I am. Sorry to intrude.")
        
        elif 'pause' in command:
            apolloResponse("Pausing")
            reactivate = myCommand()
            while not reactivate == "apollo":
                reactivate = myCommand()
            apolloResponse("Hey, I'm back")
    except:
        apolloResponse("Sorry sir, general error occured. Please try again")

appsList = initAppsList()
print("")
print("     AA       PPPPPPP     OOOO    LL         LL          OOOO"  )
print("    AAAA      PP    PP  OO    OO  LL         LL        OO    OO")
print("   AA  AA     PPPPPPP   OO    OO  LL         LL        OO    OO")
print("  AAAAAAAA    PP        OO    OO  LL         LL        OO    OO")
print(" AA      AA   PP        OO    OO  LL         LL        OO    OO")
print("AA        AA  PP          OOOO    LLLLLLLLL  LLLLLLLL    OOOO"  )
print()
print("Version 1.0 ~ Rohit Das")

assistant("hello")
apolloResponse("Apollo here.     ")
looper = True

#loop to continue executing multiple commands
while looper:
    assistant(myCommand())

"Qed"
