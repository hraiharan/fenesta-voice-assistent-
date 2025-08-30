import speech_recognition as sr
import webbrowser
import csv
import requests
import time
import wikipedia
from time import localtime, strftime
import re
import smtplib
from datetime import datetime
import winshell
import os.path
import ctypes
from tkinter.filedialog import askopenfile
import pygame
from pygame import mixer
from tkinter import *
import os
import pyttsx3
from tkinter.filedialog import askdirectory

import pyautogui



# =====================================================================================
settings = '''
1=change the format typing to voice 
2=help me
3=exit
'''
main_dict = {'hello': "hi sir what can i do for you",
             "hi":"hi sir what can i do for you",
             'your name': "I am your voice assistant,fenesta",
             'thank': "You are welcome",
             'fenesta': 'Yes what should i do?',
             'how are you': "I am fine, Thank you",
             'fine': "It's good to know that your fine",
             "good": "It's good to know that your fine",
             "who i am": "If you talk then definitely your human.",
             "why you came to world": "Thanks to hari. further It's a secret",
             'is love': "It is 7th sense that destroy all other senses",
             "who are you": "I am your virtual assistant created by hari",
             'reason for you': "I was created as a Minor project by Mister hariharan",
             "i love you": "It's hard to understand"
             }
genral_question = ['can you say', 'what is', 'who is']
my_computer = ["disk c", "open local disk c", "local disk c",
               "launch local disk d", "open local disk d", "local disk d",
               "local disk g", "open local disk g", "launch local disk g",
               "open new volume", "launch new volume", "new volume"]

searching_in_chrome = {"youtube": "https://www.youtube.com/results?search_query=",
                       "google": "https://www.google.com/search?q=",
                       'wikipedia': "https://en.wikipedia.org/wiki/"}

Chrome = "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s"
# =============================================================================================================

def speak(rand):
    engine = pyttsx3.init()
    engine.setProperty("rate", 110)  # 110= the speed of the voice
    engine.say(rand)
     
    engine.runAndWait()

# def functions
def greet():
    time_now = datetime.now()
    dt_string = int(time_now.strftime(" %H  "))
    if dt_string < 12:  # It will check the time is less than 12
        speak("A magical morning, this is fenesta as your voice assistant")
    elif (dt_string >= 12) and (dt_string < 16):  # It will check the time is more or equal to 12 and less than 12
        speak("A wonderful afternoon,this is fenesta as your voice assistant")
    else:
        speak("A happy evening,this is fenesta as your voice assistant")

greet()
 
# ===============================(opening folder)=============================================
def open_folder(message):
    import os, re, shutil

    def giveback(path):
        return os.path.dirname(os.path.normpath(path))

    if 'local disk c' in message:
        c = "c:/"
    elif 'local disk d' in message or "new volume" in message:
        c = "d:/"
    else:
        print("No valid drive mentioned.")
        return

    prev_c = None
    if c != prev_c:
        os.startfile(c)
        prev_c = c

    while True:
        try:
            i = takeCommand().lower()

            if 'back' in i:
                if c.lower() in ["c:/", "d:/"]:
                    message = takeCommand()
                    if 'local disk c' in message:
                        c = "c:/"
                    elif 'local disk d' in message:
                        c = "d:/"
                else:
                    c = giveback(c)

            elif 'open' in i or "launch" in i or 'play' in i:
                target = i.replace('play', '').replace('open', '').replace('launch', '').replace('the', '').strip().lower()
                found = False
                for File in os.listdir(c):
                    clean_name = re.sub('[^a-zA-Z0-9. \n\.]', '', File).lower()
                    if target in clean_name:
                        path = os.path.join(c, File)
                        if os.path.isdir(path):
                            c = path  # move into folder
                        elif os.path.isfile(path):
                            os.startfile(path)  # open the file
                        found = True
                        break
                if not found:
                    print("No matching file/folder found.")

            elif "create" in i:
                if "folder" in i:
                    print("Name of folder?")
                    name = takeCommand()
                    folder_path = os.path.join(c, name)
                    os.mkdir(folder_path)
                    print("Folder created.")
                else:
                    print("File creation not yet implemented.")

            elif 'delete' in i or "remove" in i:
                target = i.replace('delete', '').replace('remove', '').strip()
                for File in os.listdir(c):
                    if target in File.lower():
                        path = os.path.join(c, File)
                        if os.path.isdir(path):
                            shutil.rmtree(path)
                        else:
                            os.remove(path)
                        print("Deleted.")
                        break

            elif 'quit' in i or 'close' in i or 'bye' in i:
                break

            # Only reopen if path changed
            if c != prev_c:
                os.startfile(c)
                prev_c = c

        except Exception as f:
            c = giveback(c)
            prev_c = c
            print("Error:", f)




# =========================(obtain audio)==============================================================================================
def takeCommand():
    # It takes microphone input from the user and returns string output
    r = sr.Recognizer()

    with sr.Microphone() as source:  # It takes microphone as the source of input
        print("listening...")
        r.pause_threshold = 1  # after during execution it) waits for 1 seconds

        audio = r.listen(source)  # It listen the audio as input or source and it omits noise
        print("Recognizing")
    try:
        command = (r.recognize_google(audio,language='en-in'))  # By using recognize_google it convert the audio to text in indian language
        print(f"User said: {command}\n")  # printing the command

    except Exception as e:
        print(e)
        return takeCommand()
    return command.lower()
# ==================================(geting path)===================================================================================================
 
# ==================================(take command)===================================================================================================
def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    your_email_id = input("your mail id::-")
    your_email_password = input('your mail password::-')
    server.login(your_email_id, your_email_password)
    server.sendmail('your email id', to, content)
    server.close()

def is_online():
    try:
        requests.get('https://www.google.com', timeout=3)
        return True
    except requests.ConnectionError:
        return False
# ================================================================================================================================
while True:
    if __name__ == "__main__":
        clear = lambda: os.system('cls')  # waste
        clear()
        if __name__ == "__main__":
             
            if is_online():
               message= takeCommand()
            else:
                message = input("enter:") 
            message = message.lower()
            print(message)
            for zn in main_dict:
                if zn in message:
                    speak(main_dict[zn])
            # ==========================================================================================================================================
            if 'open google search' in message or 'open youtube search' in message:
                messag = message.replace('open google search ', '').replace('open youtube search ', '')
                if "google" in message: 
                    webbrowser.register('brave', None, webbrowser.GenericBrowser(r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe"))
                    browser = webbrowser.get('brave')
                    browser.open_new_tab(searching_in_chrome['google'] + messag)
                if "youtube" in message:
                    webbrowser.register('brave', None, webbrowser.GenericBrowser(r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe"))
                    browser = webbrowser.get('brave')
                    browser.open_new_tab(searching_in_chrome['youtube'] + messag)
                speak('Opening' + messag)
            if '.com' in message:
                speak('Opening' + message)
                webbrowser.register('brave', None, webbrowser.GenericBrowser(r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe"))
                browser = webbrowser.get('brave')
                browser.open_new_tab( message)
            if "show desktop" in message or "minimize"in message:
                pyautogui.hotkey('win', 'd') 

            if 'bye' in message:
                speak('bye fenesta powering off in 3, 2, 1, 0')
                exit()
                
            if "wallpaper" in message:
                try:
                    path_first1 = str(askopenfile()).split(" ")
                    path_of_wallpaper = path_first1[1].split("=")
                    path_of_wallpaper = path_of_wallpaper[1].replace("/", "//").replace("'", "")
                    ctypes.windll.user32.SystemParametersInfoW(20, 0, path_of_wallpaper, 0)
                    speak("wallpaper is been changed")
                except Exception as j:
                    print("you have not inserted a image")

            if message in my_computer:
                open_folder(message)


            
            

            if 'setting' in message:
                print(settings)
                change_settings = int(input('Enter the number to change the setting'))
                if change_settings == 1:
                    voice_type = input('𝖊𝖓𝖙𝖊𝖗 (VI) 𝖋𝖔𝖗 𝖛𝖔𝖎𝖈𝖊 𝖆𝖘𝖘𝖎𝖘𝖙𝖊𝖓𝖙 𝖔𝖗 (TP) 𝖋𝖔𝖗 𝖙𝖞𝖕𝖎𝖓𝖌:- ').upper()
                    to_check_net = is_online()
                    if voice_type.lower() == 'vi':
                        to_check_net = is_online()
                        if to_check_net:
                            message= takeCommand()
                        else:
                            print('You are not connected to the internet')
                            message = input("enter:")
                         
                    else:
                        to_check_net = False

                elif change_settings == 2:
                    print(help)
                elif change_settings == 3:
                    exit()
                else:
                    print('invalid character')
            
            if 'wikipedia' in message:
                try:
                    speak('searching in wikipedia....')
                    message = message.replace('wikipedia', '')
                    result_wiki = wikipedia.summary(message, sentences=2)
                    speak(result_wiki)
                    print(result_wiki)
                    time.sleep(18)
                except Exception:
                    print('can u say what else')
            if "write a note" in message:
                speak("What should i write, sir")
                note = takeCommand()
                file = open('jarvis.txt', 'w')
                speak("Sir, Should i include date and time")
                snfm = takeCommand()
                if 'yes' in snfm or 'sure' in snfm:
                    strTime = datetime.datetime.now().strftime("% H:% M:% S")
                    file.write(strTime)
                    file.write(" :- ")
                    file.write(note)
                else:
                    file.write(note)
            if "show note" in message:
                speak("Showing Notes")
                file = open("jarvis.txt", "r")
                print(file.read())
                speak(file.read(6))
                time.sleep(10)
                
            if message in genral_question:
                try:
                    result2 = wikipedia.summary(message, sentences=1)
                    speak(result2)
                    print(result2)
                    time.sleep(18)
                except Exception:
                    print('could not find what have you said,can you say something else')
            if 'turn off' in message or 'shutdown' in message:
                speak('see you later')
                os.system("shutdown /p")
            if 'sleep mode' in message:
                speak('good night')
                os.system('rundll32.exe powrprof.dll,SetSuspendState 0,1,0')
            if "don't listen" in message or "stop listening" in message:
                speak("for how much time you want to stop fenesta from listening commands")
                a = int(takeCommand())
                time.sleep(a)
                print(a)
            if 'restart' in message:
                speak('see you in five minutes')
                os.system("shutdown /r /t 1")
            if 'what is the time' in message:
                tim = strftime("%X", localtime())
                speak(tim)
            if 'empty recycle bin' in message:
                winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
                speak("Recycle Bin Recycled")
            time.sleep(5)