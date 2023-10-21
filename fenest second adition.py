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

from tkinter import *
import os
import pyttsx3
from tkinter.filedialog import askdirectory


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
# ====================(to goback)============================================================
def giveback(path):
    k = path.split("//")
    print(k)
    k.pop()
    k.pop()
    str1 = ""
    for ele in k:
        str1 = str1 + ele + '//'
    return str1

# ===============================(opening folder)=============================================
def open_folder(message):
    if __name__ == "__main__":
        if 'local disk c' in message:
            c = "c://"
        elif 'local disk g' in message or "new volume" in message:
            c = "g://"
        elif 'local disk d' in message:
            c = "d://"
        os.startfile(c)  # It will open  the folder
        while True:
            try:
                i = takeCommand().lower()
                if 'back' in i:
                    if c == "c://" or c == "d://" or c == "g://":
                        message = takeCommand()
                        if 'local disk c' in message:
                            c = "c://"
                        elif 'local disk g' in message or "new volume" in message:
                            c = "g://"
                        elif 'local disk d' in message:
                            c = "d://"
                        os.startfile(c)
                    else:
                        c = giveback(c)
                        os.startfile(c)
                elif "search" in i:
                    for z, x, v in os.walk(r"D:\user\Desktop\hari"):
                        if "folder" in i:
                            i = i.replace("search ", "").replace("folder ", "")
                            print(i)
                        if i in z:
                            os.startfile(z)

                elif 'open' in i or "launch" in i:
                    i = i.replace('play', '').replace('open', '').replace(' ', '').replace('the', '')
                    for File in os.listdir(c):
                        l = re.sub('[^a-zA-Z0-9. \n\.]', ' ', File).lower()
                        if i in l:
                            c = c + File + "//"
                            os.startfile(c)
                elif "create" in i:
                    if "folder" in i:
                        print("please say the name for the folder")
                        i = takeCommand()
                        c = c + i + "//"
                        print("folder has been created")
                        os.mkdir(c, 0o666)
                        c = giveback(c)
                    elif "msword" in i or " Word Document" in i:
                        print("please give the name of the file")
                        i = takeCommand()
                        c = c + i + ".docx"
                        ha = open(c, "w")
                        ha.close()
                        os.startfile(c)
                        c = giveback(c)
                    elif "python" in i:
                        print("please give the name of the file")
                        i = takeCommand()
                        c = c + i + ".py"
                        ha = open(c, "w")
                        ha.close()
                        os.startfile(c)
                        c = giveback(c)
                    elif "notepad" in i or "text file" in i:
                        print("please give the name of the file")
                        i = takeCommand()
                        c = c + i + ".txt"
                        ha = open(c, "w")
                        ha.close()
                        os.startfile(c)
                        c = giveback(c)
                    elif "csv file" in i:
                        print("please give the name of the file")
                        i = takeCommand()
                        c = c + i + ".csv"
                        ha = open(c, "w")
                        ha.close()
                        os.startfile(c)
                        c = giveback(c)
                    elif "excel" in i:
                        print("please give the name of the file")
                        i = takeCommand()
                        c = c + i + ".xlsx"
                        ha = open(c, "w")
                        ha.close()
                        os.startfile(c)
                        c = giveback(c)
                    elif "power point presentation" in i or "ppt" in i:
                        print("please give the name of the file")
                        i = takeCommand()
                        c = c + i + ".pptx"
                        ha = open(c, "w")
                        ha.close()
                        os.startfile(c)
                        c = giveback(c)
                    else:
                        print("no other file is found")
                elif 'delete' in i or "remove" in i:
                    i = i.replace('delete', '').replace('remove', '').replace(' ', '').replace('the', '')
                    for File in os.listdir(c):
                        l = File.lower()
                        if i in l:
                            c = c + File + "//"
                            os.rmdir(c + "//" + File + "//")
                            c = giveback(c)

                elif "quit" in i or "close" in i or "bye" in i:
                    break
                else:
                    os.startfile(c)
            except Exception as f:
                c = giveback(c)
                print(f)

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
def getting_path_for_music(c,k, h=[]):
    for file in os.listdir(c):  # This function read all the files present in the given folder

        if file.endswith(k):  # It check the file end with mp3 ==> means songs
            file = c + '/' + file  # creating a list
            h.append(file)  # appending the mp3 files in the folder
    return h  # returning the list to the program

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

# ================================================================================================================================
while True:
    if __name__ == "__main__":
        clear = lambda: os.system('cls')  # waste
        clear()
        if __name__ == "__main__":
            message = input("enter:")
            for zn in main_dict:
                if zn in message:
                    speak(main_dict[zn])
            # ==========================================================================================================================================
            if 'open google search' in message or 'open youtube search' in message:
                messag = message.replace('open google search ', '').replace('open youtube search ', '')
                if "google" in message:
                    webbrowser.get(Chrome).open(searching_in_chrome['google'] + messag)
                if "youtube" in message:
                    webbrowser.get(Chrome).open(searching_in_chrome['youtube'] + messag)
                speak('Opening' + messag)
            if '.com' in message:
                speak('Opening' + message)
                webbrowser.get(Chrome).open('http://www.' + message)
            if message in my_computer:
                open_folder(message)
            if 'bye' in message:
                speak('bye fenesta powering off in 3, 2, 1, 0')
                exit()
            if "wallpaper" in message:
                path_first1 = str(askopenfile()).split(" ")
                path_of_wallpaper = path_first1[1].split("=")
                path_of_wallpaper = path_of_wallpaper[1].replace("/", "//").replace("'", "")
                try:
                    ctypes.windll.user32.SystemParametersInfoW(20, 0, path_of_wallpaper, 0)
                    speak("wallpaper is been changed")
                except Exception as j:
                    print("you have not inserted a image")
            if "show desktop" in message or "minimize"in message:
                os.startfile(r"C:\Users\USER\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\Shows Desktop.lnk")
            if 'musics' in message or "songs" in message or 'music' in message or "song" in message:
                path_folder = askdirectory(
                    title='Select Folder')  # Asking the user to select the folder to play the songs
                getting_files = getting_path_for_music(path_folder,".mp3")
                length_of_folder = len(getting_files) - 1
                first_mp3 = 0
                initial_volume = 0.6
                while True:
                    print(first_mp3)
                    query = "e"
                    if first_mp3 < length_of_folder:
                        try:
                            mixer.init()
                            mixer.music.load(getting_files[first_mp3])
                            mixer.music.set_volume(initial_volume)
                            mixer.music.play()
                        except Exception:
                            pass
                        while True:
                            query = input("press anyone (p),(r),(-),(+),(/n),(e),(pr):").lower()
                            if (query == 'p') or ("pause" in query):
                                mixer.music.pause()
                            elif (query == 'r') or ("resume" in query) or ("play" in query):
                                mixer.music.unpause()
                            elif (query == 'e') or ("exit" in query):
                                mixer.music.stop()
                                first_mp3 = length_of_folder
                                break
                            elif ("pr" in query) or ("privies" in query) or ("last" in query):
                                first_mp3 -= 1
                                break
                            elif ("+" in query) or ("increase volume" in query):
                                initial_volume += 0.1
                                mixer.music.set_volume(initial_volume)
                            elif ("-" in query) or ("decrease volume" in query):
                                initial_volume -= 0.1
                                mixer.music.set_volume(initial_volume)
                            else:
                                first_mp3 += 1
                                print("")
                                break
                    elif query == "e":
                        print("ok, what is the next command")
                        speak("ok, what is the next command")
                        break
                    else:
                        print("No mp3 files has been found")
                        speak("No mp3 files has been found")
                        break
            """if 'joke' in message:
                try:
                    res = requests.get('https://icanhazdadjoke.com/', headers={"Accept": "application/json"})
                    if res.status_code == requests.codes.ok:
                        print(str(res.json()['joke']))
                        speak(str(res.json()['joke']))
                    else:
                        print('oops!I ran out of jokes')
                        speak("oops!I ran out of jokes")
                except Exception:
                    print(pyjokes.get_joke())
                    speak(pyjokes.get_joke())"""
            for i in getting_path_for_music("fenesta for class",".lnk"):
                k = j = i.replace("fenesta for class/", "")
                if j.lower() in message.lower() + ".lnk":
                    k = "C:/Users/USER/.spyproject/fenesta/fenesta for class/" + k
                    os.startfile(k)


            if 'setting' in message:
                print(settings)
                change_settings = int(input('Enter the number to change the setting'))
                if change_settings == 1:
                    voice_type = input('𝖊𝖓𝖙𝖊𝖗 (VI) 𝖋𝖔𝖗 𝖛𝖔𝖎𝖈𝖊 𝖆𝖘𝖘𝖎𝖘𝖙𝖊𝖓𝖙 𝖔𝖗 (TP) 𝖋𝖔𝖗 𝖙𝖞𝖕𝖎𝖓𝖌:- ').upper()
                    if voice_type == 'VI':
                        to_check_net = True
                    else:
                        to_check_net = False

                elif change_settings == 2:
                    print(help)
                elif change_settings == 3:
                    exit()
                else:
                    print('invalid character')
            if 'send a mail' in message:
                try:
                    content = input("what should i send::-")
                    to = input("whom should i send")
                    sendEmail(to, content)
                    print("Email has been sent !")
                except Exception as e:
                    print(e)
                    print("I am not able to send this email")
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
