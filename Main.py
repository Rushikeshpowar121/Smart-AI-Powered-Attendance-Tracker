import time
import speech_recognition as sr
import pyttsx3
import datetime
import pandas as pd
import sys
import os
import openpyxl


def hello():
    hour = datetime.datetime.now().hour
    if 0 <= hour < 12:
        speak("Good morning, sir.")
        print("Good morning, sir.")
    elif 12 <= hour < 18:
        speak("Good afternoon, sir.")
        print("Good afternoon, sir.")
    else:
        speak("Good evening, sir.")
        print("Good evening, sir.")

def thank_you():
    sys.exit()

def choose_subject():
    os.system('cls' if os.name == 'nt' else 'clear')
    speak("Please choose a subject for which you want to take attendance.")
    print("I am your attendance bot. Please choose a subject for which you want to take attendance.")
    speak("Here are some subjects:")
    print("Here are some subjects:")
    speak("1. Management")
    print("1. Management")
    speak("2. Programming with Python")
    print("2. Programming with Python")
    speak("3. Emerging Trend in Information Technology")
    print("3. Emerging Trend in Information Technology")
    speak("4. Network and Information Security")
    print("4. Network and Information Securit")
    speak("5. Mobile Application Development")
    print("5. Mobile Application Development")
    speak("6. Exit")
    print("6. Exit")
    subject = int(input("Choose subject: "))
    if subject == 1:
        return "Management"
    elif subject == 2:
        return "Programming with Python"
    elif subject == 3:
        return "Emerging Trend in Information Technology"
    elif subject == 4:
        return "Web page development using PHP"
    elif subject == 5:
        return "Mobile Application Development"
    elif subject == 6:
        thank_you()
    else:
        print("Please choose a correct subject.")
        speak("Please choose a correct subject.")
        return choose_subject()

def recognize_speech():
    listener = sr.Recognizer()
    try:
        with sr.Microphone() as source:
            print("Listening...")
            voice = listener.listen(source)
            command = listener.recognize_google(voice)
            print(command)
            return command
    except:
        pass

def take_attendance(subject):
    speak(f"Taking attendance for {subject}")
    absent_numbers = []
    for i in range(1, 5):
        text = f"Number {i}"
        speak(text)
        print(f"Attendance for {text}")
        response = recognize_speech()
        print(response)
        if response and "present" in response.lower():
            print(f"Roll number {i} is present")
        else:
            print(f"Roll number {i} is absent")
            absent_numbers.append(i)
    print(absent_numbers)
    return absent_numbers


def pronounce_absent(absent_numbers, subject_name):
    print("Absent numbers are:", absent_numbers)
    date = datetime.date.today()
    time_ = time.localtime()
    current_time = time.strftime("%H:%M:%S", time_)
    data = [(date, current_time, absent_numbers)]
    df = pd.DataFrame(data, columns=["Date", "Time", "Absentees"])

    # Print the DataFrame to check its contents
    print("DataFrame:")
    print(df)

    file_exists = os.path.isfile("test.xlsx")

    with pd.ExcelWriter("test.xlsx", engine="openpyxl", mode="a") as writer:
        if file_exists:
            existing_sheets = pd.ExcelFile("test.xlsx").sheet_names
            if subject_name in existing_sheets:
                print(f"Sheet '{subject_name}' already exists. Appending data to existing sheet.")
                df.to_excel(writer, index=False, header=False, sheet_name=subject_name, startrow=len(pd.read_excel("test.xlsx", sheet_name=subject_name)))
            else:
                df.to_excel(writer, index=False, header=False, sheet_name=subject_name)
        else:
            df.to_excel(writer, index=False, header=False, sheet_name=subject_name)

    speak("Absent numbers are")
    speak(absent_numbers)
    speak("Thank you.")
    speak("Do you want to continue? If yes, press 'y'. If no, press 'n'.")
    response = input("Do you want to continue? If yes, press 'y'. If no, press 'n'.")
    if response.lower() == 'y':
        pass
    else:
        thank_you()


def speak(text):
    engine.say(text)
    engine.runAndWait()

engine = pyttsx3.init()
voices = engine.getProperty("voices")
engine.setProperty("voice", voices[1].id)

hello()
while True:
    subject_name = choose_subject()
    absent_numbers = take_attendance(subject_name)
    pronounce_absent(absent_numbers, subject_name)
