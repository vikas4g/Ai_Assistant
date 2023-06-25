import datetime
import os.path
import webbrowser
import openai
import random
from config import apikey
import speech_recognition as sr
import win32com.client

speaker = win32com.client.Dispatch("SAPI.SpVoice")

# speech recognition
chatStr = ""


def chat(query):
    openai.api_key = apikey
    global chatStr
    # print(chatStr)
    chatStr += f"Vikas: {query}\n jarvis: "

    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=chatStr,
        temperature=1,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )

    speaker.speak(response["choices"][0]["text"])
    chatStr = f" {response['choices'][0]['text']}\n"
    return response["choices"][0]["text"]
    # if not os.path.exists("openAI"):
    #     os.mkdir("OpenAI")
    #
    # with open(f"openAI/{prompt[0:30]}.txt", "w") as f:
    #     f.write(text)


def ai(prompt):
    openai.api_key = apikey

    text = f"responce for openai prompt : {prompt} \n #####################\n\n"
    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=prompt,
        temperature=1,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )

    print(response["choices"][0]["text"])
    text += response["choices"][0]["text"]

    if not os.path.exists("openAI"):
        os.mkdir("OpenAI")

    with open(f"openAI/{prompt[0:30]}.txt", "w") as f:
        f.write(text)


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 0.5
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language="en-in")
            print(f"User Said : {query}")
            return query
        except Exception as e:
            return "some error ocured"


if __name__ == '__main__':
    print("WELLCOME")
    # text = input()
    speaker.speak("Hello Sir I'M jarvise")
    while True:
        print("listinig.....")
        query = takeCommand()
        speaker.speak(query)
        sites = [["youtube", "https://youtube.com"], ["google", "https://google.com"]]
        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                speaker.speak(f" Opening {site[0]} Sir..")
                webbrowser.open(site[1])

            if "the time" in query:
                hour = datetime.datetime.now().strftime("%H")
                min = datetime.datetime.now().strftime("%M")
                speaker.speak(f"sir the time is {hour} o'clock {min} minutes")

            if "using ai".lower() in query.lower():
                ai(prompt=query)

            else:
                chat(query)
