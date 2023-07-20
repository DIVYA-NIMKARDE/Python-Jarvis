import os
import speech_recognition as sr
import win32com.client
import webbrowser
import openai
import datetime
import random

# Update the API key here
apikey = "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"

speaker = win32com.client.Dispatch("SAPI.SpVoice")

chatStr = ""

def chat(query):
    global chatStr
    print(chatStr)
    openai.api_key = apikey
    chatStr += f"Divya: {query}\n Jarvis: "
    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=chatStr,
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    say(response["choices"][0]["text"])
    chatStr += f"{response['choices'][0]['text']}\n"
    return response["choices"][0]["text"]

def ai(prompt):
    openai.api_key = apikey
    text = f"OpenAI response for Prompt: {prompt} \n *************************\n\n"
    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=prompt,
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    text += response["choices"][0]["text"]
    if not os.path.exists("Openai"):
        os.mkdir("Openai")
    with open(f"Openai/{''.join(prompt.split('intelligence')[1:]).strip() }.txt", "w") as f:
        f.write(text)

def say(text):
    speaker.Speak(text)

def takecommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        audio = r.listen(source)
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query
        except Exception as e:
            print(f"An error occurred: {str(e)}")
            return "Some error occurred! Sorry."

if __name__ == '__main__':
    j = "Hi, I am shambhu. How can I help you Divya?"
    speaker.Speak(j)
    while True:
        print('Listening...')
        query = takecommand()
        sites = [["youtube", "https://youtube.com"], ["wikipedia", "https://wikipedia.com"], ["google", "https://google.com"]]
        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                speaker.Speak(f"Opening {site[0]}")
                webbrowser.open(site[1])

        if "play music" in query:
            music_url = "D:\saudebazi.mp3"  # Update the correct music file path
            speaker.Speak("Playing music...")
            os.system(f"start {music_url}")

        if "the time" in query:
            current_time = datetime.datetime.now().strftime("%H:%M")
            speaker.Speak(f"Sir, the time is {current_time}")

        if "start camera" in query.lower():
            os.system("start microsoft.windows.camera:")

        elif "using artificial intelligence" in query.lower():
            ai(prompt=query)

        elif "bie shamhu" in query.lower():
            exit()

        elif "delete chat" in query.lower():
            chatStr = " "

        else:
            print("Chatting...")
            chat(query)
