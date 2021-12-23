import requests
import json


def speak(st):
    from win32com.client import Dispatch

    speak = Dispatch("SAPI.SpVoice")

    speak.Speak(st)


if __name__ == '__main__':
    url = ('https://newsapi.org/v2/top-headlines?'
           'q=business&'
           'from=2021-12-21&'
           'to=2021-12-30&'
           'sortBy=popularity&'
           'apiKey=4f81d511df2542bdafe823645662004e')

    try:
        response = requests.get(url)
        read = json.loads(response.text)

        i = 1
        for news in read['articles']:
            print(f"Title {i} ->", str(news['title']))
            speak(str(i) + " Title News is")
            speak(str(news['title']))
            print(f"Description {i} ->", str(news['description']))
            speak(str(i) + " news description is")
            speak(str(news['description']))
            print("-------------------------------xxx----------------------\n")
            i += 1


    except Exception as e:
        print(e, "Some error occurred.Requests has not respond.")
