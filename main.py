import requests
import win32com.client as wincom
import json

speak = wincom.Dispatch("SAPI.SpVoice")

city = input("Enter the name of city\n")

url = f"https://api.weatherapi.com/v1/current.json?key=88d26221360d42cdbfe210332230108&q={city}"

r = requests.get(url)

weatherdic = json.loads(r.text)
temp = weatherdic["current"]["temp_c"]
print(temp)

text = "Python text-to-speech test. using win32com.client"
speak.Speak(f"The temperature of city {city} is {temp} degree celcius")
