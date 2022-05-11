import encodings
import json
import requests
import pandas as pd
import win32com.client as win32
from io import StringIO
import urllib
import os

numberOfErrors = 0

endpoint = "entries"
language_code = "en-gb"

dflist = []
errorlist = []
wordlist = []

print("\nHello! This program will create a Microsoft Excel Sheet to store all information about the words you have.\n")
print("It uses the public API of Oxford Dictionary.\n")
print("Before you may continue, please get yourself an API ID and an API key.\n")
print("You can get them here: " + "https://developer.oxforddictionaries.com/")
print("\nNOTE: For the free plan, a limit of 1000 words is imposed; once reached, please get a new API ID and Key.")

print("\n\n")

app_id = input("Please enter your API ID: ")
app_key = input("Please enter your API Key: ")

print("\n\n")

print("Please prepare a .txt file with your word list, place this in the same folder as this python script.")
filename = input("Please enter the name of the .txt file: ")

print("\nScanning file...\n")

with open(f"{os.path.dirname(os.path.abspath(__file__))}/{filename}.txt", 'r') as fp:
    numberOfLines = len(fp.readlines())

print(f"Scan complete. {numberOfLines} words detected.\n")

for id, line in enumerate(open(f"{os.path.dirname(os.path.abspath(__file__))}/{filename}.txt", "r").readlines()):
    
    word_id = line.strip()

    print(f"Attempting line {id+1} of {numberOfLines}...")

    try:
        url = "https://od-api.oxforddictionaries.com/api/v2/" + endpoint + "/" + language_code + "/" + word_id.lower()

        r = requests.get(url, headers = {"app_id": app_id, "app_key": app_key})

        if r.status_code == 200:
            
            result = r.json()["results"][0]["lexicalEntries"][0]["entries"][0]["senses"]
            df = pd.json_normalize(result)

            dflist.append(df)

            wordlist.append(word_id)

            print("Success!")
        
        elif r.status_code == 429:

            print("\nYou have exceed the 1000 words limit, exiting program with completed words so far...\n")

            break

        else:
            errorlist.append(id+1)
            print(f"\nThere was a problem calling Oxford API. Status code is {r.status_code}. The problem is with line {id+1}.")
            print("Possibly invalid word or multiple words not connected with dashes")
            print("Refer to HTTP documentations for more information.\n")

    except:
        numberOfErrors += 1
        errorlist.append(id+1)

        if numberOfErrors == 1:
            print(f"{numberOfErrors} unknown error found")

        else:
            print(f"{numberOfErrors} unknown errors found")
        
        continue

try:

    print("\nFinished query, constructing data sheet.\n")

    newdf = pd.concat(dflist, names=['word'], keys=[*wordlist])

    pd.DataFrame(newdf).to_excel("C:/Users/david/OneDrive - Raffles Institution/Documents/wordbank/wordbank.xlsx")

    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open("C:/Users/david/OneDrive - Raffles Institution/Documents/wordbank/wordbank.xlsx")
    ws = wb.Worksheets("Sheet1")
    ws.Columns.AutoFit()
    wb.Save()
    excel.Application.Quit()

    if errorlist:

        print("The following lines have errors (Seperated by space):")
        print(*errorlist)

    print("\n")
    print(f"The process has finished with {numberOfErrors} errors.")
    print("Please mind the JSON residual mess.")
    print("\n")

except:
    print("An error has occured whilst parsing the data sheet.\n")
