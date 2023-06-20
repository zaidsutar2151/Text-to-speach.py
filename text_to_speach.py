import win32com.client as wincom
speak = wincom.Dispatch("SAPI.SpVoice")
while(True):
    text = input("Enter what you want to speak ")
    if(text=="0"):
        break
    speak.Speak(text)
print("the program is finished")
