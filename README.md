# ClipSpeak

This is a variation of ClipMon (https://github.com/hikriss/clipmon). This app use Speech API from Microsoft to read the text in clipboard.

pyttsx is cross platform of text to speech libray. however, when I used with Wxpython, there is unpredictable behavior due to message pumping. I cannot solve it, therefore, I used SAPI directly.

however, to record it, if I need to use event from SAPI in the future, I can check the code from the link below:
https://github.com/RapidWareTech/pyttsx/blob/master/pyttsx/drivers/sapi5.py

Sample code to use SAPI 5.x from python

import win32com.client as wincl
speak = wincl.Dispatch("SAPI.SpVoice")
speak.Speak("Hello World")

https://msdn.microsoft.com/en-us/library/ee125024(v=vs.85).aspx
Speak function can have second parameters which is Flags

https://msdn.microsoft.com/en-us/library/ee431843(v=vs.85).aspx
SPEAKFLAGS
    SPF_ASYNC = 1
    SPF_PURGEBEFORESPEAK = 3
