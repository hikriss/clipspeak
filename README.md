# ClipSpeak
    This is a variation of ClipMon (https://github.com/hikriss/clipmon). This app use Speech API from Microsoft to read the text in clipboard.

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
