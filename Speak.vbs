Dim a, msg, sapi
a = 1
Do While a=1
msg=InputBox("Speak Now")
Set sapi=CreateObject("sapi.spvoice")
sapi.speak msg
loop