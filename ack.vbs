Dim msg, sapi
msg=("acknowledge the document by writing in it and saving or closing it......... or else")
Set sapi=CreateObject("sapi.spvoice")
sapi.Speak msg
