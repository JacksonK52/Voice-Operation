Dim speech
Set speech = CreateObject("sapi.spvoice")
speech.rate = -2
If Minute(time)>1 then
	speech.speak "Its " & Hour(time) & " and " & Minute(time) & "minutes"
Else
	speech.speak "Its " & Hour(time) & " and " & Minute(time) & "minute"
End If