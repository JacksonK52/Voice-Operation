Dim speech
Dim h
Set speech = CreateObject("sapi.spvoice")
If Hour(time)>12 then
	h = Hour(tme)
	h = h - 12
	If Minute(time)>1 then
		speech.speak "Its " & h & "pm and " & Minute(time) & "minutes"
	Else
		speech.speak "Its " & h & "pm and " & Minute(time) & "minute"
	End If
Else
	If Minute(time)>1 then
		speech.speak "Its " & h & "am and " & Minute(time) & "minutes"
	Else
		speech.speak "Its " & h & "am and " & Minute(time) & "minute"
	End If
End If