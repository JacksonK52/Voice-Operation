Dim speech
Set speech = CreateObject("sapi.spvoice")
speech.rate = -2
speech.volume = 50
Set wshShell = CreateObject("WScript.Shell")
userName = wshShell.ExpandEnvironmentStrings("%USERNAME%")
If Hour(time)<12 then
	speech.speak "Good Morning " & userName
Else
	If Hour(time)=>16 then
		speech.speak "Good Evening " & userName
	Else
        speech.speak "Good Afternoon " & userName
	End If
End If