Dim speech
Dim m
Dim w
Set speech = CreateObject("sapi.spvoice")
speech.speak "Echelon activated"
speech.speak "Initializing"
If Hour(time)<12 then
	speech.speak "Good Morning sir"
Else
	If Hour(time)=>16 then
		speech.speak "Good Evening sir"
	Else
		speech.speak "Good Afternoon sir"
	End If
End If

If Hour(time)>12 then
    m = Hour(time)
    m = m - 12
    speech.speak "Its " & m & " O clock and" & Minute(time) & "minute"
End If

'Converting array value of month into word
If(Month(date)=1) Then
	m = "january"
ElseIf(Month(date)=2) Then
	m = "february"
ElseIf(Month(date)=3) Then
	m = "march"
ElseIf(Month(date)=4) Then
	m = "april"
ElseIf(Month(date)=5) Then
	m = "may"
ElseIf(Month(date)=6) Then
	m = "june"
ElseIf(Month(date)=7) Then
	m = "july"
ElseIf(Month(date)=8) Then
	m = "august"
ElseIf(Month(date)=9) Then
	m = "september"
ElseIf(Month(date)=10) Then
	m = "october"
ElseIf(Month(date)=11) Then
	m = "november"
ElseIf(Month(date)=12) Then
	m = "december"
End If

'Converting array value of day into word'
If(Weekday(date)=1) Then
	w = "sunday"
ElseIf(Weekday(date)=2) Then
	w = "monday"
ElseIf(Weekday(date)=3) Then
	w = "tuesday"
ElseIf(Weekday(date)=4) Then
	w = "wednesday"
ElseIf(Weekday(date)=5) Then
	w = "thursday"
ElseIf(Weekday(date)=6) Then
	w = "friday"
ElseIf(Weekday(date)=7) Then
	w = "saturday"
End If

speech.speak "Today is " & w & " " & m & " " & Day(date) & " " & Year(date)
