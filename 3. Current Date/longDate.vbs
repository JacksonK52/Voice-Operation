Dim speech
Dim m
Dim w
Set speech = CreateObject("sapi.spvoice")
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
