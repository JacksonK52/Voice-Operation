Dim speech
Dim m
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
speech.speak m & " " & Day(date) & " " & Year(date)
