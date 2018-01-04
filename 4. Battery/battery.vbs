Dim val
Dim speech
Set speech=CreateObject("sapi.spvoice")

Function IsLaptop( myComputer )
    On Error Resume Next
    Set objWMIService = GetObject( "winmgmts://" & myComputer & "/root/cimv2" )
    Set colItems = objWMIService.ExecQuery( "Select * from Win32_Battery" )
    IsLaptop = True
    For Each objItem in colItems
        IsLaptop = False

		speech.speak "Checking for battery information"
	'Manufacture Name
		If IsNumeric(val) Then 
			Set objRegEx = CreateObject("VBScript.RegExp")
			objRegEx.Global = True
			objRegEx.Pattern = "[^A-Za-z ]"
			val = objItem.DeviceID
			val = objRegEx.Replace(val, "")
			speech.speak "Manufacture by " & val
		End If

	'Voltage Calculation
		val = objItem.DesignVoltage
		val = val/1000
		speech.speak "Voltage is " & val

	'Battery Type
		speech.speak "Type " & objItem.Description

	'Chemical use in battery
		val = objItem.Chemistry
		If(val=1) Then
			speech.speak "Chemical type other"
		Elseif(val=2) Then
			speech.speak "Chemical type unknown"
		Elseif(val=3) Then
			speech.speak "Chemical type Lead Acid"
		Elseif(val=4) Then
			speech.speak "Chemical type Nickel cadmium"
		Elseif(val=5) Then
			speech.speak "Chemical type nickel metal hydride"
		Elseif(val=6) Then
			speech.speak "Chemical type lithium ion"
		Elseif(val=7) Then
			speech.speak "Chemical type zinc air"
		Elseif(val=8) Then
			speech.speak "Chemical type lithium polymer"
		End if

	'Battery Status
		val = objItem.BatteryStatus
		If(val=1) Then
			speech.speak "Charger not pluggin"
			'This calculation does not work while battery is charging 
			speech.speak "Estimated Remaining Time is " & CLng(objItem.EstimatedRunTime/60) & " hours and" & CLng(((objItem.EstimatedRunTime/60) - CLng(objItem.EstimatedRunTime/60)) * 100) & " minutes"
		Elseif(val=2) Then
			speech.speak "Charger is pluggin"
		Elseif(val=3) Then
			speech.speak "Battery is fully charged"
		Elseif(val=4) Then
			speech.speak "Battery is low"
		Elseif(val=5) Then
			speech.speak "Battery is in critical condition"
		Elseif(val=6) Then
			speech.speak "Battery is charging"
		Elseif(val=7) Then
			speech.speak "Battery is charging and high"
		Elseif(val=8) Then
			speech.speak "Battery is charging and low"
		Elseif(val=9) Then
			speech.speak "Battery is charging and it is in critical"
		Else
			speech.speak "Battery status unknown"
		End if

	'Battery Percentage
		speech.speak "Battery percent is " & objItem.EstimatedChargeRemaining

	'Check Battery condition
		val = objItem.Status
		If(val="OK") Then
			speech.speak "Battery is in good condition"
		Else
			speech.speak "Battery may be in bad condition"
		End If
    Next
    If Err Then Err.Clear
    On Error Goto 0
End Function
If IsLaptop( "." ) Then
	speech.speak "Unable to check battery"
	speech.speak "You are using desktop or server"
End If