On Error Resume Next
strComputer = "."
Set objWMI = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMI.ExecQuery("Select * from Win32_Battery",,48)
For Each objItem in colItems
    batCharge = objItem.EstimatedChargeRemaining
    batStatus = objItem.BatteryStatus
Next

If batCharge > 80 And batStatus = 2 Then 'if charge is >80 and AC power is still on
	Dim objSAPI
	Set objSAPI = CreateObject("sapi.spvoice")
	objSAPI.speak "The battery is " & batCharge & " % charged. Please switch off charger."
	msgbox "The battery is " & batCharge & "% charged. Please switch off charger.", vbOKOnly+vbCritical+vbSystemModal, "Switch off charger"
End If

Wscript.Quit 0
