Dim objShell, lngMinutes, boolValid

Set objShell = CreateObject("WScript.Shell")
lngMinutes = InputBox("How long you want to keep your system awake?" & Replace(Space(5), " ", vbNewLine) & "Enter minutes:", "Awake Duration") 'we are replacing 5 spaces with new lines
If lngMinutes = vbEmpty Then 'If the user opts to cancel the process
	'Do nothing
Else
	On Error Resume Next
	Err.Clear
	boolValid = False
	lngMinutes = CLng(lngMinutes)
	If Err.Number = 0 Then 'input is numeric
		If lngMinutes > 0 Then 'input is greater than zero
			For i = 1 To lngMinutes
				WScript.Sleep 60000 '60 seconds
				objShell.SendKeys "{SCROLLLOCK 2}"
			Next
			boolValid = True
			MsgBox "Forced awake time over. Back to normal routine.", vbOKOnly+vbInformation, "Task Completed"
		End If
	End If
	On Error Goto 0
	If boolValid = False Then
		MsgBox "Incorrect input, script won't run" & vbNewLine & "You can only enter a numeric value greater than zero", vbOKOnly+vbCritical, "Task Failed"
	End If
End If

Set objShell = Nothing
Wscript.Quit 0
