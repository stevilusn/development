# $language = "VBScript"
# $interface = "1.0"
' SendToAll.vbs
Option Explicit

Sub Main()
	Dim objCurrentTab
	Dim objScriptTab
	Dim blnJumpboxExists
	Dim blnSkipJumpboxs
	Dim strCommand
	Dim intIndex
	
	blnJumpboxExists = False
	blnSkipJumpboxs = False
		
	if Not crt.Session.Connected then
		szSession = crt.Dialog.Prompt("Enter session:", "", "", False)
		if szSession = "" then exit sub
		
		crt.Session.ConnectInTab("/S " & szSession)
		crt.Session.ConnectInTab("/S " & szSession)
		crt.Session.ConnectInTab("/S " & szSession)
	end if
	
	' Find out what should be sent to all tabs
	strCommand = crt.Dialog.Prompt("Enter command to be sent to all tabs:", _
	"Send To All Connected Tabs", "clear; ls -ll", False)
	if strCommand = "" then exit Sub
	
	For intIndex = 1 To crt.GetTabCount
		Set objCurrentTab = crt.GetTab(intIndex)
		If (InStr(objCurrentTab.Caption, "Jumpbox") > 0) Then 
			blnJumpboxExists = True
			Exit For 
		End If 
		Set objCurrentTab = nothing
	Next
	If (blnJumpboxExists And crt.Dialog.Messagebox("Skip Jumpboxes?", "SKIP Jumpboxes", vbQuestion + vbYesNo) = vbYes) Then
		blnSkipJumpboxs = true
	End If
	
	If crt.Dialog.MessageBox(_
			"Are you sure you want to send the following command to " & _
			"__ALL__ tabs?" & vbcrlf & vbcrlf & strCommand, _
			"Send Command To All Tabs - Confirm", _
			vbYesNo) = vbYes  Then
		' Connect to each tab in order from left to right, issue a command, and
		' then disconnect...
		For intIndex = crt.GetTabCount to 1 Step -1
			Set objCurrentTab = crt.GetTab(intIndex)
			If (blnSkipJumpboxs = False Or (blnSkipJumpboxs = True And InStr(objCurrentTab.Caption, "Jumpbox") = 0)) Then 
				objCurrentTab.Activate
				if (objCurrentTab.Session.Connected = True) then
					crt.Sleep 500
					objCurrentTab.Screen.Send strCommand & vbcr
					crt.Sleep 1000
				end If
			End if
		Next
		
		Set objScriptTab = crt.GetScriptTab()
		objScriptTab.Activate
		
		crt.Dialog.Messagebox "The following command was sent to all connected tabs:" _
				& vbcrlf & vbcrlf & strCommand, "REVIEW", vbInformation + vbOKOnly
	End If 
End Sub
