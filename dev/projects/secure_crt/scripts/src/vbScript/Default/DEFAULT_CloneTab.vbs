# $language = "VBScript"
# $interface = "1.0"
' CloneTab.vbs
' Takes a session window and creates a duplicate (clone) window
Option Explicit

Sub Main()
	Dim objScriptTab
	Dim objConfig
	Dim strSession
	Dim objClonedTab
	
	Set objScriptTab = crt.GetScriptTab()
	
'This is left as an example.  It's faster to use the SecureCRT tab pop-up 
'menu to clone.  But this shows a cleaner method of cloning a tab in code 
'since authentication has already been done.
'	Set objClonedTab = objScriptTab.Clone
'	MsgBox "Cloned Tab index: " & objClonedTab.Index & vbcrlf & _
'			"Cloned Tab Name: " & objClonedTab.Caption  
	
'This has the added benifit of allowing the user to edit which tab to clone 
'before executing.  This will re-authenticate the connection.
	strSession = crt.Dialog.Prompt("Enter session:", "", """" & objScriptTab.Session.Path & """", False)
	If (strSession = "") then 
		Exit Sub
	End If 
	Set objConfig = objScriptTab.Session.Config
	Set objClonedTab = crt.Session.ConnectInTab("/S " & strSession)
'	If MsgBox("Current Tab index: " & objClonedTab.Index & vbcrlf & _
'			"Current Tab name: " & objClonedTab.Caption & vbcrlf & vbcrlf & _
'			"Cloning OK...", vbokCancel) <> vbOK then exit sub

End Sub
