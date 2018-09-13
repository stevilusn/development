'#$language = "VBScript"
'#$interface = "1.0"
'Name: DEFAULT_ClearStatusText.vbs

Option Explicit

Dim s_objScriptTab

'********'*********'*********'*********'*********'*********'*********'*********'
'Occationally a script will end without clearing the SecureCRT Status Text line.
'If this is an issue, this script will re-set the test to "Ready".
'********'*********'*********'*********'*********'*********'*********'*********'
Sub Main()
	Set s_objScriptTab = crt.GetScriptTab
	s_objScriptTab.Session.SetStatusText "Ready"
	Set s_objScriptTab = Nothing 	
End Sub