# $language = "VBScript"
# $interface = "1.0"
' GoogleSelectedText.vbs
'
' Description:
' When this script is launched, the text selected within the terminal
' window is used as the search term for a web search using google.com.
' This script demonstrates capabilities only available in SecureCRT 6.1
' and later (Screen.Selection property).
'
' Demonstrates:
' - How to use the Screen.Selection property in SecureCRT 6.1 and later
' to get access to the text selected in the terminal window.
' - How to use the WScript.Shell object to launch an external application.
' - How to branch code based on the version of SecureCRT in which this script
' is being run.
Sub Main()
	' Extract SecureCRT's version components to determine how to go about
	' getting the current selection (version 6.1 provides a scripting API
	' for accessing the screen's selection, but earlier versions do not)
	strVersionPart = Split(crt.Version, " ")(0)
	vVersionElements = Split(strVersionPart, ".")
	nMajor = vVersionElements(0)
	nMinor = vVersionElements(1)
	nMaintenance = vVersionElements(2)
	If nMajor >= 6 And nMinor > 0 Then
		' Use available API to get the selected text:
		strSelection = Trim(crt.Screen.Selection)
	Else
		MsgBox "The Screen.Selection object is available" & vbCrLf & _
		"in SecureCRT version 6.1 and later." & vbCrLf & _
		vbCrLf & _
		"Exiting script."
		Exit Sub
	End If
	' Now search on Google for the information.
	g_strSearchBase = "http://www.google.com/search?hl=en&q="
	Set g_shell = CreateObject("WScript.Shell")
	' Instead of launching Internet Explorer, we'll run the URL, so that the
	' default browser gets used :).
	If strSelection = "" Then
		g_shell.Run Chr(34) & "http://www.google.com/" & Chr(34)
	Else
		g_shell.Run Chr(34) & g_strSearchBase & strSelection & Chr(34)
	End If
End Sub