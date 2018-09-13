#$language = "VBScript"
#$interface = "1.0"
Option Explicit

Const SCRIPTS_DIR = "C:\Scripts\"

Include SCRIPTS_DIR&"SecureCRT\Constants.txt"
'Include SCRIPTS_DIR&"Global\clsDialogBoxUtils.vbs"
Include SCRIPTS_DIR&"Global\clsStringUtils.vbs"
Include SCRIPTS_DIR&"Global\clsArrayUtils.vbs"
Include SCRIPTS_DIR&"SecureCRT\clsSecureCRTUtils.vbs"

'Dim DialogBox
Dim Strings
Dim Arrays
Dim Utils

'Set DialogBox = New DialogBoxUtils
Set Strings = New StringUtils
Set Arrays = New ArrayUtils
Set Utils = New SecureCRTUtilis

Dim s_strNextdata
Dim s_strOrig
Dim s_strO_Tele
Dim s_blnImage

' This automatically generated script may need to be
' edited in order to work correctly.

Sub Main
	Dim strPW
	
	scriptStart
	
'	strPW = testInputBox()
'	testSend
'	testSendAndWait strPW
'	crt.Sleep(2000)
	
'	testScreenCapture

	testGetSwitch
	
	scriptEnd
End Sub

Sub scriptStart
	crt.Session.SetStatusText "" & ": " & crt.ScriptFullName
	crt.Session.SetStatusText "Starting" & ": " &  crt.ScriptFullName 
	crt.Sleep(1000)
	crt.Session.SetStatusText "Running" & ": " &  crt.ScriptFullName 
End Sub

Sub scriptEnd
	crt.Session.SetStatusText "Done" & ": " &  crt.ScriptFullName 
	crt.Sleep(1000)
	crt.Session.SetStatusText "Ready" 
End Sub

Function testInputBox()
	Dim dlg, strPW

	strPW = InputBox("Please enter your password:")
	Set dlg = crt.Dialog
	strPW = dlg.Prompt("Enter your password:", "Logon Script", "", True)
	testInputBox = strPW
End Function

Sub testSend()
	crt.Screen.Send "ls"
	crt.Screen.Send " -ll"
	crt.Screen.Send Chr(13)	
	crt.Screen.Send ""  & Chr(13)
	crt.Session.SetStatusText "" & ": " & crt.ScriptFullName	
End Sub


Sub testSendAndWait(strPW)
	If strPW = "" Then
		strPW = "Great9#me"
'		MsgBox "User Canceled."
'		Exit Sub
	End If
		
	crt.Screen.Send "ssh -l ke066443 itplt04.lab.sprint.com" & chr(13)
	crt.Screen.WaitForString "ke066443@itplt04.lab.sprint.com's password: "
	crt.Screen.Send strPW & chr(13)
	crt.Screen.WaitForString "@itplt04:/home/ke066443>"
	crt.Sleep(2000)
	crt.Screen.Send "ll" & chr(13)
	crt.Screen.Send "exit" & Chr(13)

End Sub
	
Sub testScreenCapture()
	Dim Num1, Num2, strScreenData
	crt.Screen.Synchronous = False

	crt.Screen.Send "clear" & chr(13)
	crt.Sleep(500)
	Num1 = crt.Screen.CurrentRow
	crt.Screen.Send "ll" & chr(13)
	crt.Sleep(500)
	Num2 = crt.Screen.CurrentRow
	strScreenData = crt.Screen.Get(Num1+4,1,Num2-1,80)
End Sub 

Sub testGetSwitch()
	Dim strFileName
	Dim arrData
	Dim arrFile
	Dim s_strSiteid
	
	strFileName =  SCRIPTS_DIR + "UnitTests\TestData\" & Utils.getSessionFileName()
	crt.Dialog.Messagebox strFileName
	arrData = Array("9991112222")
	Arrays.WriteFile arrData, strFileName
	
	arrFile = Arrays.ReadFile(strFileName)
	s_strSiteid = arrFile(UBound(arrFile))
	crt.Dialog.MessageBox(s_strSiteid)
End Sub 



'*******************************************************************************
' Include is used to import code from external files at runtime.  Code 
' imported with this Sub can be called on as if it existed in this file.
Sub Include(strFileName)
	Dim objFSO, objOTF, strData
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objOTF = objFSO.OpenTextFile(strFileName, 1)
	strData = objOTF.ReadAll
	objOTF.Close
	ExecuteGlobal strData
End Sub


