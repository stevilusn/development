'#$language = "VBScript"
'#$interface = "1.0"
'Name: C7RTESET.vbs

Option Explicit

Dim s_intTimeout
Dim s_objScriptTab	'Gets set during Initialize.
Dim s_objUtils

'********'*********'*********'*********'*********'*********'*********'*********'
'Main
'********'*********'*********'*********'*********'*********'*********'*********'
Sub Main
	dim strC7rteset

	If (Not Initialize) Then Exit Sub
	s_objUtils.ScriptStart
		
	If (s_objUtils.QuitAll(10)) Then 'This tests if you have access to the screen.
		If (s_objUtils.Prompt("Enter the RouteSet:", "ENTER ROUTESET", strC7rteset) = False) Then
			ExitMain : Exit Sub 
		End If
		s_objUtils.SendAndWaitFor "", ">", s_intTimeout
		s_objUtils.SendAndWaitFor "mapci;mtc;ccs;ccs7;c7rteset;post c " & strC7rteset, ">", s_intTimeout
	Else 
		s_objUtils.CancelMsgBox
	End If

	ExitMain
End Sub 'Main
Sub ExitMain()
	s_objUtils.ScriptEnd
	Set s_objUtils = Nothing
	Set s_objScriptTab = Nothing 
End Sub 'ExitMain


'********'*********'*********'*********'*********'*********'*********'*********'
'Initialize the system by importing Constants.txt and classes, setting up 
'objects, and testing the connection.
'********'*********'*********'*********'*********'*********'*********'*********'
Function Initialize
	Dim objFSO, objFile, strFileData, strCurPath
	
	Initialize = False 
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strCurPath = objFSO.GetParentFolderName(crt.ScriptFullName) & "\"
	If objFSO.FileExists(strCurPath & "BootStrap.txt") Then
		Set objFile = objFSO.OpenTextFile(strCurPath & "BootStrap.txt")
		strFileData = objFile.ReadAll
		objFile.Close
		Set objFile = Nothing
		ExecuteGlobal strFileData	
		If (BootStrap) Then
			Initialize = True 
		End If
	Else
		MsgBox "Could not find file " & strCurPath & "BootStrap.txt.", vbOK & vbCritical, "FILE NOT FOUND!"
	End If 
	s_intTimeout = DEFAULT_TIMEOUT
	Set objFSO = Nothing
End Function 'Initialize
