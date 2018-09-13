'#$language = "VBScript"
'#$interface = "1.0"
'Name: NODISP.vbs

Option Explicit

Dim s_intTimeout
Dim s_objScriptTab	'Gets set during Initialize.
Dim s_objUtils

'********'*********'*********'*********'*********'*********'*********'*********'
'Main
'********'*********'*********'*********'*********'*********'*********'*********'
Sub Main
 	
	If (Not Initialize) Then Exit Sub
	s_objUtils.ScriptStart
	
	crt.Sleep 2000
	If (s_objUtils.QuitAll(s_intTimeout)) Then 'This tests if you have access to the screen.
		s_objUtils.Send "MAPCI NODISP PRTMAP;MTC;TRKS;TTP"
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
