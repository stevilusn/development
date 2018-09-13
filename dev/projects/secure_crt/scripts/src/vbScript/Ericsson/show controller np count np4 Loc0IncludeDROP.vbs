'#$language = "VBScript"
'#$interface = "1.0"
'Name: <SCRIPT NAME>.vbs

Option Explicit

Dim s_intTimeout
Dim s_objScriptTab	'Gets set during Initialize.
Dim s_objUtils
Dim s_objDialogBox
'Dim s_objStrings
'Dim s_objArrays
'Dim s_objWatchFor
Dim s_strPort
'Dim s_strSvlan

'********'*********'*********'*********'*********'*********'*********'*********'
'Main
'********'*********'*********'*********'*********'*********'*********'*********'
Sub Main
	Dim blnRet
	 	
	If (Not Initialize) Then
		Exit Sub
	End If
	s_objUtils.ScriptStart
	
	blnRet = s_objUtils.SendAndWaitFor("", "#", s_intTimeout) 
	If (blnRet) Then 'This tests if you have access to the screen.
			Call issue_PORTSVLAN_command
	Else
		s_objUtils.CancelMsgBox
	End If 	
	
	s_objUtils.ScriptEnd
	Set s_objUtils = Nothing
	Set s_objDialogBox = Nothing
'	Set s_objStrings = Nothing
'	Set s_objArrays = Nothing
'	Set s_objWatchFor = Nothing
End Sub 'Main

'********'*********'*********'*********'*********'*********'*********'*********'
'issue_PORTSVLAN_command
'********'*********'*********'*********'*********'*********'*********'*********'
Sub issue_PORTSVLAN_command()
s_objUtils.Send "show controller np count np4 loc 0/3/CPU0 | include DROP" & chr(13) + _
                "show controller np count np4 loc 0/4/CPU0 | include DROP" & chr(13) + _
                "show controller np count np4 loc 0/5/CPU0 | include DROP" & chr(13) + _
                "show controller np count np4 loc 0/6/CPU0 | include DROP" & chr(13) + _
                "show controller np count np4 loc 0/7/CPU0 | include DROP"
'	s_objUtils.SendAndWaitFor "show circuit counters " & s_strPort & " | include " & s_strSvlan, "#", s_intTimeout
End Sub 'issue_PORTSVLAN_command



'********'*********'*********'*********'*********'*********'*********'*********'
'Initialize the system by importing Constants.txt and classes, setting up 
'objects, and testing the connection.
'********'*********'*********'*********'*********'*********'*********'*********'
Function Initialize()
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
