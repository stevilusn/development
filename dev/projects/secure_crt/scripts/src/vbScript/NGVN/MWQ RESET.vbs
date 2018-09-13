#$language = "VBScript"
#$interface = "1.0"
'Name: MWQ RESET.vbs

Option Explicit

Dim s_objScriptTab
Dim s_objUtils
Dim s_objStrings
Dim s_intTimeout

'********'*********'*********'*********'*********'*********'*********'*********'
'Main
'********'*********'*********'*********'*********'*********'*********'*********'
Sub Main
	Dim blnRet
	Dim strMwq
	Dim strReturn
	
	If (Not Initialize) Then
		Exit Sub
	End If
	s_objUtils.ScriptStart

	blnRet = s_objUtils.SendAndWaitFor("leave all", ">", s_intTimeout) 
	If (blnRet) Then 'This tests if you have access to the screen.
		strMwq = s_objUtils.ReadMwqDn()
		
		strReturn = crt.Dialog.Prompt("DN:", "MWQ RESET", strMwq)
		'TODO: Should strMwq be TestLength()?  YES
		strMwq = s_objStrings.RemoveNonNumericCharacters(strReturn)
	
		If (Len(strMwq) > 0) Then
			s_objUtils.SendAndWaitFor "MWQ", ">", s_intTimeout
			s_objUtils.SendAndWaitFor "STATUS " & strMwq, ">", s_intTimeout
			s_objUtils.SendAndWaitFor "RESET "  & strMwq, ">", s_intTimeout
			s_objUtils.SendAndWaitFor "STATUS " & strMwq, ">", s_intTimeout
			s_objUtils.Send "leave all"
		
			s_objUtils.WriteMwqDn(strMwq)
		End If
	Else
		s_objUtils.CancelMsgBox
	End If 

	s_objUtils.ScriptEnd
	Set s_objStrings = Nothing
	Set s_objUtils = Nothing
End Sub 'Main


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