'#$language = "VBScript"
'#$interface = "1.0"
'Name: CNAMDVER.vbs

Option Explicit

Dim s_objScriptTab
Dim s_objStrings
Dim s_objUtils
Dim s_intTimeout
	
Dim s_strOrigTele	' to store user input

'********'*********'*********'*********'*********'*********'*********'*********'
'Main
'********'*********'*********'*********'*********'*********'*********'*********'
Sub Main
	Dim blnRet
	
	If (Not Initialize) Then
		Exit Sub
	End If
	s_objUtils.ScriptStart
	
	blnRet = s_objUtils.SendAndWaitFor("leave all", ">", s_intTimeout) 
	If (blnRet) Then 'This tests if you have access to the screen.
		If (get_user_data = vbOK) Then 
			Call issue_CNAMDVER_command
		End If 
	Else
		s_objUtils.CancelMsgBox
	End If 	
	
	s_objUtils.ScriptEnd
	Set s_objStrings = Nothing
	Set s_objUtils = Nothing
End Sub 'Main

'********'*********'*********'*********'*********'*********'*********'*********'
'get_user_data
'********'*********'*********'*********'*********'*********'*********'*********'
Function get_user_data()
	Dim blnTryAgain

	blnTryAgain = True
	Do 	
		s_strOrigTele = crt.Dialog.Prompt(s_objUtils.CleanTrashMsg("DN") _
				& "Enter DN:", "ENTER DN", s_strOrigTele)
		If (Len(s_strOrigTele) = 0) Then 
			get_user_data = vbCancel
			Exit Function
		End If

		' return only numeric characters...   
		s_strOrigTele = s_objStrings.RemoveNonNumericCharacters(s_strOrigTele)
		If (s_objUtils.TestLength(s_strOrigTele, "DN", 10)) then 
			blnTryAgain = False 
			get_user_data = vbOK
		End if
	Loop While blnTryAgain
End Function 'get_user_data

'********'*********'*********'*********'*********'*********'*********'*********'
'issue_CNAMDVER_command
'********'*********'*********'*********'*********'*********'*********'*********'
Sub issue_CNAMDVER_command()
	s_objUtils.SendAndWaitFor "leave all", ">", s_intTimeout
	s_objUtils.SendAndWaitFor "REPEAT 5 (cnamdver " & s_strOrigTele & " 0)", ">", s_intTimeout
End Sub 'issue_CNAMDVER_command


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