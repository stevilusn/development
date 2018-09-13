'#$language = "VBScript"
'#$interface = "1.0"
'Name: POST DN.vbs

Option Explicit

Dim s_objScriptTab
Dim s_objUtils
Dim s_objStrings
Dim s_intTimeout

Dim s_strInfodigs	' to store user input
Dim s_strO_Tele	' to store user input
Dim s_strOrig

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
		s_strO_Tele = s_objUtils.ReadMwqDn()
		
		s_objUtils.Message "Running: get_user_data"
		If (get_user_data() = vbOK) Then 
			s_objUtils.Message "Running: issue_POST_command"
			Call issue_POST_command
			s_objUtils.Message "Running: Saving MWQ"
			s_objUtils.WriteMwqDn(s_strO_Tele)
		End if
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
	Dim strReturn
	
	blnTryAgain = True 
	Do 
		strReturn = crt.Dialog.Prompt(s_objUtils.CleanTrashMsg("DN") _
				& "Enter Phone#:", "ENTER DN", s_strO_Tele)
		If (Len(strReturn) = 0) Then
			get_user_data = vbCancel
			Exit Function
		End If
		
		s_strO_Tele = s_objStrings.RemoveNonNumericCharacters(strReturn)  
		s_strOrig = s_strInfodigs & s_strO_Tele   '  dinatest wants the infodigs and orig tele together   'FIXME:
		If (s_objUtils.TestLength(s_strOrig, "DN", 10)) Then 
			blnTryAgain = False
		End If 
	Loop While blnTryAgain
	get_user_data = vbOk
End Function  'get_user_data

'********'*********'*********'*********'*********'*********'*********'*********'
'issue_POST_command
'********'*********'*********'*********'*********'*********'*********'*********'
Sub issue_POST_command()
	s_objUtils.SendAndWaitFor "leave all", ">", s_intTimeout
	s_objUtils.Send "TABLE SERVCHNG;HEADING;POS " & s_strOrig
End Sub 'issue_POST_command


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