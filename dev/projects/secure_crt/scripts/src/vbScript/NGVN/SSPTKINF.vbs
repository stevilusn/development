'#$language = "VBScript"
'#$interface = "1.0"
'Name: SSPTKINF.vbs

Option Explicit

Dim s_objScriptTab
Dim s_objUtils
Dim s_intTimeout

Dim s_strTWC_LATA	' to store user input
'Dim s_strOrig	'TODO: Variable not used.  What was it's purpose?  Can it be removed?

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
		s_objUtils.Message "Running: get_user_data"
		If (get_user_data() = vbOK) Then 
			s_objUtils.Message "Running: issue_SSPTKINF_command"
			Call issue_SSPTKINF_command
		End If 
	Else
		s_objUtils.CancelMsgBox
	End If 
		
	s_objUtils.ScriptEnd
	Set s_objUtils = Nothing 
End Sub 'Main

'********'*********'*********'*********'*********'*********'*********'*********'
'get_user_data
'********'*********'*********'*********'*********'*********'*********'*********'
Function get_user_data()
	Dim blnTryAgain

	blnTryAgain = True 
	Do 
		s_strTWC_LATA = crt.Dialog.Prompt(s_objUtils.CleanTrashMsg("LATA") _
				& "Enter LATA:", "ENTER LATA", s_strTWC_LATA)
				
		'  clean the spaces, underlines and dashes from the orig and term teles...   
		s_strTWC_LATA = s_objUtils.stripTN(s_strTWC_LATA)
		If (Len(s_strTWC_LATA) = 0) Then
			get_user_data = vbCancel
			blnTryAgain = False 
		ElseIf (Not s_objUtils.TestLength(s_strTWC_LATA, "", 3)) Then 
			blnTryAgain = True 
		Else
			get_user_data = vbOK
			blnTryAgain = False
		End If 
	Loop While (blnTryAgain)
End Function 'get_user_data

'********'*********'*********'*********'*********'*********'*********'*********'
'issue_SSPTKINF_command
'********'*********'*********'*********'*********'*********'*********'*********'
Sub issue_SSPTKINF_command()
	s_objUtils.SendAndWaitFor "leave all", ">", s_intTimeout
	s_objUtils.Send "Table SSPTKINF;lis all (2 eq "  & s_strTWC_LATA & ")"
End Sub 'issue_SSPTKINF_command


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


