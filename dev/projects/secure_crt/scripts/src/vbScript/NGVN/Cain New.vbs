'#$language = "VBScript"
'#$interface = "1.0"
'Name: Cain New.vbs

Option Explicit

Dim s_objScriptTab
Dim s_objUtils
Dim s_objDialogBox
Dim s_objStrings
Dim s_intTimeout

Dim s_strInfodigs	' to store user input  'TODO: s_strInfodigs is never set, so in effect: Variable not used.  What was it's purpose?  Can it be removed?
Dim s_strOrigTele	' to store user input
Dim s_strTermTele	' to store user input
Dim s_strCtc		' to store user input

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
		Call get_MwqDN
		If (get_user_data = vbOK) Then
			s_objUtils.WriteMwqDn(s_strOrigTele) 
			Call issue_Diskut_commands
			Call issue_caintest_command
		End If 
	Else
		s_objUtils.CancelMsgBox
	End If 
	
	s_objUtils.ScriptEnd
	Set s_objDialogBox = Nothing
	Set s_objStrings = Nothing
	Set s_objUtils = Nothing
End Sub 'Main

'********'*********'*********'*********'*********'*********'*********'*********'
'get_MwqDN
'********'*********'*********'*********'*********'*********'*********'*********'
Sub get_MwqDN
	Dim strMwq_Dn

	strMwq_Dn = s_objUtils.ReadMwqDn()
	s_strOrigTele = strMwq_Dn
	s_strTermTele = strMwq_Dn
End Sub 'get_MwqDN

'********'*********'*********'*********'*********'*********'*********'*********'
'get_user_data
'********'*********'*********'*********'*********'*********'*********'*********'
Function get_user_data()
	Dim blnTryAgain
	
	blnTryAgain = True 
	Do 	
		With s_objDialogBox
			.Clear 
			.Method = "HTA"		
			.Height = 200
			.Width = 375
			.LabelWidths = "75px"
			.Title = "Class 4 Cain Test"
			.AddItem Array("type=paragraph", "name=label1", "value=" & s_objUtils.CleanTrashMsg("input"), "terminate=newline")
			.AddItem Array("type=text", "name=s_strOrigTele", "accesskey=r", "label=Orig TN:", "value=" & s_strOrigTele, "size=10", "terminate=newline")
			.AddItem Array("type=text", "name=s_strTermTele", "accesskey=e", "label=Term TN:", "value=" & s_strTermTele, "size=10", "terminate=newline")
			.AddItem Array("type=text", "name=s_strCtc", "accesskey=T", "label=CTC:", "value=" & s_strCtc, "size=3", "terminate=newline")
			.Show
			get_user_data = .Status
			If .Status <> vbOK Then 
				Exit Function  
			End If 
			s_strTermTele = s_objStrings.RemoveNonNumericCharacters(.Responses.Item("s_strTermTele"))         
			s_strOrigTele = s_objStrings.RemoveNonNumericCharacters(.Responses.Item("s_strOrigTele"))        
			s_strCtc = s_objStrings.RemoveNonNumericCharacters(.Responses.Item("s_strCtc"))
		End With
		' return only numeric characters...   
		s_strTermTele = s_objStrings.RemoveNonNumericCharacters(s_strTermTele)         
		s_strOrigTele = s_objStrings.RemoveNonNumericCharacters(s_strOrigTele)        
		s_strCtc = s_objStrings.RemoveNonNumericCharacters(s_strCtc)

		If (s_objUtils.TestLength(s_strOrigTele, "Orig TN", 10) _
				And s_objUtils.TestLength(s_strTermTele, "Term TN", 10) _
				And s_objUtils.TestLength(s_strCtc, "CTC", 3)) Then 
			blnTryAgain = False 
		End if
	Loop While blnTryAgain
End Function 'get_user_data

'********'*********'*********'*********'*********'*********'*********'*********'
'issue_Diskut_commands
'********'*********'*********'*********'*********'*********'*********'*********'
Sub issue_Diskut_commands()
	s_objUtils.SendAndWaitFor "diskut;lf sd00perm;Read saver;quit all", ">", s_intTimeout
End Sub 'issue_Diskut_commands

'********'*********'*********'*********'*********'*********'*********'*********'
'issue_caintest_command
'********'*********'*********'*********'*********'*********'*********'*********'
sub issue_caintest_command()
	s_objUtils.SendAndWaitFor "cainjuxn " & s_strTermTele & " " & s_strOrigTele & " " & s_strCtc, ">", s_intTimeout
	s_objUtils.SendAndWaitFor "table ofcvar;pos orig_switch_id", ">", s_intTimeout
	s_objUtils.SendAndWaitFor "send response 1", ">", s_intTimeout
End Sub 'issue_caintest_command


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