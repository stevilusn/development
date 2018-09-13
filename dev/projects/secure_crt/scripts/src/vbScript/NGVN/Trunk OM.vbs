'#$language = "VBScript"
'#$interface = "1.0"
'Name: Trunk OM.vbs

Option Explicit

Dim s_objScriptTab
Dim s_objUtils
Dim s_intTimeout

Dim s_strClli

'********'*********'*********'*********'*********'*********'*********'*********'
'Main
'********'*********'*********'*********'*********'*********'*********'*********'
Sub Main
	Dim blnRepeat
	Dim blnRet
	
	If (Not Initialize) Then
		Exit Sub
	End If
	s_objUtils.ScriptStart

	blnRet = s_objUtils.SendAndWaitFor("leave all", ">", s_intTimeout) 
	If (blnRet) Then 'This tests if you have access to the screen.
		s_objUtils.Message "Running: get_user_data"
		If (get_user_data() = vbOK) Then 
			s_objUtils.Message "Running: issue_trunkom_commands"
			Call issue_trunkom_commands
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
	s_strClli = crt.Dialog.Prompt("Issue OMSHOW TRK commands for " & WeekdayName(Weekday(Now())) & "." _
			& vbCrLf & vbCrLf & "Enter Trunk Group:", "NGVN TRUNK OMSHOW", s_strClli)
	If (Len(s_strClli) > 0) Then
		get_user_data = vbOK 
	Else
		get_user_data = vbCancel 		
	End If
End Function 'get_user_data

'********'*********'*********'*********'*********'*********'*********'*********'
'issue_trunkom_commands
'********'*********'*********'*********'*********'*********'*********'*********'
Sub issue_trunkom_commands()
	s_objUtils.SendAndWaitFor "leave all", ">", s_intTimeout
	
	issue_omshowtrk_command "ACTIVE" & " " & s_strClli
	issue_omshowtrk_command "HOLDING" & " " & s_strClli
	Select Case (Weekday(Now()))
	    Case vbSunday
	    	issue_omshowtrk_command "SAT" & " " & s_strClli
	    	issue_omshowtrk_command "SUN" & " " & s_strClli
	    Case vbMonday
	    	issue_omshowtrk_command "SUN" & " " & s_strClli
	    	issue_omshowtrk_command "MON" & " " & s_strClli
	    Case vbTuesday
	    	issue_omshowtrk_command "MON" & " " & s_strClli
	    	issue_omshowtrk_command "TUE" & " " & s_strClli
	    Case vbWednesday
	    	issue_omshowtrk_command "TUE" & " " & s_strClli
	    	issue_omshowtrk_command "WED" & " " & s_strClli
	    Case vbThursday
	    	issue_omshowtrk_command "WED" & " " & s_strClli
	    	issue_omshowtrk_command "THU" & " " & s_strClli
	    Case vbFriday
	    	issue_omshowtrk_command "THU" & " " & s_strClli
	    	issue_omshowtrk_command "FRI" & " " & s_strClli
	    Case vbSaturday
	    	issue_omshowtrk_command "FRI" & " " & s_strClli
	    	issue_omshowtrk_command "SAT" & " " & s_strClli
	End Select
End Sub 'issue_trunkom_commands

'********'*********'*********'*********'*********'*********'*********'*********'
'issue_omshowtrk_command
'********'*********'*********'*********'*********'*********'*********'*********'
Sub issue_omshowtrk_command(strCommand)
	s_objUtils.SendAndWaitFor "OMSHOW TRK " & strCommand, ">", s_intTimeout
	s_objUtils.SendAndWaitFor "", ">", s_intTimeout
End Sub 'issue_omshowtrk_command


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