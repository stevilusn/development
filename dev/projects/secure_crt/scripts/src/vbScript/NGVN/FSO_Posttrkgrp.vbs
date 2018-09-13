'#$language = "VBScript"
'#$interface = "1.0"
'Name: Posttrkgrp.vbs

Option Explicit

'The following get set during Initialize.
Dim s_intTimeout
Dim s_objScriptTab	
Dim s_objUtils
Dim s_objStrings
Dim s_objWatchFor

'********'*********'*********'*********'*********'*********'*********'*********'
'Main
'********'*********'*********'*********'*********'*********'*********'*********'
Sub Main
	Dim strTrkgrpNo
	Dim strCLLI
	Dim strMsg
	
	If (Not Initialize) Then Exit Sub
	s_objUtils.ScriptStart
	
	'Get strCLLI
	If (s_objUtils.Prompt("CLLI (FMS TG Number)", "Enter Switch CLLI", strCLLI) = False) Then
		ExitMain : Exit Sub 
	End If 
	
	s_objUtils.IgnoreCase = True
	If (s_objUtils.LeaveAll(s_intTimeout)) Then 'This tests if you have access to the screen.
		If (s_objUtils.IsPosInteger(strCLLI)) Then
			strTrkgrpNo = strCLLI
			strCLLI = s_objWatchFor.SearchClliCdr(s_objStrings.IntVal(strCLLI))
				If (strCLLI = "") Then
				crt.Dialog.MessageBox "Trunk group not found: " & strTrkgrpNo, "NOT FOUND", vbOKOnly + vbExclamation
				ExitMain : Exit Sub 
			End If
		Else
			Dim strCLLIStatus
			strCLLIStatus = CLLIExists(strCLLI, s_intTimeout)
			Select Case strCLLIStatus
				Case "TIMEOUT"
					'TODO:
					s_objUtils.LeaveAll(s_intTimeout)
					strMsg = "TimeOut occured for '" & strCLLI & "' while searching the screen for one of 'TUPLE NOT FOUND', 'COMMAND DISALLOWED DURING DUMP', or '>'.  Ending process." 
					crt.Dialog.MessageBox  strMsg, "TIMEOUT!", vbOKOnly + vbCritical
					ExitMain : Exit Sub
				Case "TUPLE NOT FOUND"
					s_objUtils.LeaveAll(s_intTimeout)
					crt.Dialog.MessageBox  "Trunk group not found: " & strCLLI, "NOT FOUND", vbOKOnly + vbExclamation
					ExitMain : Exit Sub
				Case "COMMAND DISALLOWED DURING DUMP"
					s_objUtils.LeaveAll(s_intTimeout)
					crt.Dialog.MessageBox "Image Dump In Progress!", "IN PROGRESS", vbOKOnly + vbExclamation
					ExitMain : Exit Sub
				Case ">"
					s_objUtils.LeaveAll(s_intTimeout)
			End Select 
		End If 
		s_objUtils.SendAndWaitFor "send sink;leave all;mapci;mtc;trks;ttp;SEND PREVIOUS;post g " & strCLLI, ">", s_intTimeout
		s_objUtils.SendAndWaitFor "", ">", s_intTimeout
	Else 
		s_objUtils.CancelMsgBox
	End If

	ExitMain
End Sub 'Main
Sub ExitMain()
	s_objUtils.ScriptEnd
	Set s_objUtils = Nothing
	Set s_objStrings = Nothing
	Set s_objWatchFor = Nothing
	Set s_objScriptTab = Nothing 
End Sub 'ExitMain


'********'*********'*********'*********'*********'*********'*********'*********'
'CLLIExists
'********'*********'*********'*********'*********'*********'*********'*********'
Function CLLIExists(strClli, intTimeout) 'returns string
	Dim strWork
	Dim intIndex
	Dim arrstrSearch
	
	strWork = Trim(strClli)

	s_objUtils.LeaveAll(s_intTimeout)
	s_objUtils.SendAndWaitFor "send sink;table clli;format pack;send previous", ">", s_intTimeout
	arrstrSearch = Array("TUPLE NOT FOUND", "COMMAND DISALLOWED DURING DUMP", ">")
	intIndex = s_objUtils.SendAndWaitFor("pos " & strWork, arrstrSearch, s_intTimeout)
	If (intIndex = 0) Then
		CLLIExists = "TIMEOUT"
	Else
		CLLIExists = arrstrSearch(intIndex - 1)
	End If
End Function


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
			s_intTimeout = DEFAULT_TIMEOUT
		End If
	Else
		MsgBox "Could not find file " & strCurPath & "BootStrap.txt.", vbOK & vbCritical, "FILE NOT FOUND!"
	End If 
	Set objFSO = Nothing
End Function 'Initialize

