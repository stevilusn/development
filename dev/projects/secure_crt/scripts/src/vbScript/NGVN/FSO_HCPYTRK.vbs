'#$language = "VBScript"
'#$interface = "1.0"
'Name: FSO's HCPYTRK.vbs

Option Explicit

Dim s_intTimeout
Dim s_objScriptTab	'Gets set during Initialize.
Dim s_objUtils
Dim s_objStrings
Dim s_objWatchFor

Const s_intIMAGEINPROGRESS = -1	
Const s_intCLLINOTFOUND = -2

crt.Screen.Synchronous = True 

'********'*********'*********'*********'*********'*********'*********'*********'
'Main
'********'*********'*********'*********'*********'*********'*********'*********'
Sub Main
	Dim intError
	Dim intCount
	Dim strCLLI
	Dim strCLLINum
 	
	If (Not Initialize) Then Exit Sub
	s_objUtils.ScriptStart
	s_objScriptTab.Screen.Synchronous = True 
			
	If (s_objUtils.Prompt("NOTE: If you know the letter designation, go ahead and enter it." & vbcrlf & _
						"Otherwise enter just the numerical portion and the script will" & vbcrlf & _
						"search for it in table CLLICDR." _
						& vbcrlf & vbcrlf & "CLLI (FMS TG Number)", "Enter Switch CLLI", strCLLI) = False) Then
		ExitMain : Exit Sub 
	End If 
	If (s_objUtils.LeaveAll(s_intTimeout)) Then 'This tests if you have access to the screen.
		intCount = 0
		intError = 0
		Do
			intCount = intCount + 1
			intError = fnHcpytrk(strCLLI)
			If (intError = s_intCLLINOTFOUND And intCount < 2) Then
				If (Not s_objUtils.IsPosInteger(strCLLI)) Then 
					strCLLINum = s_objStrings.RemoveLeftNonNumericCharacters(strCLLI)
					If (s_objUtils.Prompt("CLLI """ & strCLLI & """ was not found.  Would you like to search for the CLLI by ", "Enter Switch CLLI", strCLLINum) = True) Then
						strCLLI = strCLLINum
						If (Not s_objUtils.IsPosInteger(strCLLI)) Then
							intCount = 0
							s_objUtils.LeaveAll(s_intTimeout)
						End If 
					Else 
						ExitMain : Exit Sub 
					End If 
				End If 
				If (s_objUtils.IsPosInteger(strCLLI)) Then 
					strCLLI = s_objWatchFor.SearchClliCdr(s_objStrings.IntVal(strCLLI))
					If (strCLLI = "") Then
						intCount = 2
					End If 
				End If
			End If
		Loop Until (intError = 0 Or intError = s_intIMAGEINPROGRESS Or intCount >= 2)
	
		Select Case intError 
			Case s_intIMAGEINPROGRESS
				crt.Dialog.MessageBox "Image in progress, please try later.", "IN PROGRESS", vbOKOnly & vbExclamation
				s_objUtils.SendAndWaitFor "", ">", s_intTimeout
				s_objUtils.Send "date"
				ExitMain : Exit Sub
			Case s_intCLLINOTFOUND
				crt.Dialog.MessageBox "CLLI not found", "NOT FOUND", vbOKOnly & vbExclamation
				s_objUtils.SendAndWaitFor "", ">", s_intTimeout
				ExitMain : Exit Sub
			Case -1
				ExitMain : Exit Sub 
		End Select
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
'fnHcpytrk
'********'*********'*********'*********'*********'*********'*********'*********'
Private Function fnHcpytrk(strTrkGrp) 'returns integer
	Dim intResult
	
	intResult = s_objUtils.SendAndWaitFor("MAPCI NODISP;MTC;TRKS;STAT;SELGRP " & strTrkGrp, Array("COMMAND DISALLOWED DURING DUMP", "CLLI NAME NOT FOUND", ">"), s_intTimeout)
	Select Case intResult
	    Case 0
	    	crt.Dialog.MessageBox "Command timed out waiting for expected response (""COMMAND DISALLOWED DURING DUMP"", ""CLLI NAME NOT FOUND"", or "">"").", "TIMEOUT", vbOKOnly & vbExclamation
	    	intResult = -1
	    Case 1 '"COMMAND DISALLOWED DURING DUMP"
	    	intResult = s_intIMAGEINPROGRESS
	    	s_objUtils.WaitFor ">", s_intTimeout
	    Case 2 '"CLLI NAME NOT FOUND"
	    	intResult = s_intCLLINOTFOUND
	    	s_objUtils.WaitFor ">", s_intTimeout
	    Case 3 '">"
	    	s_objUtils.SendAndWaitFor "hcpytrk", ">", s_intTimeout
	    	s_objUtils.LeaveAll(s_intTimeout)
	    	intResult = 0
	    Case Else
	    	crt.Dialog.MessageBox "Unexpected response while waiting for (""COMMAND DISALLOWED DURING DUMP"", ""CLLI NAME NOT FOUND"", or "">"").", "TIMEOUT", vbOKOnly & vbExclamation
	    	intResult = -1
	End Select
	fnHcpytrk = intResult 
End Function 'fnHcpytrk


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
