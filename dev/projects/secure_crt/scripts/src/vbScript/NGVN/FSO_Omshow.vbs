'#$language = "VBScript"
'#$interface = "1.0"
'Name: Omshow.vbs

Option Explicit

Dim s_intTimeout
Dim s_objScriptTab	'Gets set during Initialize.
Dim s_objUtils
Dim s_objStrings
Dim s_objWatchFor

'********'*********'*********'*********'*********'*********'*********'*********'
'Main
'********'*********'*********'*********'*********'*********'*********'*********'
Sub Main
	Dim strCLLI
 	
	If (Not Initialize) Then Exit Sub
	s_objUtils.ScriptStart
	
	If (s_objUtils.LeaveAll(s_intTimeout)) Then 'This tests if you have access to the screen.
		If (s_objUtils.Prompt("CLLI (FMS TG Number)", "Enter Switch CLLI", strCLLI) = False) Then
			ExitMain : Exit Sub
		End If 
		
		If (s_objUtils.IsPosInteger(strCLLI)) Then
			strCLLI = s_objWatchFor.SearchClliCdr(s_objStrings.IntVal(strCLLI))
			If (strCLLI = "") Then 
				crt.Dialog.MessageBox "Trunk group not found", "NOT FOUND", vbOKOnly + vbExclamation
				ExitMain : Exit Sub
			End If 
		End If 
	
		s_objUtils.SendAndWaitFor "OMSHOW TRK SWDAY  " & strCLLI, ">", s_intTimeout 
		s_objUtils.SendAndWaitFor "", ">", s_intTimeout 
	
		s_objUtils.SendAndWaitFor "OMSHOW TRK ACTIVE  " & strCLLI, ">", s_intTimeout 
		s_objUtils.SendAndWaitFor "", ">", s_intTimeout 
	
		s_objUtils.SendAndWaitFor "OMSHOW TRK HOLDING  " & strCLLI, ">", s_intTimeout 
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
