'#$language = "VBScript"
'#$interface = "1.0"
'Name: clsTSTLNOIND.vbs

Option Explicit

Dim s_intTimeout
Dim s_objScriptTab	'Gets set during Initialize.
Dim s_objUtils

'********'*********'*********'*********'*********'*********'*********'*********'
'Main
'********'*********'*********'*********'*********'*********'*********'*********'
Sub Main
	Dim strTstnoind 
	Dim strMsg 
	
	If (Not Initialize) Then Exit Sub
	s_objUtils.ScriptStart
	
	s_objUtils.DebugLevel = 1	

	If (s_objUtils.QuitAll(s_intTimeout)) Then 'This tests if you have access to the screen.
		s_objUtils.SendAndWaitFor "", ">", s_intTimeout 
		
		strMsg = strMsg & "Use this Button to find out which TSTNOIND are not being used.  " 
		strMsg = strMsg & "It will display Table Cllimtce and all of its fields.  Below " 
		strMsg = strMsg & "the field ""TSTNOIND"" there is a number populated stating which " 
		strMsg = strMsg & "TSTLCONT it is using.  For example, if you choose to start with " 
		strMsg = strMsg & "a TSTNOIND 25, but do not see the Table display until 27 or " 
		strMsg = strMsg & "you see ""TOP/BOTTOM"" this means the TSTNOIND 25 & 26 are " 
		strMsg = strMsg & "available.  Remember, that when you see ""TOP/BOTTOM"" displayed, " 
		strMsg = strMsg & "it means that particular TSTNOIND is available." 
		strMsg = strMsg & vbCrLf & vbCrLf 
		strMsg = strMsg & "Press Ok to Begin..." 
		If (crt.Dialog.MessageBox(strMsg, "TSTNOIND", vbOKCancel + vbInformation) <> vbOk) then 
			ExitMain : Exit Sub
		End If 	
'This next line was in the Crosstalk code but it's usage could not be determined.
'		s_objUtils.WaitFor "|", s_intTimeout 
	
		s_objUtils.SendAndWaitFor "", ">", s_intTimeout 
		If (Not s_objUtils.Prompt("Enter a TSTNOIND to start with (1-100): ", "ENTER TSTNOIND", strTstnoind)) Then
			ExitMain : Exit Sub
		End If

		s_objUtils.SendAndWaitFor "", ">", s_intTimeout 
		s_objUtils.SendAndWaitFor "table cllimtce", ">", s_intTimeout 
		s_objUtils.Send strTstnoind & "->x"
		s_objUtils.SendAndWaitFor "repeat 5(lis all(7 eq x);x+1->x)", ">", s_intTimeout 
		s_objUtils.QuitAll(s_intTimeout)
	Else 
		s_objUtils.CancelMsgBox
	End If
	
	ExitMain
End Sub 'Main
Sub ExitMain()
	s_objUtils.ScriptEnd
	Set s_objUtils = Nothing
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
