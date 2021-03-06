'#$language = "VBScript"
'#$interface = "1.0"
'Name: <SCRIPT NAME>.vbs

Option Explicit

Dim s_intTimeout
Dim s_objScriptTab	'Gets set during Initialize.
Dim s_objUtils
Dim s_objDialogBox
'Dim s_objStrings
'Dim s_objArrays
'Dim s_objWatchFor
Dim s_strPort
'Dim s_strSvlan

'********'*********'*********'*********'*********'*********'*********'*********'
'Main
'********'*********'*********'*********'*********'*********'*********'*********'
Sub Main
	Dim blnRet
	 	
	If (Not Initialize) Then
		Exit Sub
	End If
	s_objUtils.ScriptStart
	
	blnRet = s_objUtils.SendAndWaitFor("", "#", s_intTimeout) 
	If (blnRet) Then 'This tests if you have access to the screen.
		If (get_user_data = vbOK) Then 
			Call issue_PORTSVLAN_command
		End If 
	Else
		s_objUtils.CancelMsgBox
	End If 	
	
	s_objUtils.ScriptEnd
	Set s_objUtils = Nothing
	Set s_objDialogBox = Nothing
'	Set s_objStrings = Nothing
'	Set s_objArrays = Nothing
'	Set s_objWatchFor = Nothing
End Sub 'Main

'********'*********'*********'*********'*********'*********'*********'*********'
'get_user_data
'********'*********'*********'*********'*********'*********'*********'*********'
Function get_user_data()
	Dim blnTryAgain
	
	blnTryAgain = True 
	Do 	
		With s_objDialogBox
			s_objDialogBox.Clear
			.Method = "HTA"
			.Height = 200
			.Width = 375
			.LabelWidths = "75px"
			.Title = "show controller tengige [port]"
			.AddItem Array("type=paragraph", "name=label1", "value= Please enter TenGige Port:", "terminate=newline")
			.AddItem Array("type=paragraph", "name=label1", "value= Example: Port 0/6/0/3", "terminate=newline")
			.AddItem Array("type=text", "name=s_strPort", "label=Enter Port:", "accesskey=p", "value=" & s_strPort, "size=10", "terminate=newline")
'			.AddItem Array("type=text", "name=s_strSvlan", "label=Enter S-VLan:", "accesskey=s", "value=" & s_strSvlan, "size=10", "terminate=newline")
			.Show
			get_user_data = .Status
			If (.Status <> vbOK) Then 
				Exit Function  
			End If 
			s_strPort = .Responses.Item("s_strPort")           
'			s_strSvlan = .Responses.Item("s_strSvlan")           
		End With

		If (s_objUtils.TestLengthMinMax(s_strPort, "Port Number", 3, 8)) then
'				And s_objUtils.TestLengthMinMax(s_strSvlan, "S-Vlan", 2, 9)) 
			blnTryAgain = False 
		End if
	Loop While blnTryAgain
End Function 'get_user_data

'********'*********'*********'*********'*********'*********'*********'*********'
'issue_PORTSVLAN_command
'********'*********'*********'*********'*********'*********'*********'*********'
Sub issue_PORTSVLAN_command()
	s_objUtils.Send "show controller tengige " & s_strPort
'	s_objUtils.SendAndWaitFor "show circuit counters " & s_strPort & " | include " & s_strSvlan, "#", s_intTimeout
End Sub 'issue_PORTSVLAN_command



'********'*********'*********'*********'*********'*********'*********'*********'
'Initialize the system by importing Constants.txt and classes, setting up 
'objects, and testing the connection.
'********'*********'*********'*********'*********'*********'*********'*********'
Function Initialize()
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
