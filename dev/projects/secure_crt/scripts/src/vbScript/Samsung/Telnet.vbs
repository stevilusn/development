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
Dim s_strSvlan
Dim s_strCascade

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
			.Height = 225
			.Width = 375
			.LabelWidths = "100px"
			.Title = "Telnet"
			.AddItem Array("type=paragraph", "name=label1", "value= Please enter exact requested information:", "terminate=newline")
			.AddItem Array("type=paragraph", "name=label1", "value= Example: IP Address: 113.9.58.226", "terminate=newline")
			.AddItem Array("type=paragraph", "name=label1", "value= Example: Port: 0/6/0/3.1740", "terminate=newline")
'			.AddItem Array("type=text", "name=s_strCascade", "label=Enter Cascade ID:", "accesskey=p", "value=" & s_strCascade, "size=9", "terminate=newline")
			.AddItem Array("type=text", "name=s_strSvlan", "label=Enter IP Address:", "accesskey=s", "value=" & s_strSvlan, "size=10", "terminate=newline")
			.AddItem Array("type=text", "name=s_strPort", "label=Enter Port:", "accesskey=p", "value=" & s_strPort, "size=10", "terminate=newline")
			.Show
			get_user_data = .Status
			If (.Status <> vbOK) Then 
				Exit Function  
			End If 
'			s_strCascade = .Responses.Item("s_strCascade")   
			s_strPort = .Responses.Item("s_strPort")           
			s_strSvlan = .Responses.Item("s_strSvlan")           
		End With

		If (s_objUtils.TestLengthMinMax(s_strSvlan, "IP Address", 7, 15)) then
'				And s_objUtils.TestLengthMinMax(s_strPort, "Port", 2, 6))
'				(s_objUtils.TestLengthMinMax(s_strSvlan, "S-Vlan", 2, 9)) then
			blnTryAgain = False 
		End if
	Loop While blnTryAgain
End Function 'get_user_data

'********'*********'*********'*********'*********'*********'*********'*********'
'issue_PORTSVLAN_command
'********'*********'*********'*********'*********'*********'*********'*********'
Sub issue_PORTSVLAN_command()
	s_objUtils.Send "telnet vrf sat-exdmz " & s_strSvlan & " source Te" & s_strPort & "" & chr(13)
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
