'#$language = "VBScript"
'#$interface = "1.0"
'Name: NEW.vbs

Option Explicit

Dim s_objScriptTab
Dim s_objUtils
Dim s_objDialogBox
Dim s_intTimeout

Dim s_strLen	' to store user input
Dim s_strLinetype	' to store user input
Dim s_strLataname	' to store user input
Dim s_strLtg	' to store user input
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
		If (get_user_data() = vbOK) Then 
			Call NEW_COMMAND
		End If 
	Else
		s_objUtils.CancelMsgBox
	End If 	

	s_objUtils.ScriptEnd
	Set s_objDialogBox = Nothing 
	Set s_objUtils = Nothing
End Sub 'Main

'********'*********'*********'*********'*********'*********'*********'*********'
'get_user_data
'********'*********'*********'*********'*********'*********'*********'*********'
Function get_user_data()
	Dim strTN
	Dim blnTryAgain
	
	strTN = ""
	s_strLinetype = "1FR"
	s_strLataname = ""
	s_strLtg = ""
	s_strLen = ""
	blnTryAgain = True 
	Do 	
		With s_objDialogBox
			.Clear
			.Method = "HTA"		
			.Height = 200
			.Width = 375
			.LabelWidths = "100px"
			.Title = "OSSGATE NEW Command"
			.AddItem Array("type=text", "name=strTN", "label=TN:", "accesskey=T", "value=" & strTN, "size=10", "terminate=newline")
			.AddItem Array("type=text", "name=s_strLinetype", "label=LINE TYPE:", "accesskey=L", "value=" & s_strLinetype, "size=10", "terminate=newline")
			.AddItem Array("type=text", "name=s_strLataname", "label=LATA NAME:", "accesskey=T", "value=" & s_strLataname, "size=10", "terminate=newline")
			.AddItem Array("type=text", "name=s_strLtg", "label=LTG:", "accesskey=T", "value=" & s_strLtg, "size=10", "terminate=newline")
			.AddItem Array("type=text", "name=s_strLen", "label=LEN:", "accesskey=T", "value=" & s_strLen, "size=10", "terminate=newline")
			.Show
			get_user_data = .Status
			If (.Status <> vbOK) Then 
				Exit Function 
			End If 
			strTN = .Responses.Item("strTN")           
			s_strLinetype = .Responses.Item("s_strLinetype")           
			s_strLataname = .Responses.Item("s_strLataname") 
			s_strLtg = .Responses.Item("s_strLtg") 
			s_strLen = .Responses.Item("s_strLen") 
		End With

		'clean the spaces, underlines and dashes from the orig and term teles...   
		s_strOrig=s_objUtils.stripTN(strTN)  
		
		If (s_objUtils.TestLength(s_strOrig, "TN", 10)) Then 
			blnTryAgain = false
		End if
	Loop While blnTryAgain
	get_user_data = vbOK
End Function 'get_user_data


'********'*********'*********'*********'*********'*********'*********'*********'
'NEW_COMMAND
'********'*********'*********'*********'*********'*********'*********'*********'
Sub NEW_COMMAND()
	s_objUtils.Send "SERVORD; NEW $ " & s_strOrig & " " & s_strLinetype & " " & s_strLataname & " " & s_strLtg & " " & s_strLen & " dgt $"
End Sub 'NEW_COMMAND


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