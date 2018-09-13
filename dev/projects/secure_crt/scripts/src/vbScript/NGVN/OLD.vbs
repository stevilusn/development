'#$language = "VBScript"
'#$interface = "1.0"
'Name: OLD.vbs

Option Explicit

Dim s_objScriptTab
Dim s_objUtils
Dim s_objDialogBox
Dim s_intTimeout

Dim s_strOrig
Dim s_strTnum	' to store user input
Dim s_strLen	' to store user input

'********'*********'*********'*********'*********'*********'*********'*********'
'
'********'*********'*********'*********'*********'*********'*********'*********'
Sub Main
	Dim blnRet
	Dim intRet
	
	If (Not Initialize) Then
		Exit Sub
	End If
	s_objUtils.ScriptStart

	blnRet = s_objUtils.SendAndWaitFor("leave all", ">", s_intTimeout) 
	If (blnRet) Then 'This tests if you have access to the screen.
		intRet = crt.Dialog.MessageBox("This script will remove the TN from the switch.  Use with caution! ! ! ", "CAUTION", vbExclamation + vbOKCancel)
		If (intRet = vbOK) Then 
			If (get_user_data() = vbOK) then
				Call OUT_COMMAND
			End If
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
Function get_user_data
	Dim strTN
	Dim blnTryAgain

	s_objUtils.Message "Running: get_user_data"

	blnTryAgain = True 
	Do 	
		With s_objDialogBox
			.Clear
			.Method = "HTA"		
			.Height = 150
			.Width = 350
			.LabelWidths = "75px"
			.Title = "SERVORD OUT Command"
			.AddItem Array("type=text", "name=s_strTnum", "label=TN:", "accesskey=T",  "value=" & s_strTnum, "size=10", "terminate=newline")
			.AddItem Array("type=text", "name=s_strLen",  "label=LEN:", "accesskey=L", "value=" & s_strLen,  "size=10", "terminate=newline")
			.Show
			get_user_data = .Status
			If (.Status <> vbOK) Then 
				Exit Function  
			End If 
			s_strTnum = .Responses.Item("s_strTnum")           
			s_strLen = .Responses.Item("s_strLen")           
		End With
					
		'clean the spaces, underlines and dashes from the orig and term teles...   
		s_strTnum = s_objUtils.stripTN(s_strTnum)
		s_strOrig = s_strTnum  
		
		if (s_objUtils.TestLength(s_strOrig, "TN", 10)) then 
			blnTryAgain = False 
		End if
	Loop While blnTryAgain
	get_user_data = vbOK
End Function 'get_user_data

'********'*********'*********'*********'*********'*********'*********'*********'
'OUT_COMMAND
'********'*********'*********'*********'*********'*********'*********'*********'
Sub OUT_COMMAND
	s_objUtils.Send "SERVORD;OUT $ " & s_strOrig & " " & s_strLen & " bldn y "
End Sub 'OUT_COMMAND


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