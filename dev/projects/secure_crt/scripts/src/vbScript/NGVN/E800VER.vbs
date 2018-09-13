'#$language = "VBScript"
'#$interface = "1.0"
'Name: E800VER.vbs

Option Explicit

Dim s_objScriptTab
Dim s_objUtils
Dim s_objDialogBox
Dim s_objStrings
Dim s_intTimeout

Dim s_strOrigTele	' to store user input
Dim s_strTermTele	' to store user input
Dim s_strOrigLata	' to store user input
Dim s_strTerm
Dim s_strTermOne	'TODO: This is never set.  What was it's purpose?  Can it be removed?

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
			Call e800ver
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
			.Title = "E800VER"
			.AddItem Array("type=paragraph", "name=label1", "value=" & s_objUtils.CleanTrashMsg("input"), "terminate=newline")
			.AddItem Array("type=text", "name=s_strOrigTele", "label=Orig TN:", "accesskey=r", "value=" & s_strOrigTele, "size=10", "terminate=newline")
			.AddItem Array("type=text", "name=s_strTermTele", "label=Term TN:", "accesskey=e", "value=" & s_strTermTele, "size=10", "terminate=newline")
			.AddItem Array("type=text", "name=s_strOrigLata", "label=Orig LATA:", "accesskey=L","value=" & s_strOrigLata, "size=3", "terminate=newline")
			.Show
			get_user_data = .Status
			If (.Status <> vbOK) Then 
				Exit Function  
			End If 
			s_strTermTele = .Responses.Item("s_strTermTele")           
			s_strOrigTele = .Responses.Item("s_strOrigTele")           
			s_strOrigLata = .Responses.Item("s_strOrigLata") 
		End With
		' clean the spaces, underlines and dashes from the orig and term teles...   
		s_strTermTele = s_objStrings.Strip(s_objStrings.RemoveNonNumericCharacters(s_strTermTele), "1", 2)
		s_strOrigTele = s_objStrings.RemoveNonNumericCharacters(s_strOrigTele)
		s_strOrigLata = s_objStrings.RemoveNonNumericCharacters(s_strOrigLata)

		If (s_objUtils.TestLength(s_strOrigTele, "Term TN", 10) _
				And s_objUtils.TestLength(s_strTermTele, "Orig TN", 10) _
				And s_objUtils.TestLength(s_strOrigLata, "Orig LATA", 3)) Then 
			blnTryAgain = False 
		End if
	Loop While blnTryAgain
End Function 'get_user_data

'********'*********'*********'*********'*********'*********'*********'*********'
'e800ver
'********'*********'*********'*********'*********'*********'*********'*********'
Sub e800ver()
	s_objUtils.SendAndWaitFor "e800ver " & s_strOrigTele & " " & s_strOrigLata & " " & s_strTermTele, ">", s_intTimeout
End Sub 'e800ver


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