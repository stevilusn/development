'#$language = "VBScript"
'#$interface = "1.0"
'Name: QDN & QLRN.vbs

Option Explicit

Dim s_objScriptTab
Dim s_objUtils
Dim s_objDialogBox
Dim s_objStrings
Dim s_intTimeout

Dim s_strO_Tele	' to store user input
Dim s_blnQlrnonly

'********'*********'*********'*********'*********'*********'*********'*********'
'Main
'********'*********'*********'*********'*********'*********'*********'*********'
Sub Main
	Dim blnRet
	Dim arrScreenData
	Dim strNxx
	
	If (Not Initialize) Then
		Exit Sub
	End If
	s_objUtils.ScriptStart
	
	blnRet = s_objUtils.SendAndWaitFor("leave all", ">", s_intTimeout) 
	If (blnRet) Then 'This tests if you have access to the screen.
		s_blnQlrnonly = False
		If (get_user_data() = vbOK) Then 
			If (s_blnQlrnonly) Then 
				s_objUtils.SendAndWaitFor "qlrn " & s_strO_Tele, ">", s_intTimeout
			Else 
				s_objUtils.SendAndWaitFor "QDN  " & s_strO_Tele, ">", s_intTimeout
				arrScreenData = s_objUtils.SendAndReadStrings("qlrn " _
						& s_strO_Tele, ">", s_intTimeout)
				s_objUtils.SendAndWaitFor "", ">", s_intTimeout
				strNxx = parse_Nxx(arrScreenData)
				HomeLRN strNxx
			End If 
			s_objUtils.WriteMwqDn s_strO_Tele
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
			.Height = 230
			.Width = 375
			.LabelWidths = "100px"
			.Title = "QDN and QLRN Check"
			.AddItem Array("type=paragraph", "name=label1", "value=" _
					& s_objUtils.CleanTrashMsg("DN") _
					& "<br><br>When you check QLRN only the HomeLRN command will not be issued. ")
			.AddItem Array("type=text", "name=s_strO_Tele", "label=DN:" , "accesskey=D"_
					, "value="& s_strO_Tele, "size=10", "terminate=newline")
			.AddItem Array("type=checkbox", "name=s_blnQlrnonly", "label=QLRN Only:", "value=1", "checked=" _
					& s_blnQlrnonly, "accesskey=Q", "terminate=newline")
			.Show
			get_user_data = .Status
			If (.Status <> vbOK) Then 
				Exit Function  
			End If 
			s_strO_Tele = .Responses.Item("s_strO_Tele")
			s_blnQlrnonly = .Responses.Item("s_blnQlrnonly")
		End With
		s_strO_Tele = s_objStrings.RemoveNonNumericCharacters(s_strO_Tele)  
		If (s_objUtils.TestLength(s_strO_Tele, "DN", 10)) Then 
			blnTryAgain = False
		End If 
	Loop While blnTryAgain
End Function 'get_user_data

'********'*********'*********'*********'*********'*********'*********'*********'
'parse_data
'********'*********'*********'*********'*********'*********'*********'*********'
Function parse_Nxx(arrData)
	Dim strData
	Dim strNxx 
	
	For Each strData In arrData
		'determine if number is ported   
		If (InStr(strData, "Routing number:") > 0) Then
			strNxx = Mid(strData, InStr(strData, ":") + 5, 3)
			'TODO: Exit the For?  Or get the last NXX in the array?
		End If
	Next
	parse_Nxx = strNxx
End Function 'parse_Nxx

'********'*********'*********'*********'*********'*********'*********'*********'
'HomeLRN
'********'*********'*********'*********'*********'*********'*********'*********'
Sub HomeLRN(strNxx)
	s_objUtils.SendAndWaitFor "Table homelrn ; lis all (2 eq " _
			& strNxx & " )", ">", s_intTimeout
	s_objUtils.SendAndWaitFor "leave all", ">", s_intTimeout
End Sub 'HomeLRN


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