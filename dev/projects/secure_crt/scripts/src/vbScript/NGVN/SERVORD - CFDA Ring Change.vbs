'#$language = "VBScript"
'#$interface = "1.0"
'Name: SERVORD - CFDA Ring Change.vbs

Option Explicit

Dim s_objScriptTab
Dim s_objUtils
Dim s_objDialogBox
'Dim s_objStrings
'Dim s_objArrays
Dim s_intTimeout

Dim s_strOrig
Dim s_strO_Tele
Dim s_intMdcm
Dim s_intMsln
Dim s_intWave
Dim s_intWehco
Dim s_intAntm
Dim s_intMlnm
Dim s_intSelco
Dim s_intBaja
Dim s_strPicvalue

Dim s_strSeconds

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
		If (get_dn_data() = vbOK) Then 
			Call issue_qdn_command
			If (get_ring_change_info = vbOk) Then
				Call chg_cfda
				Call issue_qdn_command
			End If
		End if
	Else
		s_objUtils.CancelMsgBox
	End If 
	
	s_objUtils.ScriptEnd
	Set s_objDialogBox = Nothing
	'Set s_objStrings = Nothing
	'Set s_objArrays = Nothing
	Set s_objUtils = Nothing 
End Sub 'Main

'********'*********'*********'*********'*********'*********'*********'*********'
'get_dn_data
'********'*********'*********'*********'*********'*********'*********'*********'
Function get_dn_data()
	Dim blnTryAgain
	Dim strReturn
	Dim strInfodigs	'TODO: This is never set.  Can it be removed.

	blnTryAgain = True 
	Do 
		With s_objDialogBox
			.Clear
			.Height = 175
			.Width = 350
			.LabelWidths = "50px"
			.Title = "Enter DN"
			.AddItem Array("type=paragraph", "name=label1", "value=" & s_objUtils.CleanTrashMsg("DN"), "terminate=newline")
			.AddItem Array("type=text", "name=s_strO_Tele", "label=DN:", "value=" & s_strO_Tele, "size=13", "accesskey=D", "terminate=newline")
			.Show
			get_dn_data = .Status
			If .Status <> vbOK Then 
				Exit Function  
			End If 
			s_strO_Tele = .Responses.Item("s_strO_Tele")           
		End With
		s_strO_Tele = s_objUtils.stripTN(s_strO_Tele)  
		s_strOrig = s_strO_Tele
		If (s_objUtils.TestLength(s_strOrig, "DN", 10)) Then 
			blnTryAgain = false
		End If 
	Loop While blnTryAgain

End Function 'get_dn_data

'********'*********'*********'*********'*********'*********'*********'*********'
'issue_qdn_command
'********'*********'*********'*********'*********'*********'*********'*********'
Sub issue_qdn_command()
	s_objUtils.SendAndWaitFor "leave all", ">", s_intTimeout
	s_objUtils.Send "QDN " & s_strOrig
End Sub 'issue_qdn_command

'********'*********'*********'*********'*********'*********'*********'*********'
'get_ring_change_info
'********'*********'*********'*********'*********'*********'*********'*********'
Function get_ring_change_info()
	Dim blnTryAgain
	Dim strSeconds
	Dim strReturn

	s_objUtils.SendAndWaitFor "leave all", ">", s_intTimeout

	blnTryAgain = True 
	Do 
		With s_objDialogBox
			.Clear
			.Height = 330
			.Width = 350
			.LabelWidths = "50px"
			.Title = "Voicemail Number of Rings Change"
			.AddItem Array("type=paragraph", "name=label1", "value=Select the number of Rings you want to change CFDA to.", "terminate=newline")
			.AddItem Array("type=html", "Select Ring Value:")
			.AddItem Array("type=radio", "name=strSeconds", "label=", "value=1", "accesskey=S", _
					"choices=2 Rings - 12 Seconds|3 Rings - 18 Seconds|4 Rings - 24 Seconds|5 Rings - 30 Seconds|6 Rings - 36 Seconds|7 Rings - 42 Seconds|8 Rings - 48 Seconds|9 Rings - 54 Seconds|10 Rings - 60 Seconds", _
					"termination=newline")
			.Show
			get_ring_change_info = .Status
			If .Status <> vbOK Then 
				Exit Function  
			End If 
			strSeconds = .Responses.Item("strSeconds")           
		End With

		blnTryAgain = False 
		Select Case strSeconds
		    Case "2 Rings - 12 Seconds"
		    	s_strSeconds = "12"
		    Case "3 Rings - 18 Seconds"
		    	s_strSeconds = "18"
		    Case "4 Rings - 24 Seconds"
		    	s_strSeconds = "24"
		    Case "5 Rings - 30 Seconds"
		    	s_strSeconds = "30"
		    Case "6 Rings - 36 Seconds"
		    	s_strSeconds = "36"
		    Case "7 Rings - 42 Seconds"
		    	s_strSeconds = "42"
		    Case "8 Rings - 48 Seconds"
		    	s_strSeconds = "48"
		    Case "9 Rings - 54 Seconds"
		    	s_strSeconds = "54"
		    Case "10 Rings - 60 Seconds"
		    	s_strSeconds = "60"
		    Case Else
		    	crt.Dialog.MessageBox "please select a ring value", vbOKOnly
		    	blnTryAgain = True 
		End Select
	Loop While blnTryAgain
End Function 'get_ring_change_info

'********'*********'*********'*********'*********'*********'*********'*********'
'chg_cfda
'********'*********'*********'*********'*********'*********'*********'*********'
Sub chg_cfda()
	s_objUtils.SendAndWaitFor "Servord; chf $ "  & s_strOrig & " CFDA", ">", s_intTimeout
	s_objUtils.SendAndWaitFor "", ">", s_intTimeout
	s_objUtils.SendAndWaitFor "", ">", s_intTimeout
	s_objUtils.SendAndWaitFor "",">", s_intTimeout
	s_objUtils.SendAndWaitFor s_strSeconds, ">", s_intTimeout
	s_objUtils.SendAndWaitFor "", ">", s_intTimeout
	s_objUtils.SendAndWaitFor "$", ">", s_intTimeout
	s_objUtils.SendAndWaitFor "$", ">", s_intTimeout
	s_objUtils.SendAndWaitFor "Y", ">", s_intTimeout
	s_objUtils.SendAndWaitFor "leave all", ">", s_intTimeout
End Sub 'chg_cfda

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