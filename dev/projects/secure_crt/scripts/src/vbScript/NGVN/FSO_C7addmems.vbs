'#$language = "VBScript"
'#$interface = "1.0"
'Name: C7addmems.vbs

Option Explicit

Dim s_intTimeout
Dim s_objScriptTab	'Gets set during Initialize.
Dim s_objUtils
Dim s_objDialogBox
Dim s_objStrings
Dim s_objWatchFor

'********'*********'*********'*********'*********'*********'*********'*********'
'Main
'********'*********'*********'*********'*********'*********'*********'*********'
Sub Main
	Dim intError
	Dim intCount
	Dim blnDone
	Dim blnFound
	Dim intReturn
	Dim dicData
	Dim strCLLI
	
	If (Not Initialize) Then Exit Sub
	s_objUtils.ScriptStart

	Set dicData = CreateObject("scripting.dictionary")
	
	intCount = 0
	intError = 0

	blnDone = False
	blnFound = True

	If (s_objUtils.LeaveAll(10)) Then 'This tests if you have access to the screen.
		Do 
			If (C7AddMemDialog(dicData) = vbOk) Then
				'Build the C7 mems
				s_objUtils.LeaveAll(s_intTimeout)
				If (C7AddMemEntriesValid(dicData)) Then
					s_objUtils.QuitAll(4)
					s_objUtils.SendAndWaitFor "", ">", 4
					If (s_objUtils.IsPosInteger(dicData.Item("CLLI"))) Then
						blnFound = False
						strCLLI = s_objWatchFor.SearchClliCdr(s_objStrings.IntVal(dicData.Item("CLLI")))
						If (strCLLI = "") Then 
							crt.Dialog.MessageBox "FMS Trunk group not found in table cllicdr", "NOT FOUND", vbOKOnly + vbExclamation
						Else
							dicData.Item("CLLI") = strCLLI
							blnFound = True
						End If
					End If
					If (blnFound) Then
						blnDone = True
						intReturn = s_objUtils.SendAndWaitFor("SEND SINK;TABLE C7TRKMEM;FORMAT PACK;SEND PREVIOUS;VER OFF", Array("COMMAND DISALLOWED DURING DUMP", ">"), 60)
						Select Case intReturn
						    Case 0
								crt.Dialog.MessageBox "Timed out while watching for strings ""COMMAND DISALLOWED DURING DUMP"", "">""", "TIMEOUT", vbOKOnly & vbExclamation
								s_objUtils.Send ""
						    Case 1
								s_objUtils.WaitFor ">", s_intTimeout
								crt.Dialog.MessageBox "SWITCH IS CURRENTLY IN IMAGE, PLEASE TRY LATER", "SWITCH IN IMAGE", vbOKOnly + vbExclamation
								s_objUtils.SendAndWaitFor "", ">", s_intTimeout
								s_objUtils.SendAndWaitFor "table imgsched;lis all", ">", s_intTimeout
								s_objUtils.Send "time"
								main_ngvnc7addmems = False 
								ExitMain : Exit Sub
						    Case 2
								s_objUtils.Send ""
						End Select
						s_objUtils.WaitFor ">", 5
						s_objUtils.SendAndWaitFor dicData.Item("BMem") & "->X;" & dicData.Item("BCIC") & "->Y", ">", s_intTimeout
						s_objUtils.SendAndWaitFor "REPEAT " & dicData.Item("Qty") & " (ADD " & dicData.Item("CLLI") & " X Y;1+X->X;1+Y->Y);QUIT ALL", ">", s_intTimeout
					End If
				End If
			Else 
				blnDone = True
			End If 
		Loop While Not blnDone
	Else 
		s_objUtils.CancelMsgBox
	End If
	
	ExitMain
End Sub 'Main
Sub ExitMain()
	s_objUtils.ScriptEnd
	Set s_objUtils = Nothing
	Set s_objDialogBox = Nothing
	Set s_objStrings = Nothing
	Set s_objWatchFor = Nothing
	Set s_objScriptTab = Nothing 
End Sub 'ExitMain

'********'*********'*********'*********'*********'*********'*********'*********'
'C7AddMemEntriesValid
'********'*********'*********'*********'*********'*********'*********'*********'
Private Function C7AddMemEntriesValid(dicData) 'returns boolean
	Dim blnReturn
	Dim strTitle
	
	strTitle = "C7 AddMEM VALIDATION"
	If (Not s_objUtils.IsPosInteger(dicData.Item("BMem"))) Then 
		crt.Dialog.MessageBox "Bad input in Start Mem", strTitle, vbOKOnly & vbInformation
		blnReturn = False 
	ElseIf (Not s_objUtils.IsPosInteger(dicData.Item("BCIC"))) Then
		crt.Dialog.MessageBox "Bad input in Start CIC", strTitle, vbOKOnly & vbInformation
		blnReturn = False 
	ElseIf (Not s_objUtils.IsPosInteger(dicData.Item("Qty"))) Then
		crt.Dialog.MessageBox "Bad input in Quantity", strTitle, vbOKOnly & vbInformation
		blnReturn = False 
	ElseIf (s_objStrings.IntVal(dicData.Item("BMem")) + s_objStrings.IntVal(dicData.Item("Qty")) > 10000) Then
		crt.Dialog.MessageBox "DMS member numbers cannnot exceed 9999", strTitle, vbOKOnly & vbInformation
		blnReturn = False 
	ElseIf (s_objStrings.IntVal(dicData.Item("BCIC")) < 1) Then 
		crt.Dialog.MessageBox "Start CIC must be greater than 0", strTitle, vbOKOnly & vbInformation
		blnReturn = False 
	ElseIf (s_objStrings.IntVal(dicData.Item("Qty")) < 1 or s_objStrings.IntVal(dicData.Item("Qty")) > 96) Then 
		crt.Dialog.MessageBox "Quantity must be between 1 and 96", strTitle, vbOKOnly & vbInformation
		blnReturn = False 
	ElseIf (s_objStrings.IntVal(dicData.Item("BCIC")) + s_objStrings.IntVal(dicData.Item("Qty")) > 16384) Then 
		crt.Dialog.MessageBox "CIC cannot exceed 16383", strTitle, vbOKOnly & vbInformation
		blnReturn = False 
	Else
	 	blnReturn = True 
	End If 

	C7AddMemEntriesValid = blnReturn
End Function 

'********'*********'*********'*********'*********'*********'*********'*********'
'C7AddMemDialog
'********'*********'*********'*********'*********'*********'*********'*********'
Private Function C7AddMemDialog(ByRef dicData)
	Dim blnTryAgain
	Dim strCLLI
	Dim strBMem
	Dim strBCIC
	Dim strQty
	
	
	blnTryAgain = False 
	Do 	
		strCLLI = dicData.Item("CLLI")
		strBMem = dicData.Item("BMem")
		strBCIC = dicData.Item("BCIC")
		strQty = dicData.Item("Qty")
		
		With s_objDialogBox
			.Clear
			.Method = "HTA"
			.Height = 200
			.Width = 375
			.LabelWidths = "100px"
			.Title = "Build C7TRKMEMs"
			.AddItem Array("type=text", "name=strCLLI", "label=CLLI:", "accesskey=L", "value=" & strCLLI, "size=10", "terminate=newline")
			.AddItem Array("type=text", "name=strBMem", "label=Start Mem:", "accesskey=M", "value=" & strBMem, "size=10", "terminate=newline")
			.AddItem Array("type=text", "name=strBCIC", "label=Start CIC:", "accesskey=I","value=" & strBCIC, "size=3", "terminate=newline")
			.AddItem Array("type=text", "name=strQty", "label=Quantity (max 96):", "accesskey=Q","value=" & strQty, "size=3", "terminate=newline")
			.Show
			C7AddMemDialog = .Status
			If (.Status <> vbOK) Then 
				Exit Function  
			End If 
			
			strCLLI = .Responses.Item("strCLLI")           
			strBMem = .Responses.Item("strBMem")           
			strBCIC = .Responses.Item("strBCIC") 
			strQty = .Responses.Item("strQty") 
			
			dicData.Item("CLLI") = strCLLI
			dicData.Item("BMem") = strBMem
			dicData.Item("BCIC") = strBCIC
			dicData.Item("Qty") = strQty
	
		End With
		'Should add some data validation here.
	Loop While blnTryAgain
End Function 


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
