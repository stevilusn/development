'#$language = "VBScript"
'#$interface = "1.0"
'Name: Buildmems.vbs

Option Explicit

Dim s_intTimeout
Dim s_objScriptTab	'Gets set during Initialize.
Dim s_objUtils
Dim s_objDialogBox
Dim s_objStrings
Dim s_objArrays
Dim s_objWatchFor

Dim g_blnTrainMode

'********'*********'*********'*********'*********'*********'*********'*********'
'Main
'********'*********'*********'*********'*********'*********'*********'*********'
Sub Main() 
	Dim dicSwitchData
	
	If (Not Initialize) Then Exit Sub
	s_objUtils.ScriptStart
	s_objScriptTab.Screen.Synchronous = True 

	Set dicSwitchData = CreateObject("scripting.dictionary")
	
'TODO: Add test to see if the screen is availible.
	If (s_objUtils.LeaveAll(10)) Then 'This tests if you have access to the screen.
		buildmem_InitVars
		If (GetGWCs(dicSwitchData)) Then 
			BuildMemDialog dicSwitchData 
		End If 
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
	Set s_objArrays = Nothing
	Set s_objWatchFor = Nothing
	Set s_objScriptTab = Nothing 
End Sub 'Main


'********'*********'*********'*********'*********'*********'*********'*********'
'buildmem_InitVars
'********'*********'*********'*********'*********'*********'*********'*********'
Sub buildmem_InitVars()
	g_blnTrainMode = False
End Sub 'buildmem_InitVars

'********'*********'*********'*********'*********'*********'*********'*********'
'GetGWCs
'********'*********'*********'*********'*********'*********'*********'*********'
Function GetGWCs(ByRef dicSwitchData)
	Dim strNGVNSwitchesFileName
	Dim strGWCSSwitchFileName
	Dim strGWCSFileName	
	Dim blnIPFound
	Dim strHostName
	Dim strSwitchName
	Dim strIPAddress
	Dim arrstrSwitches
	Dim arrstrSwitch
	Dim arrstrGWCs
	Dim strGWC
	Dim strLine
	Dim strKey
	Dim strValue
	

	strNGVNSwitchesFileName = DIR_NGVN_DATA & "NGVN_Switches.txt"

	If (s_objUtils.FileExists(strNGVNSwitchesFileName)) Then
		arrstrSwitches = s_objArrays.ReadFile(strNGVNSwitchesFileName)
	Else 
		crt.Dialog.MessageBox "File '" & strNGVNSwitchesFileName & "' was not found.  Please contact support.", "SWITCHES NOT FOUND", vbOKOnly + vbCritical
		GetGWCs = False 
		Exit Function
	End If 
	
	blnIPFound = False
	'Check that the current tabs connection exists in the NGVN_Switches.txt file.
	strHostName = s_objUtils.getHostName(s_objScriptTab)
	For Each strLine In arrstrSwitches
		If (InStr(strLine, ",")) Then 
			arrstrSwitch = Split(strLine, ",")
			strSwitchName = arrstrSwitch(0)
			strIPAddress = arrstrSwitch(1)
			If (strIPAddress = strHostName) Then
				blnIPFound = True
				Exit For 
			End If
		End If 
	Next
	
	strGWCSFileName = DIR_NGVN_DATA & "GWCs.txt"
	strGWCSSwitchFileName = DIR_NGVN_DATA & "GWCS." & strSwitchName & ".txt"
	If (blnIPFound = False) Then
		crt.Dialog.MessageBox "Switch specific GWCs not found for '" & strHostName & "' in file '" & strNGVNSwitchesFileName & "'.  Will use the default file '" & strGWCSFileName & "' to load GWC data.", "SWITCH NOT FOUND", vbOKOnly + vbInformation
		If (s_objUtils.FileExists(strGWCSFileName)) Then
			arrstrGWCs = s_objArrays.ReadFile(strGWCSFileName)
		Else 
			crt.Dialog.MessageBox "Cannot open file " & strGWCSFileName & ".  Please contact support.", "FILE NOT FOUND", vbOKOnly + vbCritical
			GetGWCs = False 
			Exit Function 
		End If 
	Else 
		If (s_objUtils.FileExists(strGWCSSwitchFileName)) Then
			arrstrGWCs = s_objArrays.ReadFile(strGWCSSwitchFileName)
		ElseIf (s_objUtils.FileExists(strGWCSFileName)) Then
			crt.Dialog.MessageBox "Cannot open file " & strGWCSSwitchFileName & ".  Will attempt to open " & strGWCSFileName & " instead.", "FILE NOT FOUND", vbOKOnly + vbInformation
			arrstrGWCs = s_objArrays.ReadFile(strGWCSFileName)
		Else
			crt.Dialog.MessageBox "Cannot open file " & strGWCSFileName & ".  Please contact support.", "FILE NOT FOUND", vbOKOnly + vbCritical
			GetGWCs = False 
			Exit Function 
		End If 
	End If 

	For Each strGWC In arrstrGWCs
		If (InStr(strGWC, ",") > 0) Then
			strKey = Left(strGWC, InStr(strGWC, ",") - 1) 'FMS
			strValue = Mid(strGWC, InStr(strGWC, ",") + 1) 'GWC, Node, OC3
			dicSwitchData.Add strKey, strValue
		End If
	Next
	GetGWCs = True
End Function 'GetGWCs

'********'*********'*********'*********'*********'*********'*********'*********'
'BuildMemDialog
'********'*********'*********'*********'*********'*********'*********'*********'
Private Function BuildMemDialog(dicSwitchData)
	Dim i
	Dim blnTryAgain
	Dim blnGWCAssignmentsUsingFMSEqpt
	Dim blnDone
	Dim blnFound
	Dim dicData
	Dim varStatus
	Dim strCLLI
	Dim strBMem
	Dim strQty
	Dim strGWC
	Dim strNode
	Dim strStTerm	
	
	Set dicData = CreateObject("scripting.dictionary")
	dicData.Add "CLLI", ""
	dicData.Add "BMem", ""
	dicData.Add "Qty", "24"
	dicData.Add "GWC", ""
	dicData.Add "Node", ""
	dicData.Add "StTerm", ""
	dicData.Add "FMSSelection", ""
	dicData.Add "STS", ""
	dicData.Add "T1", ""
	dicData.Add "DS0", ""

	blnTryAgain = False 
	blnDone = False 
	Do
		strCLLI = dicData.Item("CLLI")
		strBMem = dicData.Item("BMem")
		strQty = dicData.Item("Qty")
		strGWC = dicData.Item("GWC")
		strNode = dicData.Item("Node")
		strStTerm = dicData.Item("StTerm")	

		blnFound = True
		With s_objDialogBox
			blnGWCAssignmentsUsingFMSEqpt = False
			.Clear
			.Method = "HTA"
			.Height = 250
			.Width = 375
			.LabelWidths = "120px"
			.Title = "Build GWC Members"
			.AddItem Array("type=text", "name=strCLLI", "label=CLLI:", "accesskey=I", "value=" & strCLLI, "size=10", "terminate=newline")
			.AddItem Array("type=text", "name=strBMem", "label=Start Mem Number:", "accesskey=M", "value=" & strBMem, "size=10", "terminate=newline")
			.AddItem Array("type=text", "name=strQty", "label=Qty:", "accesskey=Q", "value=" & strQty, "size=10", "terminate=newline")
			.AddItem Array("type=text", "name=strGWC", "label=GWC:", "accesskey=G", "value=" & strGWC, "size=10", "terminate=newline")
			.AddItem Array("type=text", "name=strNode", "label=Node:", "accesskey=N", "value=" & strNode, "size=10", "terminate=newline")
			.AddItem Array("type=text", "name=strStTerm", "label=Start Terminal:", "accesskey=T", "value=" & strStTerm, "size=10", "terminate=newline")
			.AddItem Array("type=html", "</br>"  & vbCrLf)
			.AddItem Array("type=html", "<button type='submit' name='GWCAssignmentsUsingFMSEqpt' title='GWC Assignments Using FMS Eqpt.' AccessKey='A' Onclick=document.all('ButtonHandler').value='GWCAssignmentsUsingFMSEqpt';>GWC <u>A</u>ssignments Using FMS Eqpt </button>"  & vbCrLf)
			.AddItem Array("type=html", "<br>"  & vbCrLf)
			.Show
			varStatus = .Status
			Select Case varStatus
				Case vbOK
					BuildMemDialog = vbOk
				Case "GWCAssignmentsUsingFMSEqpt"
					BuildMemDialog = vbOk
					blnGWCAssignmentsUsingFMSEqpt = True
				Case Else
					BuildMemDialog = vbCancel : Exit Function  
			End select
			strCLLI = Trim(.Responses.Item("strCLLI"))           
			strBMem = Trim(.Responses.Item("strBMem"))           
			strQty = Trim(.Responses.Item("strQty"))           
			strGWC = Trim(.Responses.Item("strGWC"))           
			strNode = Trim(.Responses.Item("strNode"))           
			strStTerm = Trim(.Responses.Item("strStTerm"))           
		End With
		'Should add some data entry validation here.
		dicData.Item("CLLI") = strCLLI           
		dicData.Item("BMem") = strBMem           
		dicData.Item("Qty") = strQty           
		dicData.Item("GWC") = strGWC           
		dicData.Item("Node") = strNode           
		dicData.Item("StTerm") = strStTerm           
	
		If (blnGWCAssignmentsUsingFMSEqpt) Then
			If (GetFMSInfoDlg(dicData, dicSwitchData) <> vbOK) Then
				BuildMemDialog = vbCancel : Exit Function 
			End If 
		Else
			s_objUtils.LeaveAll(s_intTimeout)
			If (s_objUtils.IsPosInteger(strCLLI)) Then
				strCLLI = Trim(s_objWatchFor.SearchClliCdr(s_objStrings.IntVal(strCLLI)))
				If (strCLLI = "") Then
					blnFound = False
				Else
					dicData.Item("CLLI") = strCLLI
				End If
			End If
			If (blnFound) Then
				i = DoBuildMems(dicData) 'Zero would indicate success, -1 would indicate falure.
				'TODO: What is "i" used for?  CROSSTALK code didn't use it?  Guess nothing is done with it since the next thing that happens is that the code finishes up and exits.
				blnDone = True
			Else
				crt.Dialog.MessageBox "FMS TG number not found in CLLICDR!", "NOT FOUND", vbOKOnly & vbExclamation
			End If
		End If
	Loop While Not blnDone
End Function 'BuildMemDialog

'********'*********'*********'*********'*********'*********'*********'*********'
'StrInIntRange
'********'*********'*********'*********'*********'*********'*********'*********'
Private Function StrInIntRange(strTest, intMin, intMax) 'returns boolean
	Dim i 
	Dim blnReturn
	
	blnReturn = True
	
	If (s_objUtils.IsPosInteger(strTest)) Then 
		i = s_objStrings.Val(strTest)
		If (i > intMin - 1 And i < intMax + 1) Then 
			blnReturn = True
		Else
			blnReturn = False
		End If 
	Else 
		blnReturn = False
	End If 
	StrInIntRange = blnReturn
End Function 'StrInIntRange

'********'*********'*********'*********'*********'*********'*********'*********'
'GetFMSInfoDlg
'********'*********'*********'*********'*********'*********'*********'*********'
Private Function GetFMSInfoDlg(ByRef dicData, dicSwitchData)
	Dim blnDone
	Dim strListSelection
	Dim strOC3
	Dim strSTS
	Dim strDS0
	Dim strT1
	Dim intOC3
	Dim intSTS
	Dim intDS0
	Dim intT1
	Dim arrSwitchData
	Dim intStTerm
	
	Dim CHANNELS_PER_T1
	Dim CHANNELS_PER_STS
	Dim CHANNELS_PER_OC3
	
	blnDone = False
	
	Do 	
		strListSelection = dicData.Item("FMSSelection")
		strSTS = dicData.Item("STS")
		strT1 = dicData.Item("T1")
		strDS0 = dicData.Item("DS0")
		With s_objDialogBox
			.Clear
			.Method = "HTA"
			.Height = 200
			.Width = 450
			.LabelWidths = "120px"
			.Title = "FMS Assignments"
			.AddItem Array("type=select", "name=strListSelection", "label=Site Equipment </br>Bay/Shelf/Slot/Port:" ,"accesskey=E", "value=" & strListSelection, "choices="& Join(dicSwitchData.Keys, "|"), "size=1", "multiple=False", "terminate=newline")
			.AddItem Array("type=html", "</br><hr4>Timeslot Info</hr4>")
			.AddItem Array("type=text", "name=strSTS", "label=STS:", "title=STS must be a value from 1 to 3.", "accesskey=S", "value=" & strSTS, "size=10", "terminate=newline")
			.AddItem Array("type=text", "name=strT1", "label=DS1:", "title=T1 must be a value from 1 to 28.", "accesskey=D", "value=" & strT1, "size=10", "terminate=newline")
			.AddItem Array("type=text", "name=strDS0", "label=DS0 Channel:", "title=DS0 must be a value from 1 to 24.", "accesskey=0", "value=" & strDS0, "size=10", "terminate=newline")
			.Show
			GetFMSInfoDlg = .Status
			If (.Status <> vbOK) Then 
				Exit Function  
			End If 
			strListSelection = .Responses.Item("strListSelection")           
			strSTS = .Responses.Item("strSTS")          
			strT1 = .Responses.Item("strT1") 
			strDS0 = .Responses.Item("strDS0") 
		End With
		
		dicData.Item("FMSSelection") = strListSelection
		dicData.Item("STS") = strSTS
		dicData.Item("T1") = strT1
		dicData.Item("DS0") = strDS0
		arrSwitchData = Split(dicSwitchData.Item(strListSelection), ",") 'GWC, Node, OC3
		dicData.Item("GWC") = arrSwitchData(0)
		dicData.Item("Node") = arrSwitchData(1)
		strOC3 = arrSwitchData(2)
		
		If (Not StrInIntRange(strOC3, 1, 2)) Then 
			crt.Dialog.MessageBox "OC3 must be a number from 1 to 2!  This is an issue in the GWC lookup files.  Please contact support.  Exiting now.", "INVALID NUMBER", vbOKOnly & vbCritical
			GetFMSInfoDlg = vbCancel : Exit Function
		ElseIf (Not StrInIntRange(strSTS, 1, 3)) Then 
			crt.Dialog.MessageBox "STS must be a value from 1 to 3!", "INVALID NUMBER", vbOKOnly & vbExclamation
		ElseIf (Not StrInIntRange(strT1, 1, 28)) Then 
			crt.Dialog.MessageBox "T1 must be a value from 1 to 28!", "INVALID NUMBER", vbOKOnly & vbExclamation
		ElseIf (Not StrInIntRange(strDS0, 1, 24)) Then 
			crt.Dialog.MessageBox "DS0 must be a value from 1 to 24!", "INVALID NUMBER", vbOKOnly & vbExclamation
		Else 
			intOC3 = s_objStrings.Val(strOC3)
			intSTS = s_objStrings.Val(strSTS)
			intT1  = s_objStrings.Val(strT1)
			intDS0  = s_objStrings.Val(strDS0)
			blnDone = True
		End If

	Loop While Not blnDone

	CHANNELS_PER_T1  = 24
	CHANNELS_PER_STS = CHANNELS_PER_T1 * 28
	CHANNELS_PER_OC3 = CHANNELS_PER_STS * 3
	intStTerm = (intOC3 - 1) * CHANNELS_PER_OC3 _
			+ (intSTS - 1) * CHANNELS_PER_STS _
			+ (intT1 - 1) * CHANNELS_PER_T1 _
			+ intDS0
	dicData.Item("StTerm") = CStr(intStTerm)
End Function 'GetFMSInfoDlg

'********'*********'*********'*********'*********'*********'*********'*********'
'SearchForClliDlg
'This is all it did in CROSSTALK.  Guess this is just a hard coded manual 
'switch to allow the SearchClliCdr lookup to be turned on or off?
'********'*********'*********'*********'*********'*********'*********'*********'
Private Function SearchForClliDlg 'returns boolean
	SearchForClliDlg = True
End Function 'SearchForClliDlg

'********'*********'*********'*********'*********'*********'*********'*********'
'GetActualClli
'********'*********'*********'*********'*********'*********'*********'*********'
Private Function GetActualClli(strTest) 'returns string
	Dim intResult
	Dim blnImage
	
	blnImage = False 
	intResult = s_objUtils.SendAndWaitFor("LEAVE ALL;QUIT ALL;TABLE CLLI;POS " & strTest, _
				Array("TUPLE NOT FOUND", "COMMAND DISALLOWED DURING DUMP", ">"), s_intTimeout)
	Select Case intResult
	    Case 0
			crt.Dialog.MessageBox "Timed out while watching for strings ""TUPLE NOT FOUND"", ""COMMAND DISALLOWED DURING DUMP"", "">"".  Will attempt to continue processing.", "TIMEOUT", vbOKOnly & vbExclamation
	    	GetActualClli = ""
	    Case 1 '"TUPLE NOT FOUND"
			s_objUtils.LeaveAll(s_intTimeout)
			If (SearchForClliDlg) Then 
				If (s_objUtils.IsPosInteger(strTest)) Then 
		            GetActualClli = s_objWatchFor.SearchClliCdr(s_objStrings.IntVal(strTest))
				Else 
					GetActualClli = ""
				End If 
			Else 
				s_objUtils.LeaveAll(s_intTimeout)
				GetActualClli = ""
			End If 
	    Case 2 '"COMMAND DISALLOWED DURING DUMP"
	    	blnImage = True 
	    	GetActualClli = ""
	    Case 3 '">"
			GetActualClli =  Replace(strTest, " ", "")
	End Select
End Function 'GetActualClli

'********'*********'*********'*********'*********'*********'*********'*********'
'DoBuildMems
'********'*********'*********'*********'*********'*********'*********'*********'
Private Function DoBuildMems(dicData) 'returns integer
	Dim intRetVal
	Dim strTitle
	Dim strActualClli
	Dim intStTerm
	Dim intQty
	Dim strCLLI
	Dim intResult
	Dim blnImage
	Dim strCmd1
	Dim strCmd2
	
	intRetVal = 0
	blnImage = False 
			
	strTitle = "AddMEM VALIDATION"
	
	If (dicData.Item("CLLI") = "") Then 
		crt.Dialog.MessageBox "CLLI cannot be blank!", strTitle, vbOKOnly & vbExclamation
		DoBuildMems = -1
		Exit Function 
	ElseIf (Not StrInIntRange(dicData.Item("BMem"), 0, 9999)) Then 
		crt.Dialog.MessageBox "Start Mem Num must be a value from 1 to 9999!", strTitle, vbOKOnly & vbExclamation
		DoBuildMems = -1
		Exit Function 
	ElseIf (Not StrInIntRange(dicData.Item("Qty"), 1, 24)) Then 
		crt.Dialog.MessageBox "Qty must be a value from 1 to 24!", strTitle, vbOKOnly & vbExclamation
		DoBuildMems = -1
		Exit Function 
	ElseIf (Not s_objUtils.IsPosInteger(dicData.Item("GWC"))) Then 
		crt.Dialog.MessageBox "GWC must be an integer value!", strTitle, vbOKOnly & vbExclamation
		DoBuildMems = -1
		Exit Function 
	ElseIf (Not s_objUtils.IsPosInteger(dicData.Item("Node"))) Then  
		crt.Dialog.MessageBox "Node must be an integer value!", strTitle, vbOKOnly & vbExclamation
		DoBuildMems = -1
		Exit Function 
	ElseIf (Not StrInIntRange(dicData.Item("StTerm"), 1, 4032)) Then
		crt.Dialog.MessageBox "Start Terminal must be a value from 1 to 4032!", strTitle, vbOKOnly & vbExclamation
		DoBuildMems = -1
		Exit Function 
	End If 
	
	strCLLI = GetActualClli(dicData.Item("CLLI"))
	If (strCLLI = "") Then
		crt.Dialog.MessageBox "Cannot verify CLLI!", strTitle, vbOKOnly & vbExclamation
		DoBuildMems = -1
		Exit Function 
	Else 
		strActualClli = strCLLI
	End If 
	
	intStTerm = s_objStrings.Val(dicData.Item("StTerm"))
	intQty = s_objStrings.Val(dicData.Item("Qty"))

	If ((intQty + intStTerm - 1) > 4032) Then 
		crt.Dialog.MessageBox "Too many members for this Start Terminal!", strTitle, vbOKOnly & vbExclamation
		DoBuildMems = -1
		Exit Function 
	End If 

	s_objUtils.QuitAll(60)
	
	intResult = s_objUtils.SendAndWaitFor("SEND SINK;TABLE TRKMEM;BOTTOM;FORMAT PACK;SEND PREVIOUS;VER OFF", _
				Array("COMMAND DISALLOWED DURING DUMP", ">"), 60)
	Select Case intResult
	    Case 0
			crt.Dialog.MessageBox "Timed out while waiting for strings ""COMMAND DISALLOWED DURING DUMP"", "">"".  Will attempt to continue processing.", "TIMEOUT", vbOKOnly & vbExclamation
	    Case 1 '"COMMAND DISALLOWED DURING DUMP"
	    	blnImage = True 
	    Case 2 '">"
			s_objUtils.Send ""
	End Select

	If (blnImage = True) Then 
		s_objUtils.WaitFor ">", s_intTimeout
		crt.Dialog.MessageBox "SWITCH IS CURRENTLY IN IMAGE, PLEASE TRY LATER", "IN IMAGE", vbOKOnly & vbExclamation
		s_objUtils.SendAndWaitFor "", ">", s_intTimeout
		s_objUtils.SendAndWaitFor "table imgsched;lis all", ">", s_intTimeout
'		s_objUtils.Send "date" 'This was commented out in the original Crosstalk code.
		intRetVal = -1
	Else 
		strCmd1 = dicData.Item("BMem") & "->X;" & dicData.Item("StTerm") & "->Y"
		strCmd2 = "REPEAT " & dicData.Item("Qty") & " (ADD " & strActualClli & " X 0 GWC " _
				& dicData.Item("GWC") & " " & dicData.Item("Node") & " Y;1+X->X;1+Y->Y);QUIT ALL"
		If (g_blnTrainMode) Then 
			crt.Dialog.MessageBox "Training Dialog", "TRAIN MODE", vbOKOnly & vbInformation
		Else
			s_objUtils.WaitFor ">", 60
			s_objUtils.SendAndWaitFor strCmd1, ">", 60 'Wait for the command itself (with '>' in it twice) and then wait for the '>' prompt.
			s_objUtils.Send ""
			s_objUtils.SleepSeconds 4
			s_objUtils.SendAndWaitFor strCmd2, strCmd2, 60
			s_objUtils.WaitFor ">", 60
			s_objUtils.Send ""
			s_objUtils.SleepSeconds 7
			s_objUtils.WaitFor ">", 60
		End If 
	End If 
	DoBuildMems = intRetVal
End Function 'DoBuildMems


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

'********'*********'*********'*********'*********'*********'*********'*********'
'DicToStr
'********'*********'*********'*********'*********'*********'*********'*********'
Public Function DicToStr(objDic)
	Dim arrKeys
	Dim arrItems
	Dim i
	Dim strString
	
	If (TypeName(objDic) = "Dictionary") Then
		arrKeys = objDic.Keys
		arrItems = objDic.Items
		For i = LBound(arrKeys) To UBound(arrKeys) Step 1
			strString = strString & arrKeys(i) & "=" & arrItems(i) & ", "
		Next
		DicToStr = Left(strString, Len(strString) - 2)
	Else
		DicToStr = ""
	End If 	
End Function 'DicToStr



