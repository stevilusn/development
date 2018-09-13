'#$language = "VBScript"
'#$interface = "1.0"
'Name: QDN Analyzer.vbs

Option Explicit

Dim s_objScriptTab
Dim s_objUtils
Dim s_objDialogBox
Dim s_objStrings
Dim s_objArrays
Dim s_intTimeout

'********'*********'*********'*********'*********'*********'*********'*********'
'Main
'********'*********'*********'*********'*********'*********'*********'*********'
Sub Main
	Dim objImport
	Dim strO_Tele
	Dim blnDisplayResults
	Dim objQDN
	
	If (Not Initialize) Then
		Exit Sub
	End If
	s_objUtils.ScriptStart

	Set objImport = New clsImport
	objImport.File(DIR_SECURECRT & "clsQDN.vbs")
	Set objImport = Nothing 
	Set objQDN = New clsQDN
	
	If (s_objUtils.SendAndWaitFor("leave all", ">", s_intTimeout)) Then 'This tests if you have access to the screen.
		If (get_user_data(strO_Tele, blnDisplayResults) = vbOK) Then 
			If (objQDN.GetData(strO_Tele)) Then 
 				If (Not use_data(objQDN)) Then
 					'Continue on an error to allow the display of QDN results.
 				End If 
				If (blnDisplayResults) Then 
					Call display_results(objQDN)
				End If
			End If 
		End If
	Else
		s_objUtils.CancelMsgBox
	End If 

	s_objUtils.ScriptEnd
	Set objQDN = Nothing
	Set s_objDialogBox = Nothing
	Set s_objStrings = Nothing
	Set s_objArrays = Nothing
	Set s_objUtils = Nothing
End Sub 'Main

'********'*********'*********'*********'*********'*********'*********'*********'
'get_user_data
'********'*********'*********'*********'*********'*********'*********'*********'
Function get_user_data(ByRef strO_Tele, ByRef blnDisplayResults)
	blnDisplayResults = False
	Do 
		With s_objDialogBox
			.Clear
			.Height = 200
			.Width = 350
			.LabelWidths = "100px"
			.Title = "QDN Analyzer"
			.AddItem Array("type=paragraph", "name=label1", "value=" & s_objUtils.CleanTrashMsg("DN"), "terminate=newline")
			.AddItem Array("type=text", "name=strO_Tele", "label=DN:", "accesskey=D", "value=" & strO_Tele, "size=10", "terminate=newline")
			.AddItem Array("type=checkbox", "name=blnDisplayResults", "label=Display Results:", "value=" & blnDisplayResults, "accesskey=R", "terminate=newline")
			.Show
			get_user_data = .Status
			If .Status <> vbOK Then 
				Exit Function  
			End If 
			strO_Tele = s_objStrings.RemoveNonNumericCharacters(.Responses.Item("strO_Tele"))
			blnDisplayResults = CBool(.Responses.Item("blnDisplayResults"))      
		End With
	Loop Until s_objUtils.TestLength(strO_Tele, "DN", 10)
End Function 'get_user_data

'********'*********'*********'*********'*********'*********'*********'*********'
'use_data
'********'*********'*********'*********'*********'*********'*********'*********'
Function use_data(objQDN)
	Dim arrScreeData
	Dim strRateArea
	Dim strFETXLA
	
	If (objQDN.Data.Item("PORTED-IN") > "") Then  
		s_objUtils.SendAndWaitFor "qlrn "  & objQDN.TeleNumber, ">", s_intTimeout
	Else
		s_objUtils.Send "", s_intTimeout
	End If 

	If (Not parse_lineattr_data(objQDN.Data.Item("LNATTIDX"), strRateArea)) Then
		crt.Dialog.Messagebox "Could not retrieve a Rate Area.  Exiting execution.", "MISSING DATA", vbOKOnly & vbCritical
		use_data = False
		Exit Function 
	End If 

	s_objUtils.SendAndWaitFor "lis all (dfltra eq " & strRateArea & ")", ">", s_intTimeout
	s_objUtils.SendAndWaitFor "leave all", ">", s_intTimeout
	s_objUtils.SendAndWaitFor "table ratearea ;pos " & strRateArea, ">", s_intTimeout
	s_objUtils.SendAndWaitFor "leave all", ">", s_intTimeout
	
	If (Not parse_custhead_data(objQDN.Data.Item("CUSTGRP"), strFETXLA)) Then
		crt.Dialog.Messagebox "Could not retrieve a FETXLA.  Exiting execution.", "MISSING DATA", vbOKOnly & vbCritical
		use_data = False
		Exit Function 
	End If 
	
	s_objUtils.Message "Entering Commands ..."
	s_objUtils.SendAndWaitFor "leave all", ">", s_intTimeout
	s_objUtils.SendAndWaitFor "table ibnxla;pos " & strFETXLA & " 57", ">", s_intTimeout
	s_objUtils.SendAndWaitFor "lis 30", ">", s_intTimeout
	s_objUtils.SendAndWaitFor "leave all", ">", s_intTimeout
	s_objUtils.SendAndWaitFor "QSL " & objQDN.Data.Item("LINE EQUIPMENT NUMBER") & " ALL", ">", s_intTimeout
	s_objUtils.SendAndWaitFor "leave all", ">", s_intTimeout
	If (objQDN.Flags.Item("SC1")) Then  
		s_objUtils.SendAndWaitFor "table ibnsc", ">", s_intTimeout
		s_objUtils.SendAndWaitFor "pos " & objQDN.Data.Item("LINE EQUIPMENT NUMBER") & "0 SCS 2", ">", s_intTimeout
		s_objUtils.SendAndWaitFor "list 10", ">", s_intTimeout
	End If 
	use_data = True 
End Function 'use_data

'********'*********'*********'*********'*********'*********'*********'*********'
'parse_lineattr_data
'********'*********'*********'*********'*********'*********'*********'*********'
Function parse_lineattr_data(strLNATTIdx, ByRef strRateArea)
	Dim arrScreenData
	Dim strScreenRow
	Dim intPos
	
	s_objUtils.Message "Analyzing Table LINEATTR data..."
	arrScreenData = s_objUtils.SendAndReadStrings("table lineattr;pos " _
			& strLNATTIdx, ">", s_intTimeout)
	For Each strScreenRow In arrScreenData
		If (InStr(strScreenRow, Left(strLNATTIdx, 6)) > 0) Then    
			intPos = InStr(strScreenRow, "_")
			strRateArea = Mid(strScreenRow, intPos + 5, 13)
		End If 
	Next
	If (Len(strRateArea) = 0) Then
		parse_lineattr_data = False 
	Else
		parse_lineattr_data = True
	End If
End Function 'parse_lineattr_data

'********'*********'*********'*********'*********'*********'*********'*********'
'parse_custhead_data
'********'*********'*********'*********'*********'*********'*********'*********'
Function parse_custhead_data(strCustGrp, ByRef strFETXLA)
	Dim arrScreenData
	Dim strScreenRow
	Dim intPos
	
	s_objUtils.Message "Analyzing Table CUSTHEAD data..."
	arrScreenData = s_objUtils.SendAndReadStrings("table custhead;pos " _
			& strCustGrp, ">", s_intTimeout)
	For Each strScreenRow In arrScreenData
		If (InStr(strScreenRow, "FETXLA") > 0) Then    
			intPos = InStr(strScreenRow, "F")
			strFETXLA = Mid(strScreenRow, intPos + 7, 6)
		End If 
	Next
	If (Len(strFETXLA) = 0) Then
		parse_custhead_data = False 
	Else
		parse_custhead_data = True 
	End If
End Function 'parse_custhead_data

'********'*********'*********'*********'*********'*********'*********'*********'
'display_results
'********'*********'*********'*********'*********'*********'*********'*********'
Sub display_results(objQDN)
	Dim strFileName
	Dim arrResults
	Dim objShell
	
	arrResults = objQDN.Report
	strFileName = DIR_NGVN_DATA & "QDN_Results.txt"
		
'	arrResults = s_objArrays.AppendVariant(arrResults, " QDN Results for DN " + objQDN.TeleNumber)
'	arrResults = s_objArrays.AppendVariant(arrResults, "SNPA:" + s_objStrings.Pack(objQDN.Data.Item("SNPA"), " ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "   DN:" + s_objStrings.Pack(objQDN.Data.Item("DN"), " ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "  " + s_objStrings.Pack(objQDN.Data.Item("PORTED-IN"), " ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "  TYPE:" + s_objStrings.Pack(objQDN.Data.Item("TYPE"), " ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "   SIG:" + s_objStrings.Pack(objQDN.Data.Item("SIG"), " ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "LNATTIDX:" + s_objStrings.Pack(objQDN.Data.Item("LNATTIDX"), " ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "   LEN:" + s_objStrings.Pack(objQDN.Data.Item("LINE EQUIPMENT NUMBER"), " ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "   LINE CLASS CODE:" + s_objStrings.Pack(objQDN.Data.Item("LINE CLASS CODE"), " ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "IBN TYPE:" + s_objStrings.Pack(objQDN.Data.Item("IBN TYPE"), " ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "   CUSTGRP:" + s_objStrings.Pack(objQDN.Data.Item("CUSTGRP"), " ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "   SUBGRP:" + s_objStrings.Pack(objQDN.Data.Item("SUBGRP"), " ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "   NCOS:" + s_objStrings.Pack(objQDN.Data.Item("NCOS"), " ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "   LTG:" + s_objStrings.Pack(objQDN.Data.Item("LINE TREATMENT GROUP"), " ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "CARDCODE:" + s_objStrings.Pack(objQDN.Data.Item("CARDCODE"), " ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "   GND:" + s_objStrings.Pack(objQDN.Data.Item("GND"), " ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "   PADGRP:" + s_objStrings.Pack(objQDN.Data.Item("PADGRP"), " ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "   BNV:" + s_objStrings.Pack(objQDN.Data.Item("BNV"), " ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "   MNO:" + s_objStrings.Pack(objQDN.Data.Item("MNO"), " ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "PM NODE NUMBER:" + s_objStrings.Pack(objQDN.Data.Item("PM NODE NUMBER"), " ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "   PM TERMINAL NUMBER:" + s_objStrings.Pack(objQDN.Data.Item("PM TERMINAL NUMBER"), " ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "   CFW INDEX:" + s_objStrings.Pack(objQDN.Data.Item("CFW INDEX"), " ",1))
'	if (objQDN.Flags.Item("CWT")) Then
'		arrResults = s_objArrays.AppendVariant(arrResults, "CWT is built on the line")
'	End If 
	
	s_objArrays.WriteFile arrResults, strFileName
	s_objUtils.OpenFileImmediately = True 
	s_objUtils.OpenFile strFileName

End Sub 'display_results

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
		MsgBox "Could not find file " & strCurPath & "BootStrap.txt.", vbOK, vbCritical, "FILE NOT FOUND!"
	End If 
	s_intTimeout = DEFAULT_TIMEOUT
	Set objFSO = Nothing
End Function 'Initialize