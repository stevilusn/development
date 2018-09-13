'#$language = "VBScript"
'#$interface = "1.0"
'Name: LCA6SCRN.vbs

Option Explicit

Dim s_objScriptTab
Dim s_objUtils
Dim s_objArrays
Dim s_intTimeout

'********'*********'*********'*********'*********'*********'*********'*********'
'Main
'********'*********'*********'*********'*********'*********'*********'*********'
Sub Main
	Dim blnRet
	Dim objImport
	Dim strO_Tele
	Dim objQDN
 
	If (Not Initialize) Then
		Exit Sub
	End If
	s_objUtils.ScriptStart

	Set objImport = New clsImport
	objImport.File(DIR_SECURECRT & "clsQDN.vbs")
	Set objImport = Nothing 
	Set objQDN = New clsQDN

	blnRet = s_objUtils.SendAndWaitFor("leave all", ">", s_intTimeout) 
	If (blnRet) Then 'This tests if you have access to the screen.
		If (get_user_data(strO_Tele) = vbOK) Then	
			If (objQDN.GetData(strO_Tele)) Then 			
				Call display_results(strO_Tele, objQDN.Data.Item("LNATTIDX"))
			End If 
		End If 
	Else
		s_objUtils.CancelMsgBox
	End If 
	
	s_objUtils.ScriptEnd
	Set s_objScriptTab = Nothing
	Set objQDN = Nothing
	Set s_objArrays = Nothing
	Set s_objUtils = Nothing
End Sub 'Main


'********'*********'*********'*********'*********'*********'*********'*********'
'get_user_data
'********'*********'*********'*********'*********'*********'*********'*********'
Function get_user_data(ByRef strO_Tele)
	Dim blnTryAgain
	Do 
		strO_Tele = crt.Dialog.Prompt("DN:", "Enter the DN", strO_Tele)
		If (Len(strO_Tele) = 0) Then
			get_user_data = vbCancel
			Exit Function 
		End If

		'clean the spaces, underlines and dashes from the orig teles...   
		strO_Tele = s_objUtils.stripTN(strO_Tele)
			
		blnTryAgain = s_objUtils.TestLength(strO_Tele, "Orig Tele", 10)

	Loop Until blnTryAgain
	
	get_user_data = vbOK 
End Function 'get_user_data

'********'*********'*********'*********'*********'*********'*********'*********'
'parse_lineattr_data
'********'*********'*********'*********'*********'*********'*********'*********'
Function parse_lineattr_data(strLNATTIdx)
	Dim arrScreenData
	Dim strScreenRow
	Dim intPos 
	Dim strRateArea
	
	arrScreenData = s_objUtils.SendAndReadStrings("table lineattr;pos " _
			& strLNATTIdx, _
			">", s_intTimeout)
			
	For Each strScreenRow In arrScreenData
		If (InStr(strScreenRow, Left(strLNATTIdx, 6)) > 0) Then    
			intPos = InStr(strScreenRow, "_")
			strRateArea = Mid(strScreenRow, intPos + 6, 6)
		End If 
	Next
	parse_lineattr_data = strRateArea
End Function 'parse_lineattr_data
 
'********'*********'*********'*********'*********'*********'*********'*********'
'parse_lca6scrn_data
'********'*********'*********'*********'*********'*********'*********'*********'
Function parse_lca6scrn_data(strTeleNumber, strLNATTIdx)
	Dim arrScreenData
	Dim strScreenRow
	Dim intPos
	Dim strRateArea
	Dim arrLca6scrnData

	strRateArea = parse_lineattr_data(strLNATTIdx)
	
'TODO: Need to verify that this works.  I don't know if the pos command will return a ">" before the UP command.  Look at the CROSSTALK code.
	s_objUtils.SendAndWaitFor "leave all", ">", s_intTimeout
	arrScreenData = s_objUtils.SendAndReadStrings("Table lca6scrn;pos " _
			& strRateArea & " " & Left(strTeleNumber, 3) _
			& " " & Mid(strTeleNumber, 4, 3) & DEFAULT_LINE_TERMINATOR _
			& "UP 1000;LIS 2000", _
			">", s_intTimeout)
			
	arrLca6scrnData = Array()
	For Each strScreenRow In arrScreenData
		If (InStr(strScreenRow, strRateArea) > 0) then    
			arrLca6scrnData = s_objArrays.AppendVariant(arrLca6scrnData, strScreenRow)
		End If 
	Next 
	parse_lca6scrn_data = arrLca6scrnData
End Function 'parse_lca6scrn_data

'********'*********'*********'*********'*********'*********'*********'*********'
'display_results
'********'*********'*********'*********'*********'*********'*********'*********'
Sub display_results(strTeleNumber, strLNATTIdx)
	Dim arrstrErrorLog
	Dim strFileName
	Dim arrLca6scrnData
	
	arrLca6scrnData = parse_lca6scrn_data(strTeleNumber, strLNATTIdx)

	strFileName = DIR_NGVN_DATA & "LCA6SCRN_Results.txt"
	arrstrErrorLog = Array(_
			"Areas local to DN " + strTeleNumber, _
			"       LCA6KEY TOOFC TENDLOC", _
			"----------------------------")	
	arrstrErrorLog = s_objArrays.AppendArray(arrstrErrorLog, arrLca6scrnData)
	s_objArrays.WriteFile arrstrErrorLog, strFileName
	s_objUtils.OpenFileImmediately = False 
	s_objUtils.OpenFile strFileName	'Display the file via the default program.	
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
		MsgBox "Could not find file " & strCurPath & "BootStrap.txt.", vbOK & vbCritical, "FILE NOT FOUND!"
	End If 
	s_intTimeout = DEFAULT_TIMEOUT
	Set objFSO = Nothing
End Function 'Initialize