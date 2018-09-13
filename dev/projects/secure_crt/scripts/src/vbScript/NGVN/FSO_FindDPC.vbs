'#$language = "VBScript"
'#$interface = "1.0"
'Name: FindDPC.vbs

Option Explicit

Dim s_intTimeout
Dim s_objScriptTab	'Gets set during Initialize.
Dim s_objUtils
Dim s_objStrings
Dim s_objWatchFor

Dim s_intNOTAVALIDPOINTCODE
s_intNOTAVALIDPOINTCODE = -1

'********'*********'*********'*********'*********'*********'*********'*********'
'Main
'********'*********'*********'*********'*********'*********'*********'*********'
Sub Main
	Dim strResult
	Dim strPointCode
	 	
	If (Not Initialize) Then Exit Sub
	s_objUtils.ScriptStart
	
	If (s_objUtils.LeaveAll(s_intTimeout)) Then 'This tests if you have access to the screen.
		strPointCode = Trim(crt.Dialog.Prompt("Point Code (i.e. 253-193-0):", "Search for Routeset", ""))
		If (Len(strPointCode) > 0) Then 
			If (getRteSetFromDPC(strPointCode, strResult) = False Or strResult = "") Then
				crt.Dialog.MessageBox "Point Code " & strPointCode & " not found.", "CODE NOT FOUND", vbOKOnly & vbExclamation
			End If 
		End If
	Else 
		s_objUtils.CancelMsgBox
	End If
	
	ExitMain
End Sub 'Main
Sub ExitMain()
	s_objUtils.ScriptEnd
	Set s_objUtils = Nothing
	Set s_objStrings = Nothing
	Set s_objWatchFor = Nothing
	Set s_objScriptTab = Nothing 
End Sub 'ExitMain
	
'********'*********'*********'*********'*********'*********'*********'*********'
'pointCodeStrToInt
'********'*********'*********'*********'*********'*********'*********'*********'
Private Function pointCodeStrToInt(strPC) 'returns integer
	Dim intPosition
	Dim strWork
	Dim strNet
	Dim strCluster
	Dim strMember
	
	strWork = Trim(strPC)

	If (Len(strWork) > 11) Then
		pointCodeStrToInt = s_intNOTAVALIDPOINTCODE
		Exit Function 
	End If 

	intPosition = InStr(strWork, "-")
	If (intPosition <= 1 Or intPosition > 4) Then
		pointCodeStrToInt =  s_intNOTAVALIDPOINTCODE
		Exit Function 
	End If 
	
	strNet = Left(strWork, intPosition - 1)
	strWork = Mid(strWork, intPosition + 1, Len(strWork))
	intPosition = InStr(strWork, "-")
	If (intPosition <= 1 Or intPosition > 4) Then
		pointCodeStrToInt = s_intNOTAVALIDPOINTCODE
		Exit Function 
	End If 
	
	strCluster = Left(strWork, intPosition - 1)
	strMember = Mid(strWork, intPosition + 1, Len(strWork))
	If (Not s_objUtils.IsPosInteger(strNet) Or Not s_objUtils.IsPosInteger(strCluster) Or Not s_objUtils.IsPosInteger(strMember)) Then
		pointCodeStrToInt = s_intNOTAVALIDPOINTCODE
		Exit Function 
	End If 

	If (s_objStrings.IntVal(strNet) > 255 Or s_objStrings.IntVal(strCluster) > 255 Or s_objStrings.IntVal(strMember) > 255) Then 
		pointCodeStrToInt = s_intNOTAVALIDPOINTCODE
		Exit Function 
	End If 
	
	pointCodeStrToInt = (s_objStrings.IntVal(strNet) * 256 * 256) + (s_objStrings.IntVal(strCluster) * 256) + s_objStrings.IntVal(strMember)
End Function 'pointCodeStrToInt

'********'*********'*********'*********'*********'*********'*********'*********'
'getPointCodeNet
'********'*********'*********'*********'*********'*********'*********'*********'
Private Function getPointCodeNet(intPointCode) 'returns integer
	'Perform a "Bitwise And"
	getPointCodeNet = ((intPointCode AND (255*256*256)) / (256*256))
End Function 'getPointCodeNet

'********'*********'*********'*********'*********'*********'*********'*********'
'getPointCodeCluster
'********'*********'*********'*********'*********'*********'*********'*********'
Private Function getPointCodeCluster(intPointCode) 'returns integer
	getPointCodeCluster = ((intPointCode AND (255*256)) / (256))
End Function 'getPointCodeCluster

'********'*********'*********'*********'*********'*********'*********'*********'
'getPointCodeMember
'********'*********'*********'*********'*********'*********'*********'*********'
Private Function getPointCodeMember(intPointCode) 'returns integer
	getPointCodeMember = (intPointCode AND (255))
End Function 'getPointCodeMember

'********'*********'*********'*********'*********'*********'*********'*********'
'buildPcSearchStr
'********'*********'*********'*********'*********'*********'*********'*********'
Private Function buildPcSearchStr(intPointCode) 'returns string
	Dim strRetVal
	
	strRetVal = ""
	If (intPointCode And &HFF000000) = 0 then 
		strRetVal = "ANSI7 (" & CStr(getPointCodeNet(intPointCode)) & ") (" & CStr(getPointCodeCluster(intPointCode)) & ") (" & CStr(getPointCodeMember(intPointCode)) & ")"
	End If 
	buildPcSearchStr = strRetVal
End Function ' buildPcSearchStr

'********'*********'*********'*********'*********'*********'*********'*********'
'getRteSetFromDPC
'********'*********'*********'*********'*********'*********'*********'*********'
Private Function getRteSetFromDPC(strPC, ByRef strRetVal) 'returns string
	Dim intPointCode
	Dim strSearch
	Dim arrstrPatterns
	Dim intIndex
	Dim strMatch

	strRetVal = ""
	getRteSetFromDPC = True 

	intPointCode = pointCodeStrToInt(strPC)

	If (intPointCode = s_intNOTAVALIDPOINTCODE) Then
		crt.Dialog.MessageBox strPC & " is not a valid Point Code.", "INVALID CODE", vbOKOnly & vbExclamation
		getRteSetFromDPC = False 
	Else 
		strSearch = buildPcSearchStr(intPointCode)

		s_objUtils.SendAndWaitFor "send sink;table c7rteset;format pack;send previous", ">", 5
		s_objUtils.SendAndWaitFor "count (4 eq '" & strSearch & "')", ">", 5

		' Place code here to determine how many tuples qualify and enumerate them.
		
		arrstrPatterns = Array(_
			"[0-9A-Za-z]+ [0-9A-Za-z]+ [0-9A-Za-z]+ " & Replace(Replace(strSearch, "(", "\("), ")", "\)"), _
			"BOTTOM")
		s_objWatchFor.Reset
		s_objWatchFor.Patterns = arrstrPatterns
		s_objWatchFor.UseRegExp = True
		s_objWatchFor.IgnoreCase = True
		s_objWatchFor.Global = False
		s_objWatchFor.TimeoutSeconds = 60
		s_objWatchFor.StringToSend = "lis all (4 eq '" & strSearch & "')"
		's_objWatchFor.StringToSend = "lis 1 (4 eq '" & strSearch & "')"
		s_objWatchFor.IgnoreStringToSend = True
		intIndex = s_objWatchFor.Execute
		strMatch = s_objWatchFor.MatchValue
		Select Case intIndex
			Case 0 
				crt.Dialog.Messagebox "Timed out watching for patterns after " _
						& s_objWatchFor.TimeoutSeconds _
						& " seconds.  Patterns:" & vbCrLf _
						& Join(arrstrPatterns, vbCrLf) & vbCrLf _
						& "Will continute processing now.", _
						"TIMEOUT", vbOKOnly & vbInformation 
				getRteSetFromDPC = False
			Case 1
				strRetVal = Left(strMatch, InStr(strMatch, " ") - 1)
				getRteSetFromDPC = True
			Case 2 '"BOTTOM"
				getRteSetFromDPC = False
		End Select	
		s_objUtils.LeaveAll(s_intTimeout)
	End If
End Function 'getRteSetFromDPC

'********'*********'*********'*********'*********'*********'*********'*********'
'isvalidpointcode
'This was commented out in the original Crosstalk script
'********'*********'*********'*********'*********'*********'*********'*********'
'Private function isvalidpointcode(strPC) 'returns boolean
'	Dim intPosition
'	Dim strWork
'	Dim strNet
'	Dim strCluster
'	Dim strMember
'	
'	strWork = Trim(strPC)
'	
'	If (Len(strWork) > 11) Then
'		isvalidpointcode = False
'		Exit Function 
'	End If 	
'	
'	intPosition = InStr(strWork, "-")
'	
'	If (intPosition = 0 Or intPosition > 4) Then
'		isvalidpointcode = False
'		Exit Function 
'	End If 	
'	
'	strNet = Left(strWork, intPosition - 1)
'	strWork = Mid(strWork, intPosition + 1, Len(strWork))
'	
'	intPosition = InStr(strWork, "-")
'	
'	If (intPosition = 0 Or intPosition > 4) Then
'		isvalidpointcode = False
'		Exit Function 
'	End If 	
'	
'	strCluster = left(strWork, intPosition - 1)
'	
'	strMember = Mid(strWork, intPosition + 1, Len(strWork))
'	
'	If (Not s_objUtils.IsPosInteger(strNet) Or Not s_objUtils.IsPosInteger(strCluster) Or Not s_objUtils.IsPosInteger(strMember)) Then
'		isvalidpointcode = False
'		Exit Function 
'	End If 	
'	
'	if (s_objStrings.IntVal(strNet) > 255 Or s_objStrings.IntVal(strCluster) > 255 Or s_objStrings.IntVal(strMember) > 255) Then
'		isvalidpointcode = False
'		Exit Function 
'	End If 
'		
'	isvalidpointcode = True
'End Function 


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

