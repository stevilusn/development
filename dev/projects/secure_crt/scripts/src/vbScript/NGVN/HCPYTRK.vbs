'#$language = "VBScript"
'#$interface = "1.0"
'Name: HCPYTRK.vbs
Option Explicit

Dim s_objScriptTab
Dim s_objUtils
'Dim s_objDialogBox
Dim s_objStrings
'Dim s_objArrays
Dim s_objWatchFor
Dim s_intTimeout

crt.screen.Synchronous = True 

'********'*********'*********'*********'*********'*********'*********'*********'
'Main
'********'*********'*********'*********'*********'*********'*********'*********'
Sub Main
	Dim arrstrWaitForMAPCI
	Dim strCLLI
	Dim strCLLIClean
	Dim blnFoundCLLI
	Dim blnOneTry
	Dim blnExitMain
	Dim intMAPCIIndex
	
	If (Not Initialize) Then
		Exit Sub
	End If
 	s_objUtils.ScriptStart

	If (s_objUtils.SendAndWaitFor("leave all", ">", 10)) Then 'This tests if you have access to the screen.
 		If (getUserData(strCLLI, strCLLIClean) = vbOK) Then	
			arrstrWaitForMAPCI = Array("COMMAND DISALLOWED DURING DUMP", "CLLI NAME NOT FOUND", ">")
			blnFoundCLLI = False
			blnOneTry = False
			Do 
				blnExitMain = True 
				intMAPCIIndex = s_objUtils.SendAndWaitFor("MAPCI NODISP;MTC;TRKS;STAT;SELGRP " + strCLLIClean, _
						arrstrWaitForMAPCI, s_intTimeout)
				Select Case intMAPCIIndex
					Case 0 'None of the strings were found (timed out).
						crt.Dialog.MessageBox "Timed out waiting for screen data from MAPCI command.  Exiting script.", vbOKOnly & vbCritical
					Case 1 '"COMMAND_DISALLOWED_DURING_DUMP"
						s_objUtils.WaitForString ">", s_intTimeout
						s_objUtils.Send ""
						crt.Dialog.MessageBox   "SWITCH IS CURRENTLY IN IMAGE, PLEASE TRY LATER", _
								"COMMAND DISALLOWED DURING DUMP", vbOKOnly & vbInformation
						s_objUtils.WaitForString ">", s_intTimeout
						s_objUtils.Send "date"
					Case 2 '"CLLI NAME NOT FOUND"
						crt.Screen.WaitForString ">", s_intTimeout
'						s_objUtils.SendAndWaitFor "", ">", s_intTimeout 
						If (crt.Dialog.MessageBox("WOULD YOU LIKE THE SCRIPT TO SEARCH FOR THE CLLI?", _
								"CLLI NAME NOT FOUND", vbOkCancel) = vbOK) Then
							If (CLLICDR(strCLLIClean, blnFoundCLLI, blnOneTry)) then
								blnExitMain = False 
							End If 	
						Else
							s_objUtils.Send "leave all" & vbCr 
						End If
					Case 3 '">"
						s_objUtils.Send ""
						Call wrapUp(strCLLIClean,blnFoundCLLI)
					Case Else
						crt.Dialog.MessageBox "An unexpected value was returned while waiting for screen data from MAPCI command.  Exiting script.", vbOKOnly & vbCritical
				End Select
			Loop Until blnExitMain
		Else
			s_objUtils.Message "Canceled: getUserData"
		End If 	
	Else
		s_objUtils.CancelMsgBox
	End If 	

	s_objUtils.ScriptEnd
'	Set s_objDialogBox = Nothing
	Set s_objStrings = Nothing
'	Set s_objArrays = Nothing
	Set s_objWatchFor = Nothing
	Set s_objUtils = Nothing
End Sub 'Main

'********'*********'*********'*********'*********'*********'*********'*********'
'getUserData
'********'*********'*********'*********'*********'*********'*********'*********'
Function getUserData(ByRef strCLLI, ByRef strCLLIClean)
	strCLLI = crt.Dialog.Prompt("Enter CLLI:", "HCPYTRK Command")
	strCLLIClean = strCLLI
	If (Len(strCLLIClean) = 0) Then
		getUserData = vbCancel  
	Else
		getUserData = vbOK  
	End If 
End Function

'********'*********'*********'*********'*********'*********'*********'*********'
'CLLICDR
'********'*********'*********'*********'*********'*********'*********'*********'
Public Function CLLICDR(ByRef strCLLIClean, ByRef blnFoundCLLI, ByRef blnOneTry)
	Dim strTrunk
	
	s_objUtils.SendAndWaitFor "", ">", s_intTimeout
	
	CLLICDR = False 
	If (blnOneTry) Then 
		crt.Dialog.MessageBox _
				"UNABLE TO COMPLETE DISPLAY, PLEASE PERFORM HCPYTRK MANUALLY.", _
				"", vbOKOnly & vbInformation
	Else 
		strTrunk = s_objStrings.RemoveLeftNonNumericCharacters(strCLLIClean)			
		If (strTrunk <> "") Then 'NONSTANDARD. No numeric data was found
			'dblTrunkNumber = s_objUtils.CDblExt(strTrunk) 	'TODO: This is set but then never used.  Can it be removed?
		
			s_objUtils.SendAndWaitFor ";" + strTrunk + "->XX", ">", s_intTimeout
			crt.Screen.WaitForString ">", s_intTimeout
			s_objUtils.SleepSeconds 2
			
			If (getCLLI(strTrunk, strCLLIClean, blnFoundCLLI, blnOneTry)) Then
				CLLICDR = True 
			End If 
		Else 
			Call nonStandard
		End If 
	End If 
End Function 'CLLICDR

'********'*********'*********'*********'*********'*********'*********'*********'
'getCLLI
'********'*********'*********'*********'*********'*********'*********'*********'
Function getCLLI(strTrunk, ByRef strCLLIClean, ByRef blnFoundCLLI, ByRef blnOneTry)
	Dim intIndex
	Dim arrstrPatterns
	Dim blnNoCLLI
'	Dim intSize
		
	arrstrPatterns = Array(_
			"[A-Z][A-Z][A-Z][A-Z]" &  strTrunk, _
			"[A-Z][A-Z][A-Z]" & strTrunk, _
			"[A-Z][A-Z]" & strTrunk, _
			"[A-Z]" & strTrunk, _
			"BOTTOM", _
			">")

	s_objWatchFor.Reset
	s_objWatchFor.StringToSend  = "TABLE CLLICDR;LIS 1 (2 EQ XX)"
	s_objWatchFor.Patterns = arrstrPatterns
	s_objWatchFor.UseRegExp = True
	s_objWatchFor.IgnoreCase = False
	s_objWatchFor.Global = False
	s_objWatchFor.TimeoutSeconds = 60
	intIndex = s_objWatchFor.Execute
	Select Case intIndex
		Case 0 
			blnFoundCLLI = False
			crt.Dialog.Messagebox "Timed out watching for patterns after " _
					& s_objWatchFor.TimeoutSeconds _
					& " seconds.  Patterns:" & vbCrLf _
					& Join(arrPatterns, vbCrLf) & vbCrLf _
					& "Will continute processing now.", _
					"TIMEOUT", vbOKOnly & vbInformation 
			s_objUtils.Send ""
		Case 1, 2, 3, 4 '"[A-Z]" & strTrunk or "[A-Z][A-Z]" & strTrunk or "[A-Z][A-Z][A-Z]" & strTrunk or "[A-Z][A-Z][A-Z][A-Z]" &  strTrunk
'TODO: this didn't do anything.  Verify.
'			Select Case inIndex
'				Case 1
'					intSize = 4
'				Case 2
'					intSize = 3
'				Case 3
'					intSize = 2
'				Case 4
'					intSize = 1
'			End Select 
'			strCLLIClean = Right(s_objWatchFor.MatchValue, Len(strCLLIClean) + intSize)
			strCLLIClean = s_objWatchFor.MatchValue
			blnFoundCLLI = True
			blnOneTry = True 
		Case 5 '"BOTTOM"
			blnFoundCLLI = False
			blnNoCLLI = True
		Case 6 '">"
			blnFoundCLLI = False
			s_objUtils.Send ""
	End Select
	crt.Screen.WaitForString ">", s_intTimeout
'	s_objUtils.SendAndWaitFor "", ">", s_intTimeout
	s_objUtils.SleepSeconds 2

	If (blnNoCLLI) Then 
		crt.Dialog.Messagebox "CLLI IS NOT BUILT IN THE SWITCH", _
				"NO CLLI", vbOKOnly & vbInformation
		s_objUtils.Send ""
		getCLLI = False 
	Else 
		blnFoundCLLI = False
		getCLLI = True  
	End If 
End Function 'getCLLI

'********'*********'*********'*********'*********'*********'*********'*********'
'nonStandard
'********'*********'*********'*********'*********'*********'*********'*********'
Public Sub nonStandard()
	s_objUtils.Send ""
	crt.Dialog.Messagebox "PLEASE RESEARCH THE CLLI MANUALLY.", _
			"CLLI NOT FOUND", vbOKOnly & vbInformation
	s_objUtils.Send ""
End Sub

'********'*********'*********'*********'*********'*********'*********'*********'
'wrapUp
'********'*********'*********'*********'*********'*********'*********'*********'
Sub wrapUp(strCLLIClean, blnFoundCLLI)
 	Dim strCommand
	Dim blnSIPTrunkGroup
	Dim arrstrHCPYTRKStrings
	
	blnSIPTrunkGroup = False 
				 				 
	s_objUtils.WaitForString ">", s_intTimeout
	s_objUtils.SendAndWaitFor strCommand, ">", s_intTimeout	'TODO: strCommand is never set, so Variable is not used.  Can it be removed.  This currently is doing the same as a 's_objUtils.Send "" ', I think.
	s_objUtils.SleepSeconds 2
	
	arrstrHCPYTRKStrings = Array("DPT group", ">")				
	Select Case s_objUtils.SendAndWaitFor("HCPYTRK; QUIT ALL", arrstrHCPYTRKStrings, s_intTimeout)
		Case 0 'None of the strings were found (timed out).
			crt.Dialog.MessageBox "Timed out waiting for screen data from HCPYTRK command.  Exiting script.", vbOKOnly & vbCritical
	    Case 1 '"DPT group"
	    	'TODO: Test if need to wait for the ">" after finding the "DPT group".  Synchronous=True would mean that the screen capture would stop afer the "p".  Or, by sending a new "Send" command, will SecureCRT know to finish up?
	    	blnSIPTrunkGroup = True
	    	s_objUtils.Send "mapci prtmap nodisp;mtc;trks;dptrks;post g " _
	    			& strCLLIClean & ";printmap;leave all"
	    Case 2 '">"
	    	s_objUtils.Send ""
		Case Else
			crt.Dialog.MessageBox "An unexpected value was returned while waiting for screen data from HCPYTRK command.  Exiting script.", vbOKOnly & vbCritical
	End Select

	If (blnFoundCLLI) Then 
		crt.Dialog.MessageBox  "    THE CLLI IS BUILT AS " + strCLLI, _
				"CLLI BUILT", vbOKOnly & vbInformation
	End If 	
	
	s_objUtils.Send ""
	If (blnSIPTrunkGroup) Then 
		crt.Dialog.MessageBox  "This Trunk Group is a SIP Trunk Group.  ", _
				"SIP TRUNK GROUP", vbOKOnly & vbInformation
	End If 
End Sub


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
		MsgBox "Could not find file " & strCurPath & "BootStrap.txt.", _
				vbOKOnly & vbCritical, "FILE NOT FOUND!"
	End If 
	s_intTimeout = DEFAULT_TIMEOUT
	Set objFSO = Nothing
End Function 'Initialize