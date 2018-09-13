'********'*********'*********'*********'*********'*********'*********'*********'
'FileName:	clsSecureCRTUtils.vbs
'Author:	Kent Garvis
'Purpose:	This class acts as a container for generic utility functions 
'	specifically used with SecureCRT.
'Dependency:
'	This Class is dependant on the Constants.txt being included/imported first.
'	This Class can only be used from within a SecureCRT script.
'	This Class is not intened to be used directly by SecureCRT, but to
'		be included/imported into and used by other SecureCRT scripts at runtime.
'Properties:
'  	Tab - The SecureCRT Tab window object to interact with.
'Methods:
'	...
'
'Note: This Class will default to interacting with the Tab that the script was
'	called from.  But you can programically open another Tab connection and 
'	then Set a new instance of this class for use with the new Tab by using 
'	the Tab property. 
'********'*********'*********'*********'*********'*********'*********'*********'
Option Explicit

Class clsSecureCRTUtils
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'Initialize and Terminate the class.
	'********'*********'*********'*********'*********'*********'*********'*********'
	Private Sub Class_Initialize   ' Setup Initialize event.
		Set c_objTab = crt.GetScriptTab
		c_strOpenFileMethod = "NotePad"
		c_blnOpenFileImmediately = False
		c_intDebugLevel = DEBUG_LEVEL
		c_blnIgnoreCase = False 
	End Sub 
	Private Sub Class_Terminate()    ' Setup Terminate event.
		Set c_objTab = Nothing 
	End Sub 	
		
	'********'*********'*********'*********'*********'*********'*********'*********'
	'IgnoreCase
	'Perform a case insensitive pattern match.
	'********'*********'*********'*********'*********'*********'*********'*********'
	Private c_blnIgnoreCase
	Public Property Let IgnoreCase(blnIgnoreCase)
		c_blnIgnoreCase = blnIgnoreCase
'		c_objTab.Screen.IgnoreCase = blnIgnoreCase
	End Property
	Public Property Get IgnoreCase
		IgnoreCase = c_blnIgnoreCase
	End Property 	
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'DebugLevel
	'********'*********'*********'*********'*********'*********'*********'*********'
	Private c_intDebugLevel
	Public Property Let DebugLevel(intDebugLevel)
		c_intDebugLevel = intDebugLevel
	End Property
	Public Property Get DebugLevel
		DebugLevel = c_intDebugLevel
	End Property 

	'******************************************************************************
	'CancelMsgBox
	'Cancel Message Box is a conviniance sub to display a generic messagebox warning 
	'the user that processing can not continue.
	'******************************************************************************
	Public Sub CancelMsgBox()
		crt.Dialog.Messagebox _
				"Can not continue prosessing." _
				& vbCrLf & vbCrLf _
				& "Check your connection and try again.", _
				"CANCELED", vbCritical + vbOKOnly
	End Sub 'CancelMsgBox
	
	'******************************************************************************
	'CDbl Extended.  
	'CDblExt returns a Double; it evaluates strString for its 
	'numerical meaning and returns that meaning as the result. Leading 
	'white-space characters are ignored, and string is evaluated until a 
	'non-numeric character is encountered.
	'******************************************************************************
	Public Function CDblExt(strString)
		Dim i
		Dim strChr
		Dim intAsc
		Dim strNewString
		Dim lngString
		
		For i = 1 To Len(strString) Step 1
			strChr = Left(strString, 1)
			intAsc = Asc(strChr)
			If (Asc(strChr) <= 32) Then
				strString = Mid(strString, 2)
			Else 
				Exit For
			End If 
		Next
		strString = Trim(strString)
		
		strChr = Left(strString, 1)
		If (InStr("+-", strChr) > 0) Then
			strNewString = Left(strString, 1)
			strString = Mid(strString, 2)
		End If
		For i = 1 To Len(strString) Step 1
			strChr = Mid(strString, i, 1)
			If (Asc(strChr) >= 48 And Asc(strChr) <= 57) Then
				strNewString = strNewString & strChr
			Else 
				Exit For
			End If
		Next	
		If (IsNumeric(strNewString)) Then 
			lngString = CDbl(strNewString)
		Else
			lngString = 0
		End If 
		CDblExt = lngString
	End Function 'CDblExt

	'******************************************************************************
	'CleanTrashMsg
	'Clean Trash Message is a conviniance function to display a generic message 
	'about cleaning text entry.
	'******************************************************************************
	Function CleanTrashMsg(strLabel)
		CleanTrashMsg =    "Enter Exact Information As Requested " _
				& vbCrLf & "Enter " & strLabel & "." _
				& vbCrLf & "." _
				& vbCrLf & vbCrLf
	End Function 'CleanTrashMsg
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'Connected
	'********'*********'*********'*********'*********'*********'*********'*********'
	Public Function Connected()
		Dim blnConnected
		
		blnConnected = c_objTab.session.connected
		If (Not blnConnected) Then
	        MsgBox "Not Connected.  Please connect before running this script.", _
	        		vbOK + vbExclamation, "NOT CONNECTED!"
        End If 
		Connected = blnConnected
	End Function 'Connected
		
	'********'*********'*********'*********'*********'*********'*********'*********'
	'DebugEcho
	'This function will send a Debug echo command to the screen and wait for the prompt.
	'********'*********'*********'*********'*********'*********'*********'*********'
	Function DebugEcho(strMsg, intDebugLevel)
		If (DebugLevel >= intDebugLevel) Then 
			strMsg = "echo 'DEBUG " & g_strDebugLocator & ": " & strMsg & "'"
			Send strMsg
			DebugEcho = WaitFor(DEFAULT_LINE_TERMINATOR, DEFAULT_TIMEOUT)
		End if
	End Function 'DebugEcho
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'Echo
	'This function will send an echo command to the screen and wait for the prompt.
	'********'*********'*********'*********'*********'*********'*********'*********'
	Function Echo(strMsg)
		Send "echo '" & strMsg & "'"
		Echo = WaitFor(DEFAULT_LINE_TERMINATOR, DEFAULT_TIMEOUT)
	End Function 'Echo
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'getSessionFileName
	'********'*********'*********'*********'*********'*********'*********'*********'
	Public Function getSessionFileName()
		Dim objConfig
		Dim strUsername, strSessionName
		Set objConfig = c_objTab.Session.Config
		strUsername = objConfig.GetOption("Username")
		strSessionName = c_objTab.Session.Path
		Set objConfig = Nothing 
		getSessionFileName = Replace(Replace(Replace(Replace(strSessionName, "/", "_"), "\", "_"), "-", "_"), " ", "") + ".txt"
	End Function 'getSessionFileName
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'LeaveAll
	'TODO: Used in Buildmems.xws, c7addmems.xws, FindDPC.xws, HCPYTRK.XWS, omshow.xws, posttrkgrp.xws
	'********'*********'*********'*********'*********'*********'*********'*********'
	Function LeaveAll(intTimeout)
		LeaveAll = SendAndWaitFor("leave all", ">", intTimeout)
	End Function 'LeaveAll

	'********'*********'*********'*********'*********'*********'*********'*********'
	'QuitAll
	'TODO: Used in Buildmems.xws, c7addmems.xws, FindDPC.xws, HCPYTRK.XWS, omshow.xws, posttrkgrp.xws
	'********'*********'*********'*********'*********'*********'*********'*********'
	Function QuitAll(intTimeout)
		QuitAll = SendAndWaitFor("quit all", ">", intTimeout)
	End Function 'QuitAll

	'********'*********'*********'*********'*********'*********'*********'*********'
	'Message
	'Friendly name for SetStatusText.  Allows setting the SecureCRT status bar text.
	'SetStatusText is specifically for setting the SecureCRT status bar text.  The
	'idea behind Message is that it could be modified as needed to send Messages by 
	'other means (MsgBox, log file, etc.).
	'********'*********'*********'*********'*********'*********'*********'*********'
	Public Sub Message(strMessage)
		StatusText = strMessage
	End Sub 'Message

	'********'*********'*********'*********'*********'*********'*********'*********'
	'MessageBox
	'Wrapper for crt.Dialog.MessageBox.  No added features.  But using this will 
	'allow the Main script file to be run outside of SecureCRT without generating 
	'an error for the crt opject.
	'********'*********'*********'*********'*********'*********'*********'*********'
	Public Sub MessageBox(strMessage, strTitle, intButtons)
		MessageBox = crt.Dialog.MessageBox(strMessage, strTitle, intButtons)
	End Sub 'Message
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'OpenFile
	'Set the OpenFileMethod property and then call this to open the specifed file.
	'********'*********'*********'*********'*********'*********'*********'*********'
	Public Sub OpenFile(strFileName)
		Dim blnOpenFile
		
		blnOpenFile = true
		If (Not c_blnOpenFileImmediately) Then
			blnOpenFile = (MsgBox("Would you like to display the file """ & strFileName & """?", vbYesNo) = vbYes)
		End if
		If (blnOpenFile) Then
			Select Case LCase(OpenFileMethod)
			    Case "notepad"
			    	Run("NotePad.exe " & strFileName)
			    Case "ie"
'FIXME: Change the hard mapped path to Windows Default pathing.
			    	Run """C:\Program Files\Internet Explorer\IExplore.exe"" """ & strFileName & """"
			    Case Else
			    	MsgBox "The method """ & OpenFileMethod & """ to open the file is not recognized.", vbOKOnly & vbExclamation, "Unrecognized Method"
			End Select	
		End If
	End Sub 'OpenFile
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'OpenFileImmediately
	'Used by the OpenFile Sub.  OpenFileImmediately = True will bypasse the message  
	'box asking if the file should be opened.  
	'********'*********'*********'*********'*********'*********'*********'*********'
	Private c_blnOpenFileImmediately
	Public Property Let OpenFileImmediately(blnOpenFileImmediately)
		c_blnOpenFileImmediately = blnOpenFileImmediately
	End Property
	Public Property Get OpenFileImmediately
		OpenFileImmediately =c_blnOpenFileImmediately
	End Property 'OpenFileImmediately
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'OpenFileInIE
	'********'*********'*********'*********'*********'*********'*********'*********'
	Public Sub OpenFileInIE(strFileName)
		OpenFileMethod = "IE"
		OpenFile strFileName
	End Sub 'OpenFileInIE			
		
	'********'*********'*********'*********'*********'*********'*********'*********'
	'OpenFileInNotepad
	'********'*********'*********'*********'*********'*********'*********'*********'
	Public Sub OpenFileInNotepad(strFileName)
		OpenFileMethod = "Notepad"
		OpenFile strFileName
	End Sub 'OpenFileInNotepad
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'OpenFileMethod
	'********'*********'*********'*********'*********'*********'*********'*********'
	Private c_strOpenFileMethod
	Public Property Let OpenFileMethod(strOpenFileMethod)
		c_strOpenFileMethod = strOpenFileMethod
	End Property
	Public Property Get OpenFileMethod
		OpenFileMethod = c_strOpenFileMethod
	End Property 
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'QDNDataParse
	'********'*********'*********'*********'*********'*********'*********'*********'
	Public Function QDNDataParse(strNumber, ByRef dicstrQDNData, ByRef dicblnFlags, ByRef dicblnFeatures, blnDisplayMessages)
		Dim arrScreenData
		Dim strScreenRow
		Dim intPos 	
	
		SendAndWaitFor "leave all", ">", s_intTimeout
		arrScreenData = SendAndReadStrings("QDN " & strNumber, ">", s_intTimeout)
	
		For Each strScreenRow In arrScreenData
			If (InStr(strScreenRow, "UNASSIGNED") > 0) Then 
				If (blnDisplayMessages) Then  
					crt.Dialog.MessageBox "The DN is unassigned in the switch.", "ENDING EXECUTION", vbOKOnly & vbInformation
				End If 
				QDNDataParse = False 
				Exit Function 
			End if
			If (InStr(strScreenRow, "INVALID FOR THIS OFFICE") > 0) Then    
				If (blnDisplayMessages) Then  
					crt.Dialog.MessageBox "The DN is not valid for this switch.", "ENDING EXECUTION", vbOKOnly & vbInformation
				End If 
				QDNDataParse = False 
				Exit Function 
			End If
			
			'Parse Data
			If (InStr(strScreenRow, "BNV:") > 0) Then
				intPos = InStr(strScreenRow, " :")
				dicstrQDNData("BNV") = Mid(strScreenRow, intPos + 50, 2)
			End If 
			If (InStr(strScreenRow, "CARDCODE:") > 0) Then
				intPos = InStr(strScreenRow, " : ")
				dicstrQDNData("CARDCODE") = Mid(strScreenRow, intPos + 12, 6)
			End If 
			If (InStr(strScreenRow, "CUSTGRP:") > 0) Then
				intPos = InStr(strScreenRow, " : ")
				dicstrQDNData("CUSTGRP") = Mid(strScreenRow, intPos + 14, 10)
			End If 
			If (InStr(strScreenRow, "CFW INDEX:") > 0) Then
				intPos = InStr(strScreenRow, " :")
				dicstrQDNData("CFW INDEX") = Mid(strScreenRow, intPos + 11, 6)
			End If 
			If (InStr(strScreenRow, "DN:") > 0) Then    
				intPos = InStr(strScreenRow, ":")
				dicstrQDNData("DN") = Mid(strScreenRow, intPos + 6, 7)
			End If 
			If (InStr(strScreenRow, "GND:") > 0) Then
				intPos = InStr(strScreenRow, " :")
				dicstrQDNData("GND") = Mid(strScreenRow, intPos + 27, 2)
			End If 
			If (InStr(strScreenRow, "IBN TYPE:") > 0) Then
				intPos = InStr(strScreenRow, ":")
				if (intPos = 9) Then 
					dicstrQDNData("IBN TYPE") = Mid(strScreenRow, intPos + 1, 15)
				End If 
			End If 
			If (InStr(strScreenRow, "LINE CLASS CODE:") > 0) Then                           
				intPos = InStr(strScreenRow, " : ")
				dicstrQDNData("LINE CLASS CODE") = Mid(strScreenRow, intPos + 23, 15)
			End If 
			If (InStr(strScreenRow, "LINE EQUIPMENT NUMBER:") > 0) Then
				intPos = InStr(strScreenRow, " : ")
				dicstrQDNData("LINE EQUIPMENT NUMBER") = Mid(strScreenRow, intPos + 28, 25)
			End If 
			If (InStr(strScreenRow, "LINE TREATMENT GROUP:") > 0) Then
				intPos = InStr(strScreenRow, " : ")
				dicstrQDNData("LINE TREATMENT GROUP") = Mid(strScreenRow, intPos + 27, 6)
			End If 
			If (InStr(strScreenRow, "LNATTIDX:") > 0) Then
				intPos = InStr(strScreenRow, " : ")
				dicstrQDNData("LNATTIDX") = Replace(Mid(strScreenRow, intPos + 34, 15), " ", "")
			End If 
			If (InStr(strScreenRow, "MNO:") > 0) Then
				intPos = InStr(strScreenRow, " :")
				dicstrQDNData("MNO") = Mid(strScreenRow, intPos + 58, 2)
			End If 
			If (InStr(strScreenRow, "NCOS:") > 0) Then
				intPos = InStr(strScreenRow, " : ")
				dicstrQDNData("NCOS") = Mid(strScreenRow, intPos + 47, 3)
			End If 
			If (InStr(strScreenRow, "PADGRP:") > 0) Then
				intPos = InStr(strScreenRow, " :")
				dicstrQDNData("PADGRP") = Mid(strScreenRow, intPos + 38, 5)
			End If 
			If (InStr(strScreenRow, "PM NODE NUMBER     :") > 0) Then
				intPos = InStr(strScreenRow, " :")
				dicstrQDNData("PM NODE NUMBER") = Mid(strScreenRow, intPos + 6, 6)
			End If 
			If (InStr(strScreenRow, "PM TERMINAL NUMBER :") > 0) Then
				intPos = InStr(strScreenRow, " :")
				dicstrQDNData("PM TERMINAL NUMBER") = Mid(strScreenRow, intPos + 6, 6)
			End If 
			If (InStr(strScreenRow, "PORTED-IN") > 0) Then
				intPos = InStr(strScreenRow, "-")
				dicstrQDNData("PORTED-IN") = Mid(strScreenRow, intPos - 7, 11)
			End If 
			If (InStr(strScreenRow, "SIG:") > 0) Then
				intPos = InStr(strScreenRow, " : ")
				dicstrQDNData("SIG") = Mid(strScreenRow, intPos + 18, 3)
			End If 
			If (InStr(strScreenRow, "SNPA:") > 0) Then
				intPos = InStr(strScreenRow, ":")
				dicstrQDNData("SNPA") = Mid(strScreenRow, intPos + 2, 3)
			End If 
			If (InStr(strScreenRow, "SUBGRP:") > 0) Then     
				intPos = InStr(strScreenRow, " : ")
				dicstrQDNData("SUBGRP") = Mid(strScreenRow, intPos + 36, 10)
			End If 
			If (InStr(strScreenRow, "TYPE:") > 0) Then
				intPos = InStr(strScreenRow, ":")
				If (intPos = 5) Then
				  dicstrQDNData("TYPE") = Mid(strScreenRow, intPos + 2, 20)
				End If 
			End If 	
			
			'Parse Features. 			
			If (InStr(strScreenRow, "3WC") > 0 And InStr(strScreenRow, "U3WC") = 0) Then
				dicblnFlags("3WC") = True
				dicblnFeatures("3WC") = True 
			End If
			If (InStr(strScreenRow, "ACB") > 0) Then
				dicblnFlags("ACB") = True
				dicblnFeatures("ACB") = True
			End if
			If (InStr(strScreenRow, "ACRJ") > 0) Then
				dicblnFlags("ACRJ") = True
				dicblnFeatures("ACRJ") = True 
			End if
			If (InStr(strScreenRow, "AR ") > 0) Then
				dicblnFlags("AR") = True
				dicblnFeatures("AR") = True
			End if
			If (InStr(strScreenRow, "CCW") > 0) Then
				dicblnFlags("CCW") = True
				dicblnFeatures("CCW") = True
			End if
			If (InStr(strScreenRow, "CFBL") > 0) Then
				dicblnFlags("CFBL") = True
				dicblnFeatures("CFBL") = True 
			End if
			If (InStr(strScreenRow, "CFDA") > 0) Then
				dicblnFlags("CFDA") = True
				dicblnFeatures("CFDA") = True
			End if
			If (InStr(strScreenRow, "CFW C") > 0) Then
				dicblnFlags("CFW") = True
				dicblnFeatures("CFW") = True
			End if
			If (InStr(strScreenRow, "CFW F") > 0) Then
				dicblnFlags("CFW-FIXED") = True
				dicblnFeatures("CFW-FIXED") = True
			End If
			If (InStr(strScreenRow, "CNAMD") > 0) Then
				dicblnFlags("CNAMD") = True
				dicblnFeatures("CNAMD") = True
			End if
			If (InStr(strScreenRow, "CND NOAMA") > 0) Then
				dicblnFlags("CND") = True
				dicblnFeatures("CND") = True
			End if
			If (InStr(strScreenRow, "CNDB") > 0) Then
				dicblnFlags("CNDB") = True
				dicblnFeatures("CNDB") = True 
			End if
			If (InStr(strScreenRow, "COT") > 0) Then
				dicblnFlags("COT") = True
				dicblnFeatures("COT") = True
			End if
			If (InStr(strScreenRow, "CWT") > 0) Then
				dicblnFlags("CWT") = True
				dicblnFeatures("CWT") = True
			End if
			If (InStr(strScreenRow, "DGT") > 0) Then
				dicblnFlags("DGT") = True
				dicblnFeatures("DGT") = True
			End if
			If (InStr(strScreenRow, "INTPIC") > 0) Then
				dicblnFlags("INTPIC") = True
				dicblnFeatures("INTPIC") = True 
			End if
			If (InStr(strScreenRow, "LPIC") > 0) Then
				dicblnFlags("LPIC") = True
				dicblnFeatures("LPIC") = True 
			End if
			If (InStr(strScreenRow, "LSPAO") > 0) Then
				dicblnFlags("LSPAO") = True
				dicblnFeatures("LSPAO") = True 
			End if
			If (InStr(strScreenRow, "MTZ") > 0) Then
				dicblnFlags("MTZ") = True
				dicblnFeatures("MTZ") = True
			End if
			If (InStr(strScreenRow, "MWT") > 0) Then
				dicblnFlags("MWT") = True
				dicblnFeatures("MWT") = True
			End if
			If (InStr(strScreenRow, " PIC") > 0) Then
				dicblnFlags("PIC") = True
				dicblnFeatures("PIC") = True
			End if
			If (InStr(strScreenRow, "PORT ") > 0) Then
				dicblnFlags("PORT") = True
				dicblnFeatures("PORT") = True 
			End if
			If (InStr(strScreenRow, "SC1") > 0) Then
				dicblnFlags("SC1") = True
				dicblnFeatures("SC1") = True
			End if
			If (InStr(strScreenRow, "SCA") > 0) Then
				dicblnFlags("SCA") = True
				dicblnFeatures("SCA") = True
			End if
			If (InStr(strScreenRow, "SCF") > 0) Then
				dicblnFlags("SCF") = True
				dicblnFeatures("SCF") = True
			End if
			If (InStr(strScreenRow, "SCRJ") > 0) Then
				dicblnFlags("SCRJ") = True
				dicblnFeatures("SCRJ") = True 
			End if
			If (InStr(strScreenRow, "SCWID") > 0) Then
				dicblnFlags("SCWID") = True
				dicblnFeatures("SCWID") = True 
			End if
			If (InStr(strScreenRow, "SPB") > 0) Then
				dicblnFlags("SPB") = True
				dicblnFeatures("SPB") = True
			End if	
		
		Next 
		QDNDataParse = True 
	End Function 'QDNDataParse
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'ReadMwqDn
	'********'*********'*********'*********'*********'*********'*********'*********'
	Public Function ReadMwqDn()
		Dim objFSO
		Dim strFile
		Dim objTextFile
		Dim strData
		
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		strFile = DIR_NGVN_DATA & DEFAULT_MWQ_DN_FILE
		If (objFSO.FileExists(strFile)) Then 
			Set objTextFile = objFSO.OpenTextFile(strFile, ForReading)
		
			Do Until objTextFile.AtEndOfStream
				strData = objTextFile.Readline
			Loop
			ReadMwqDn = strData
		Else
			ReadMwqDn = ""
		End If 
	End Function 'ReadMwqDn
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'RemoveFirstAndLastRow
	'********'*********'*********'*********'*********'*********'*********'*********'
	Public Function RemoveFirstAndLastRow(arrData)
		arrData(LBound(arrData)) = ""
		arrData(UBound(arrData)) = ""
		strData = Trim(Join(arrData, vbCrLf))
		RemoveFirstAndLastRow = Split(strData, vbCrLf)
	End Function
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'Run
	'********'*********'*********'*********'*********'*********'*********'*********'
	Public Sub Run(strCommand)
		Dim objShell
		Set objShell = CreateObject("WScript.Shell")
		objShell.Run strCommand
		crt.Sleep 800 
		Set objShell = Nothing 
	End Sub 'Run

	'********'*********'*********'*********'*********'*********'*********'*********'
	'ScriptEnd
	'Adds a generic ending status message to the SecureCRT status bar text.
	'********'*********'*********'*********'*********'*********'*********'*********'
	Public Sub ScriptEnd()
		StatusText = "Done " & crt.ScriptFullName
		crt.Sleep(500)
		c_objTab.Session.SetStatusText "Ready"
	End Sub 'ScriptEnd
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'ScriptStart
	'Adds a generic starting status message to the SecureCRT status bar text.
	'********'*********'*********'*********'*********'*********'*********'*********'
	Public Sub ScriptStart()
		StatusText = "Starting"
		crt.Sleep(500)
		StatusText = "Running"
	End Sub 'ScriptStart

	'********'*********'*********'*********'*********'*********'*********'*********'
	'Send
	'********'*********'*********'*********'*********'*********'*********'*********'
	Public Sub Send(strSend)
		c_objTab.Screen.Send strSend & DEFAULT_LINE_TERMINATOR
	End Sub 'Send
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'SendAndReadStrings
	'This will send strCommand to the screen and return an array of screen lines 
	'until a string in varWaitFor is found.  This attempts to not collect the 
	'original command in the returned array.  varWaitFor can be a single string or
	'an Array of strings.
	'********'*********'*********'*********'*********'*********'*********'*********'
	Public Function SendAndReadStrings(strCommand, varWaitFor, intTimeOut)
		Dim arrWaitFor, strScreenData, arrScreenData
		Dim blnSynchronous
		
		'Capture the existing Synchronous setting so that it can be reset correctly after.
		blnSynchronous = c_objTab.screen.Synchronous
		c_objTab.screen.Synchronous = True 
		
		'Make sure that varStrings is an array.
		If (IsArray(varWaitFor)) Then
			arrWaitFor = varWaitFor
		Else
			arrWaitFor = Array(varWaitFor)	
		End If
		
		c_objTab.Screen.IgnoreEscape = True
		'Try to ignore the command string itself.
'TODO: The SendAndWaitFor ignores the DLT already so can't use it.
'TODO:		SendAndWaitFor strCommand, DEFAULT_LINE_TERMINATOR, intTimeOut
		Send strCommand
		WaitFor DEFAULT_LINE_TERMINATOR, intTimeOut
		strScreenData = c_objTab.Screen.ReadString(arrWaitFor, intTimeOut, IgnoreCase)
'FIXME: ReadString2 is used with clsSecureCRTSimulator since vbscript isn't polymorphic.
'FIXME:		strScreenData = c_objTab.Screen.ReadString2(arrWaitFor, intTimeOut, IgnoreCase)
		arrScreenData = Split(strScreenData, DEFAULT_LINE_TERMINATOR)
		SendAndReadStrings = arrScreenData
		
		c_objTab.screen.Synchronous = blnSynchronous 
	End Function 'SendAndReadStrings
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'SendAndWaitFor
	'This function automates the c_objTab.Screen.Send command and either the 
	'c_objTab.Screen.WaitForString or c_objTab.Screen.WaitForStrings command.  Either
	'a String or an Array may be passed into varWaitFor.
	'For String, a True or False is returned.
	'For Array, an interger position is returned starting with Array index 0 being 
	'position 1.  A 0 is returned if no strings were found.  So you can test for 
	'True (non-zero) or False (zero).
	'This attemts to not collect the strSend command string that was sent itself 
	'by waiting for the DEFAULT_LINE_TERMINATOR and then waiting for the varWaitFor.
	'
	'NOTE: It's assumed that most all commands sent to the screen require a DLT to 
	'	commit the command to the server, and that the command itself is not 
	'	explected to be searched within.  If you are sending a command where you 
	'	do not expect a DLT to wait for or you whish to search within the command 
	'	string itself, you will need to instead perform a Send and then a WaitFor 
	'	wrapped in a "Synchronous = True".
	'********'*********'*********'*********'*********'*********'*********'*********'
	Public Function SendAndWaitFor(strSend, varWaitFor, intTimeout)
		Dim blnSynchronous
		Dim intRet
		Dim blnRet
		intRet = 0
		blnRet = False 
		
		'Capture the existing Synchronous setting so that it can be rest correctly after.
		blnSynchronous = c_objTab.screen.Synchronous
		c_objTab.screen.Synchronous = True 
		
		Send strSend
		'Ignore the strSend itself
		WaitFor DEFAULT_LINE_TERMINATOR, intTimeOut
		If (IsArray(varWaitFor)) Then
			SendAndWaitFor = WaitForStrings(varWaitFor, intTimeout)
		Else
			SendAndWaitFor = WaitForString(varWaitFor, intTimeout)
		End If
		c_objTab.screen.Synchronous = blnSynchronous 
	End Function 'SendAndWaitFor	
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'SetStatusText
	'DEPRECIATED: Legacy method.  Use StatusText going forward.
	'********'*********'*********'*********'*********'*********'*********'*********'
	Public Sub SetStatusText(strText)
'		c_objTab.Session.SetStatusText strText & ": " & crt.ScriptFullName
		StatusText = strText
	End Sub 'SetStatusText
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'StatusText
	'Added to allow Get'ing the status text.
	'********'*********'*********'*********'*********'*********'*********'*********'
	Private c_strStatusText
	Public Property Let StatusText(strStatusText)
		c_strStatusText = strStatusText
		c_objTab.Session.SetStatusText strStatusText & ": " & crt.ScriptFullName
	End Property
	Public Property Get StatusText
		StatusText = c_strStatusText
	End Property 
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'Sleep
	'********'*********'*********'*********'*********'*********'*********'*********'
	Public Sub Sleep(intMilliSeconds)
		Dim strStatusText
		strStatusText = StatusText
		StatusText = "Sleeping for " & intMilliSeconds/1000 & " seconds" 
		crt.Sleep intMilliSeconds
		StatusText = strStatusText
	End Sub 'Sleep
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'SleepSeconds
	'********'*********'*********'*********'*********'*********'*********'*********'
	Public Sub SleepSeconds(intSeconds)
		Sleep intSeconds * 1000
	End Sub 'SleepSeconds
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'StripTN
	'********'*********'*********'*********'*********'*********'*********'*********'
	Public Function StripTN(strTN)
		stripTN = Replace(Replace(Replace(strTN, "-", ""), "_", ""), " ", "")
	End Function 'StripTN
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'Tab
	'The SecureCRT Tab window to interact with.  This will default to the Tab that 
	'the currently running script was called from.  But you can Set it here to use
	'any other Tab object.
	'********'*********'*********'*********'*********'*********'*********'*********'
	Private c_objTab
	Public Property Set Tab(objTab)
		Set c_objTab = objTab
	End Property
	Public Property Get Tab
		Set Tab = c_objTab
	End Property 
	
	'******************************************************************************
	'TestLength 
	'This will test the length of strString and warn the user if it is incorrect.
	'******************************************************************************
	Public Function TestLength(strString, strLabel, intLength)
		If (Len(strString) <> intLength) Then
			TestLength = False
			crt.Dialog.MessageBox "Incorrect length for '" & strLabel & "'!" _
					& vbCrLf & vbCrLf & "Lenght=" & intLength & " was expected." _
					& vbCrLf & "Lenght=" & Len(strString) & " was returned for '" _
					& strString & "'.", "INCORRECT LENGTH", vbExclamation + vbOKOnly
		Else
			TestLength = True 
		End If
	End Function 'TestLength
	
	'******************************************************************************
	'TestLengthMinMax 
	'This will test the min and max length of strString and warn the user if it is incorrect.
	'******************************************************************************
	Public Function TestLengthMinMax(strString, strLabel, intMinLength, intMaxLength)
		If (Len(strString) >= intMinLength And Len(strString) <= intMaxLength) Then
			TestLengthMinMax = True
		Else
			TestLengthMinMax = False
			
			crt.Dialog.MessageBox "Incorrect length for '" & strLabel & "'!"_ 
				& vbCrLf & vbCrLf & "Minimun Length=" & intMinLength & " and " _
				& "Maximum Length=" & intMaxLength & " was expected." _
				& vbCrLf & "Actual Length=" & Len(strString) & " was returned for '" _
				& strString & "'.", "INCORRECT LENGTH", vbExclamation + vbOKOnly
	
		End If
	End Function 'TestLengthMinMax
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'WaitFor
	'********'*********'*********'*********'*********'*********'*********'*********'
	Public Function WaitFor(varText, intTimeout)
		If (IsArray(varText)) Then
			WaitFor = WaitForStrings(varText, intTimeout)
		Else
			WaitFor = WaitForString(varText, intTimeout)
		End If
	End Function 'WaitFor
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'WaitForString
	'Simulates the functionality of the Crosstalk Wait command.  
	'It allows you to pass a blank string and a timeout value to make 
	'SecureCRT sleep.  It can also display a message box if the Timeout 
	'is encountered depending on Debug Levels..
	'********'*********'*********'*********'*********'*********'*********'*********'
	Public Function WaitForString(strText, intTimeout)
		Dim blnRet	
		Dim strMsg	
		Dim strStatusText

		strStatusText = StatusText
		StatusText = "WaitForString with timeout=" & intTimeout

		If (Len(strText) > 0) Then 
			If (intTimeout = 0) Then 
				blnRet = c_objTab.Screen.WaitForString(strText, 0, IgnoreCase)
			Else 
				blnRet = c_objTab.Screen.WaitForString(strText, intTimeout, IgnoreCase)
			End If 	
					
			If (DebugLevel >= 1) Then 
				If (Not blnRet) Then 
					strMsg = "WaitForString timed out waiting for """ & _
							strText & """ after " & intTimeout & " seconds."
					If (Len(g_strDebugLocator) > 0) Then
						strMsg = strMsg & vbCrLf & "After Location: """ & g_strDebugLocator & """."
					End If 
					crt.Dialog.Messagebox strMsg, "TIMEOUT", 48
				End If 	
			End If 	
		Else
			SleepSeconds intTimeout
			blnRet = False
		End If 
		WaitForString = blnRet
		StatusText = strStatusText
	End Function 'WaitForString
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'WaitForStrings
	'Simulates some of the functionality of the Crosstalk Wait command.  
	'It requires an array of string and returns the index of the matched item.
	'Return: And interger position is returned starting with arrstrText index 0 being 
	'position 1.  A 0 is returned if no strings were found.  So you can test for 
	'True (non-zero) or False (zero).
	'********'*********'*********'*********'*********'*********'*********'*********'
	Public Function WaitForStrings(arrstrText, intTimeout)
		Dim intRet	
		Dim strMsg
		Dim strStatusText
		
		strStatusText = StatusText
		StatusText = "WaitForStrings with timeout=" & intTimeout
		
		If (IsArray(arrstrText)) Then  
			If (UBound(arrstrText) > 0) Then 'TODO: test this to make sure a non array or empty array work.
				If (intTimeout = 0) Then 
					intRet = c_objTab.Screen.WaitForStrings(arrstrText, 0, IgnoreCase)
				Else 
					intRet = c_objTab.Screen.WaitForStrings(arrstrText, intTimeout, IgnoreCase)
				End If 
				If (DebugLevel >= 1) Then 
					If (intRet = 0) Then 
						strMsg = "WaitForStrings timed out waiting for """ & _
								Join(arrstrText, ", ") & """ after " & intTimeout & " seconds."
						If (Len(g_strDebugLocator) > 0) Then
							strMsg = strMsg & vbCrLf & "After Location: """ & g_strDebugLocator & """."
						End If 
						crt.Dialog.Messagebox strMsg, "TIMEOUT", 48
					End If 	
				End If 	
			Else
				SleepSeconds intTimeout
				intRet = 0
			End if
		Else
			SleepSeconds intTimeout
			intRet = 0
		End If 

		WaitForStrings = intRet
		StatusText = strStatusText
	End Function 'WaitForStrings
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'WriteMwqDn
	'********'*********'*********'*********'*********'*********'*********'*********'
	Public Sub WriteMwqDn(strLine)
		Dim objFSO
		Dim objTextFile
		
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		
		If (Not objFSO.FolderExists(DIR_NGVN_DATA)) Then
		    objFSO.CreateFolder(DIR_NGVN_DATA)
		End If
		Set objTextFile = objFSO.OpenTextFile(DIR_NGVN_DATA & DEFAULT_MWQ_DN_FILE, ForWriting, True)
		objTextFile.WriteLine(strLine)
		objTextFile.Close
	End Sub 'WriteMwqDn
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'Prompt returns true if data is returned or false if the returned value is empty.
	'strPrompt = Text to be presented in the prompt box.
	'strTitle = Text for the title of the prompt box.
	'strValue = Both the default value for the prompt box and holds the returned value.
	'********'*********'*********'*********'*********'*********'*********'*********'
	Public Function Prompt(strPrompt, strTitle, ByRef strValue) 'returns boolean
		Dim strReturn

		strReturn = Trim(crt.Dialog.Prompt(strPrompt, strTitle, strValue))
		If (Len(strReturn) > 0) Then
			strValue = strReturn
			Prompt = True 
		Else
			strValue = ""
			Prompt = False 
		End If
	End Function
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'IsPosInteger tests if a string is a positive integer by testing each character.
	'TODO: Used in Buildmems.xws, c7addmems.xws, FindDPC.xws, HCPYTRK.XWS, omshow.xws, posttrkgrp.xws
	'********'*********'*********'*********'*********'*********'*********'*********'
	Public Function IsPosInteger(strTest) 'returns boolean
		Dim i
		
		strTest = Trim(strTest)
		for i = 1 to Len(strTest)
			If (Not isNumeric(Mid(strTest, i, 1))) Then
				IsPosInteger = False
				Exit Function 
			End If 
		Next
		IsPosInteger = True
	End Function 'IsPosInteger
	
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'getHostName
	'********'*********'*********'*********'*********'*********'*********'*********'
	Public  Function GetHostName(objScriptTab)
		Dim objConfig
		
		Set objConfig = objScriptTab.Session.Config
		GetHostName = objConfig.GetOption("Hostname")
	End Function 'GetHostName
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'fileExists
	'********'*********'*********'*********'*********'*********'*********'*********'
	Public Function FileExists(strFileName)
		Dim objFSO
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		
		If objFSO.FileExists(strFileName) Then
		    FileExists = True 
		Else
		    FileExists = False
		End If
	End Function 'FileExists
			
End class
