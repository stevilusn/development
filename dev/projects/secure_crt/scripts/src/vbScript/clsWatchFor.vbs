'-------------------------------------------------------------------------------
'FileName:	clsWatchFor.vbs
'Author:	ke066443
'Purpose:	This class reproduces some of the functionality of the CrossTalk 
'			"watch" command.  The current version of SecureCRT can not perform 
'			pattern matching with it's WaitForString(s) functions.
'Dependency: This class can not be used on it's own.  It is intended to be run
'	from inside another script that has set DEFAULT constants and etc.
'Properties:
'	...
'Methods:	
'	...
'
'-------------------------------------------------------------------------------
Option Explicit

Class clsWatchFor	
	Private Sub Class_Initialize   ' Setup Initialize event.
		Reset	
	End Sub

	'********'*********'*********'*********'*********'*********'*********'*********'
	'Reset
	'********'*********'*********'*********'*********'*********'*********'*********'
	Public Sub Reset()
		Set c_objTab = crt.GetScriptTab
		c_strStringToSend = ""
		c_blnIgnoreStringToSend = True
		c_arrstrLines = Array()
		c_arrstrPatterns = Array()
		c_lngTimeoutSeconds = DEFAULT_TIMEOUT
		c_blnUseRegExp = True
		c_blnIgnoreCase = True
		c_blnGlobal = True
		c_blnSearchPatternByLine = True
		c_strLineTerminator = DEFAULT_LINE_TERMINATOR	
	End Sub
	
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
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'SearchPatternByLine
	'Defalts to True.  If True, then a complete line will be read in and then the 
	'RegExp pattern will be searched against the complete line.  If Fase, each 
	'character being sent to the screen will be added to the string and the RegExp 
	'patter is searched each time until a complete line has been added.  True is 
	'the faster search method.  
	'Note: A LineTerminator must end the line or it will not be searched.   Can't 
	'search a whole line if you can't recognize where the end is.  So searching for
	'the command prompt usually will not work with SearchPatternByLine=True since 
	'it's usually the last data in the screen stream and will not include a 
	'LineTerminator.
	'
	'FIXME: This didn't work for the FSO HCPYTRK.vbs call to SearchClliCdr when I 
	'removed the StringToSend for the "count" command.  Speen was fine with the 
	'character by character search, so we'll leave it that way for now.
	'********'*********'*********'*********'*********'*********'*********'*********'
	Private c_blnSearchPatternByLine
	Public Property Let SearchPatternByLine(blnSearchPatternByLine)
		c_blnSearchPatternByLine = blnSearchPatternByLine
	End Property
	Public Property Get SearchPatternByLine
		SearchPatternByLine = c_blnSearchPatternByLine
	End Property 
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'StringToSend
	'Issue a Send command with this value before Executing.  Set StringToSendInore
	'to false if you want to include this value in the search.  Leve this blank if 
	'you have done your Send prior to issuing this command. 
	'********'*********'*********'*********'*********'*********'*********'*********'
	Private c_strStringToSend
	Public Property Let StringToSend(strStringToSend)
		c_strStringToSend = strStringToSend
	End Property
	Public Property Get StringToSend
		StringToSend = c_strStringToSend
	End Property 
		
	'********'*********'*********'*********'*********'*********'*********'*********'
	'IgnoreStringToSend
	'Default is True.  Normally, when you issue a command to the screen, you do not 
	'want to search the command itself, you just want to search the results.  This 
	'will ignore the value in StringToSend by having Execute perform a WaitForString 
	'after sending the value in StringToSend.  Setting this to False will allow you 
	'to include the value of StringToSend in the search.
	'********'*********'*********'*********'*********'*********'*********'*********'
	Private c_blnIgnoreStringToSend
	Public Property Let IgnoreStringToSend(blnIgnoreStringToSend)
		c_blnIgnoreStringToSend = blnIgnoreStringToSend
	End Property
	Public Property Get IgnoreStringToSend
		IgnoreStringToSend = c_blnIgnoreStringToSend
	End Property 
	

	'********'*********'*********'*********'*********'*********'*********'*********'
	'Lines
	'Get the array of lines captured from the SecureCRT screen.
	'********'*********'*********'*********'*********'*********'*********'*********'
	Private c_arrstrLines
	Public Property Let Lines(arrstrLines)
		c_arrstrLines = arrstrLines
	End Property 
	Public Property Let LinesAdd(strLine)
		ReDim Preserve c_arrstrLines(UBound(c_arrstrLines) + 1)
		c_arrstrLines(UBound(c_arrstrLines)) = strLine
	End Property 
	Public Property Get Lines
		Lines = c_arrstrLines
	End Property 
	Public Property Get LinesToString
		LinesToString = Join(c_arrstrLines, LineTerminator)
	End Property 

	'********'*********'*********'*********'*********'*********'*********'*********'
	'Patterns
	'Let or Get the the patterns to search for.  If the input is not an array, add 
	'the object to the array.  Each subsequent Let will continue to add to c_arrstrPatterns.
	'********'*********'*********'*********'*********'*********'*********'*********'
	Private c_arrstrPatterns
	Property Let Patterns(varPatterns)
		Dim varItem

		If (IsArray(varPatterns)) Then
			For Each varItem In varPatterns
				ReDim Preserve c_arrstrPatterns(UBound(c_arrstrPatterns) + 1)
				c_arrstrPatterns(UBound(c_arrstrPatterns)) = varItem
			Next
		Else
			ReDim Preserve c_arrstrPatterns(UBound(c_arrstrPatterns) + 1)
			c_arrstrPatterns(UBound(c_arrstrPatterns)) = varPatterns
		End If 
	End Property 
	Property Get Patterns
		Patterns = c_arrstrPatterns
	End Property
	Public Sub ClearPatterns()
		c_arrstrPatterns = Array()
		c_arrstrLines = Array()
	End Sub
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'TimeoutSeconds
	'Let or Get the Seconds to time out.  This is a by the clock time out. 
	'********'*********'*********'*********'*********'*********'*********'*********'
	Private c_lngTimeoutSeconds
	Property Let TimeoutSeconds(lngTimeoutSeconds)
		If (Not IsNumeric(lngTimeoutSeconds)) Then
			lngTimeoutSeconds = CLng(lngTimeoutSeconds)
		End If 
		c_lngTimeoutSeconds = lngTimeoutSeconds
	End Property 
	Property Get TimeoutSeconds
		TimeoutSeconds = c_lngTimeoutSeconds
	End Property
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'UseRegExp
	'Use RegExp matching, else use string literal matching.
	'********'*********'*********'*********'*********'*********'*********'*********'
	Private c_blnUseRegExp
	Public Property Let UseRegExp(blnUseRegExp)
		c_blnUseRegExp = blnUseRegExp
	End Property
	Public Property Get UseRegExp
		UseRegExp = c_blnUseRegExp
	End Property 
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'IgnoreCase
	'Perform a case insensitive pattern match.
	'********'*********'*********'*********'*********'*********'*********'*********'
	Private c_blnIgnoreCase
	Public Property Let IgnoreCase(blnIgnoreCase)
		c_blnIgnoreCase = blnIgnoreCase
	End Property
	Public Property Get IgnoreCase
		IgnoreCase = c_blnIgnoreCase
	End Property 	
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'Global
	'Set RegExp Global to true or false.
	'********'*********'*********'*********'*********'*********'*********'*********'
	Private c_blnGlobal
	Public Property Let Global(blnGlobal)
		c_blnGlobal = blnGlobal
	End Property
	Public Property Get Global
		Global = c_blnGlobal
	End Property 
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'LineTerminator
	'Set the line terminator so that screen lines can be added to c_arrstrLines.
	'********'*********'*********'*********'*********'*********'*********'*********'
	Private c_strLineTerminator
	Public Property Let LineTerminator(strLineTerminator)
		c_strLineTerminator = strLineTerminator
	End Property
	Public Property Get LineTerminator
		LineTerminator = c_strLineTerminator
	End Property 
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'MatchPosition
	'If a match is found, Get will retun the position within the current screen line
	'that the MatchValue was found.
	'********'*********'*********'*********'*********'*********'*********'*********'
'	Private c_intMatchPosition
'	Public Property Let MatchPosition(intMatchPosition)
'		c_intMatchPosition = intMatchPosition
'	End Property
	Public Property Get MatchPosition
'		MatchPosition = c_intMatchPosition
		Dim arrLines 'use this because VBScript will error if you use Lines.
			
		arrLines = Lines
		MatchPosition = InStr(arrLines(UBound(arrLines)), MatchValue)
	End Property 
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'MatchValue
	'If a match is found, Get will return which value in c_arrstrPatterns was matched.
	'********'*********'*********'*********'*********'*********'*********'*********'
	Private c_strMatchValue
	Public Property Let MatchValue(strMatchValue)
		c_strMatchValue = strMatchValue
	End Property
	Public Property Get MatchValue
		MatchValue = c_strMatchValue
	End Property 
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'Execute
	'Start a WaitFor search of the incoming screen data.  
	'Return: The position (index + 1) of the matched value in array Patterns, 
	'or 0 if no match was found (a timeout is the same as no match found).  Can 
	'return a -1 if code execution fails (primary usage in Unit Testing).
	'********'*********'*********'*********'*********'*********'*********'*********'
	Public Property Get Execute()
		Dim blnContinue
		Dim i 
		Dim strChr
		Dim strString
		Dim intRet
		Dim dtmThen
		Dim blnSynchronous
		Dim strStatusText
		
		blnSynchronous = c_objTab.Screen.Synchronous
		c_objTab.Screen.Synchronous = True 
		c_objTab.Session.SetStatusText "WatchFor with timeout=" & TimeoutSeconds & ": " & crt.ScriptFullName

		dtmThen = Now()
		If (Len(StringToSend) > 0) Then 
			c_objTab.Screen.Send StringToSend & LineTerminator
			If (IgnoreStringToSend) Then
				c_objTab.Screen.WaitForString LineTerminator, TimeOutSeconds 'Don't WatchFor anything in the string being sent itself.
			End If
		End If
		intRet = -2
		blnContinue = True
		Do 
			If (SearchPatternByLine) Then
'FIXME: If we reach the prompt, then there is no LineTerminator and ReadString will timeout. So even if we are searching for the prompt, it'll timeout.
				strString = c_objTab.Screen.ReadString(LineTerminator, TimeOutSeconds)
			Else 
				strChr = c_objTab.Screen.ReadString
				strString = strString & strChr
			End If 

			'strString can be empty.  This would indicate a timeout.
			If (Len(strString) > 0) Then
				If (UseRegExp) Then
					intRet = checkStringRegExp(strString)
				Else
					intRet = checkStringLiteral(strString)
				End If 
			Else
				intRet = -1
				blnContinue = False 
			End If
			
			If (intRet > -1) Then
				blnContinue = False 
			ElseIf (Now() > DateAdd("s", TimeOutSeconds, dtmThen)) Then
				intRet = -1
			 	blnContinue = False
			End If  
		Loop While blnContinue
		c_intMatchIndex = intRet + 1 'Shift the index by +1 to correspond to the position of the searched data.
		c_objTab.Screen.Synchronous = blnSynchronous 
		c_objTab.Session.SetStatusText "Running: " & crt.ScriptFullName
		Execute = c_intMatchIndex 
	End Property 
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'MatchIndex
	'Index number in Array c_arrstrPatterns containing the WaitFor pattern that 
	'was found. 
	'********'*********'*********'*********'*********'*********'*********'*********'
	Private c_intMatchIndex
	Public Property Get MatchIndex
		MatchIndex = c_intMatchIndex
	End Property 
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'checkStringRegExp
	'RegExp search for a WaitFor pattern found in c_arrstrPatterns within 
	'the strString .
	'Return: -1 if not found, or the index of the matched value in c_arrstrPatterns.
	'********'*********'*********'*********'*********'*********'*********'*********'
	Private Function checkStringRegExp(byref strString)
		Dim i
		Dim strPattern

		checkStringRegExp = -1
		If (SearchPatternByLine) Then
			For i = LBound(c_arrstrPatterns) To UBound(c_arrstrPatterns) Step 1
				strPattern = c_arrstrPatterns(i)
				If (RegExpTest(strString, strPattern)) Then
					checkStringRegExp = i
					Exit For
				End If 
			Next
			checkForLineTerminator(strString)
		Else
			If (Not checkForLineTerminator(strString)) Then 
				For i = LBound(c_arrstrPatterns) To UBound(c_arrstrPatterns) Step 1
					strPattern = c_arrstrPatterns(i)
					If (RegExpTest(strString, strPattern)) Then
						LinesAdd = strString
						strString = ""
						checkStringRegExp = i
						Exit For
					End If 
				Next
			End If
		End If 
	End Function
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'RegExpTest
	'Perform a RegExp test for a WaitFor pattern within strString and return True 
	'if found or False if not.
	'********'*********'*********'*********'*********'*********'*********'*********'
	Function RegExpTest(strString, strPattern)
		Dim regEx
		Dim Match
		Dim Matches
		Dim blnMatch
		
		Set regEx = New RegExp   ' Create a regular expression.
		regEx.Pattern = strPattern   ' Set pattern.
		regEx.IgnoreCase = c_blnIgnoreCase   ' Set case insensitivity.
		regEx.Global = c_blnGlobal   ' Set global applicability.
		Set Matches = regEx.Execute(strString)   ' Execute search.
		blnMatch = False
		For Each Match in Matches   ' Iterate Matches collection.
			blnMatch = True 
			'Since we stop checkStringRegExp on the first match, 
			'this will only ever return one match.
'			c_intMatchPosition = Match.FirstIndex + 1
			c_strMatchValue = Match.Value
		Next
		RegExpTest = blnMatch
	End Function

	'********'*********'*********'*********'*********'*********'*********'*********'
	'checkStringLiteral
	'String literal search for a WaitFor pattern found in c_arrstrPatterns within 
	'the strString.
	'Return: -1 if not found, or the index of the matched value in c_arrstrPatterns.
	'********'*********'*********'*********'*********'*********'*********'*********'
	Private Function checkStringLiteral(byref strString)
		Dim i

		checkStringLiteral = -1
		If (SearchPatternByLine) Then
			For i = LBound(c_arrstrPatterns) To UBound(c_arrstrPatterns) Step 1
				If (InStr(strString, c_arrstrPatterns(i))) Then
					checkStringLiteral = i
					Exit For
				End If 
			Next	 
			LinesAdd = Left(strString, Len(strString) - 1)
			strString = ""
		Else
			'Processing one character at a time to build strString, 
			'so if a LineTerminator is found process the line. 		
			If (Not checkForLineTerminator(strString)) Then 
				For i = LBound(c_arrstrPatterns) To UBound(c_arrstrPatterns) Step 1
					If (InStr(strString, c_arrstrPatterns(i))) Then
						LinesAdd = strString
						strString = ""
						checkStringLiteral = i
						Exit For
					End If 
				Next
			End if				 	
		End If
	End Function
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'checkForLineTerminator
	'Check for a Line Terminator and start a new line.  If found then store the 
	'current line in c_arrstrLines, clear strString and pass it back, and return True.
	'********'*********'*********'*********'*********'*********'*********'*********'
	Private Function checkForLineTerminator(ByRef strString)
		Dim intPos
		intPos = InStr(strString, LineTerminator)
		If (intPos > 0) Then
			LinesAdd = Left(strString, intPos - 1)
			strString = Mid(strString, intPos + 1)
			checkForLineTerminator = True 
		Else 
			checkForLineTerminator = False
		End If 
	End Function
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'toStringLines
	'Purpose:	What the routine does (not how).
	'Inputs:	Each non-obvious parameter on a separate line with
	'			in-line comments
	'Assumes:	List of each non-obvious external variable, control, open file,
	'			and so on.
	'Returns:	Explanation of value returned for functions.
	'Effects:	List of each effected external variable, control, file, and
	'			so on and the affect it has (only if this is not obvious)
	'********'*********'*********'*********'*********'*********'*********'*********'
	Public Function toStringLines
		toStringLines = Join(WatchFor.Lines, LineTerminator)
	End Function
	Public Function toStringPatterns
		toStringPatterns = Join(WatchFor.Patterns, ", ")
	End Function		
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'SearchClliCdr
	'Attempts to search for and return the CLLI (intTrkNum) CDR data.
	'
	'Used in FSO scripts Buildmems.xws, c7addmems.xws, HCPYTRK.XWS, omshow.xws, 
	'posttrkgrp.xws.  Although putting this function in the WatchFor class isn't 
	'exactly appropriate, it was added here for conviniance since it is only 
	'dependant on the WatchFor class
	'********'*********'*********'*********'*********'*********'*********'*********'
	Function SearchClliCdr(intTrkNum) 'returns string
		Dim strTupleCount
		Dim strTrkNum
		Dim strRetVal
		Dim arrstrPatterns
		Dim strMatch
		Dim blnSynchronous
		Dim intIndex
		
		blnSynchronous = c_objTab.Screen.Synchronous
		c_objTab.Screen.Synchronous = True 
		
		strTrkNum = CStr(intTrkNum)
		
		c_objTab.Screen.Send "send sink;table cllicdr;format pack;send previous" & LineTerminator
		c_objTab.Screen.WaitForString ">", TimeoutSeconds
		c_objTab.Screen.Send "count (2 eq " & strTrkNum & " )" & LineTerminator		
		c_objTab.Screen.WaitForString "BOTTOM", TimeoutSeconds

		Reset
		UseRegExp = True
		IgnoreCase = True
		Global = False
		TimeoutSeconds = 60
'TODO:		SearchPatternByLine = True 
		SearchPatternByLine = False 
		arrstrPatterns = Array( "SIZE = [0-9]{1,4}")
		Patterns = arrstrPatterns
'TODO: 		StringToSend = "count (2 eq " & strTrkNum & " )"
		intIndex = Execute
		strMatch = MatchValue
		
		strTupleCount = "0"
		Select Case intIndex
			Case 0 
				crt.Dialog.Messagebox "WatchFor: Timed out watching for patterns after " _
						& TimeoutSeconds _
						& " seconds.  Patterns:" & vbCrLf _
						& Join(arrstrPatterns, vbCrLf) & vbCrLf _
						& "Will continute processing now.", _
						"TIMEOUT", vbOKOnly & vbInformation 
				c_objTab.Screen.Send LineTerminator
				c_objTab.Screen.WaitForString ">", TimeoutSeconds
			Case 1
				strTupleCount = Mid(strMatch, 8) 
				c_objTab.Screen.WaitForString ">", TimeoutSeconds
			Case 2 '"BOTTOM"
				c_objTab.Screen.WaitForString ">", TimeoutSeconds
		End Select
		
		If (strTupleCount = "0") Then 
			strRetVal = ""
		ElseIf (strTupleCount <> "1") Then
			crt.Dialog.MessageBox "More than one Trunk Group matches criteria.", "MULTIPLE MATCHES", vbOKOnly + vbExclamation
			c_objTab.Screen.Send "lis all (2 eq " & strTrkNum & " )" & LineTerminator
			c_objTab.Screen.WaitForString ">", TimeoutSeconds
			strRetVal = ""
		Else		
			Reset
			UseRegExp = True
			IgnoreCase = True
			Global = False
			TimeoutSeconds = 60
			SearchPatternByLine = True 
			arrstrPatterns = Array( "[0-9a-zA-Z]{1,16}\s*" & strTrkNum, _
									"BOTTOM", _
									">")
			Patterns = arrstrPatterns
			
			c_objTab.Screen.Send "lis all (2 eq " & strTrkNum & " )" & LineTerminator
			' our reply string will trigger next watch for, clear buffer
			c_objTab.Screen.WaitForString LineTerminator, TimeoutSeconds					
			' Watch for string of alph/numeric character up to a length of 16, some whitespace, followed by the trunk number.
			' Example "           TZ96        96"
			' Return clli string.
			intIndex = Execute
			strMatch = MatchValue
			
			Select Case intIndex
				Case 0
					crt.Dialog.Messagebox "Timed out watching for patterns after " _
							& TimeoutSeconds _
							& " seconds.  Patterns:" & vbCrLf _
							& Join(arrstrPatterns, vbCrLf) & vbCrLf _
							& "Will continute processing now.", _
							"TIMEOUT", vbOKOnly & vbInformation 
					c_objTab.Screen.Send LineTerminator
					c_objTab.Screen.WaitForString ">", TimeoutSeconds
					strRetVal = ""
				Case 1 '"[0-9a-zA-Z]{1,16} "
					strRetVal =  Left(strMatch, InStr(strMatch, " ") -1)
					c_objTab.Screen.WaitForString ">", TimeoutSeconds
				Case 2 '"BOTTOM"
					strRetVal = ""
					c_objTab.Screen.WaitForString ">", TimeoutSeconds
				Case 3 '">"
					crt.Dialog.MessageBox "The prompt ('>') was unexpectedly found.  Will assume that there was no match and continue.", "UNEXPECTED RESULTS", vbOKOnly + vbExclamation
					strRetVal = ""
					c_objTab.Screen.WaitForString ">", TimeoutSeconds
			End Select
		End If
		 
		c_objTab.Screen.Synchronous = blnSynchronous 
		SearchClliCdr = strRetVal
	End Function 'SearchClliCdr

End Class 	
