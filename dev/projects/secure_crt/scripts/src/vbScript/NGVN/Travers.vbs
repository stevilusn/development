'#$language = "VBScript"
'#$interface = "1.0"
'Name: Travers.vbs

Option Explicit

Dim s_objScriptTab
Dim s_objUtils
Dim s_objDialogBox
'Dim s_objStrings
Dim s_objArrays
Dim s_intTimeout

Dim s_arrstrCatchit(500)	'  array to store DINA feedback
Dim s_blnNonTwc
Dim s_blnNonTwcLrn
Dim s_blnTwcLocTgp
Dim s_blnTwcLocTgpLrn
Dim s_blnTwcLd
Dim s_blnTwcIntl
Dim s_blnNonTwcTf
Dim s_blnTwcTf
Dim s_blnInboundNonTwc
Dim s_blnInboundTwc

Dim s_strNonTwcTerm
Dim s_strNonTwcOrig
Dim s_strNonTwcTf
Dim s_strLataNum
Dim s_strCarrierNum
Dim s_strLata
Dim s_strTwcOrig
Dim s_strSsptkTrkGrp
Dim s_strOrigTrkGrp
Dim s_strTwcTerm
Dim s_strCktDigit
Dim s_strTraverToOrigLrnTrkgrp
Dim s_strLataName
Dim s_strTwcTf
Dim s_strCarrierNumber
Dim s_strOrigTwctfTrkGrp

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
'		Call loadSiteID	'TODO: Non of the variables set in the call are ever used.  Can it be removed?
		get_inbound_vs_outbound
	Else
		s_objUtils.CancelMsgBox
	End If 
	
	s_objUtils.ScriptEnd
	Set s_objDialogBox = Nothing 
'	Set s_objStrings = Nothing
	Set s_objArrays = Nothing
	Set s_objUtils = Nothing
End Sub 'Main

'********'*********'*********'*********'*********'*********'*********'*********'
'loadSiteID
'TODO: None of the variables set here are used anywhere else.  Can this sub be removed?  Is it's only purpose currently is to display the Switch? 
'********'*********'*********'*********'*********'*********'*********'*********'
Sub loadSiteID()
	Dim arrstrData
	Dim strSiteid
	Dim s_strSiteidCleaned

	crt.Screen.Synchronous = True
	
	arrstrData = s_objArrays.ReadFile(DIR_NGVN_DATA & s_objUtils.getSessionFileName())
	strSiteid = arrstrData(UBound(arrstrData))
	strSiteidCleaned = Strip(strSiteid, " ", 4)
	Select Case strSiteidCleaned
	    Case "905"
	    	s_strOrig = ""
	    Case "907"
	    	s_strOrig = ""
	    Case "909"
	    	s_strOrig = ""
	    Case "910"
	    	s_strOrig = ""
	    Case "910"
	    	s_strOrig = ""
	    Case "912"
	    	s_strOrig = ""
	End Select
	
	s_objUtils.Message "Switch " + s_strSiteid
End Sub 'loadSiteID

'********'*********'*********'*********'*********'*********'*********'*********'
'get_inbound_vs_outbound
'********'*********'*********'*********'*********'*********'*********'*********'
Function get_inbound_vs_outbound()
	Dim blnTryAgain
	Dim strChoice

	Do 
		blnTryAgain = False   
		With s_objDialogBox
			.Clear
			.Height = 300
			.Width = 420
			.LabelWidths = "10px"
			.Title = "Inbound versus Outbound Traver"
			.AddItem Array("type=html", "<u>S</u>elect Type of TRAVER to be performed:<br>")
			.AddItem Array("type=radio", "name=strChoice", "label=", "accesskey=S", "value=1", _
					"choices=Outbound Traver from cable partner customer|Inbound Traver towards cable partner customer", _
					"terminate=newline", "termination=newline")
			.AddItem Array("type=paragraph", "name=label1", "value=Note: To perform an Inbound Traver, you will need to run the Outbound traver script first to determine the trunk group the call may originate over.", "terminate=newline")
			.AddItem Array("type=paragraph", "name=label1", "value=Make note of the trunk group and re-run the traver script and choose inbound.", "terminate=newline")
			.Show
			If .Status <> vbOK Then 
				get_inbound_vs_outbound = False
				Exit Function  
			End If 
			strChoice = .Responses.Item("strChoice")           
		End With
		
		Select Case strChoice
		    Case "Outbound Traver from cable partner customer"
				get_inbound_vs_outbound = get_outbound_traver_type()
		    Case "Inbound Traver towards cable partner customer"
				get_inbound_vs_outbound = get_inbound_traver_type()
		    Case Else
				blnTryAgain = True 
				crt.Dialog.MessageBox "Please choose a radio button, or click the Cancel button.", "MISSING DATA", vbExclamation + vbOKOnly
		End Select
	Loop While blnTryAgain
End Function 'get_inbound_vs_outbound

'********'*********'*********'*********'*********'*********'*********'*********'
'get_inbound_traver_type
'********'*********'*********'*********'*********'*********'*********'*********'
Function get_inbound_traver_type()
	Dim blnTryAgain
	Dim strChoice

	Do 
		blnTryAgain = False    
		With s_objDialogBox
			.Clear
			.Height = 150
			.Width = 350
			.LabelWidths = "10px"
			.Title = "Inbound CS2K TRAVER"
			.AddItem Array("type=html", "<u>S</u>elect Type of TRAVER to be performed:<br>")
			.AddItem Array("type=radio", "name=strChoice", "label=", "accesskey=S", "value=1", _
					"choices=Inbound Non-TWC|Inbound TWC", _
					"terminate=newline", "termination=newline")
			.AddItem Array("type=html", "<div style='clear: both;'></div>")
			.Show
			If .Status <> vbOK Then
				get_inbound_traver_type = False  
				Exit Function  
			End If 
			strChoice = .Responses.Item("strChoice")           
		End With
		
		Select Case strChoice
		    Case "Inbound Non-TWC"
				get_inbound_traver_type = INBOUND_NON_TWC_TRAVER()
		    Case "Inbound TWC"
				get_inbound_traver_type = INBOUND_TWC_TRAVER() 
		    Case Else
		    	crt.Dialog.MessageBox "Please choose a radio button, or click the Cancel button.", "MISSING DATA", vbExclamation + vbOKOnly
				blnTryAgain = True   	
		End Select
	Loop While blnTryAgain
End Function 'get_inbound_traver_type

'********'*********'*********'*********'*********'*********'*********'*********'
'get_outbound_traver_type
'********'*********'*********'*********'*********'*********'*********'*********'
Function get_outbound_traver_type()
	Dim blnTryAgain
	Dim strChoice

'TODO:DIALOGBOX 24,23,190,195, "Outbound CS2K TRAVER"
'TODO:  groupbox    5,5,180,165,"Select Type of TRAVER to be performed:" 
'TODO:  radiobutton 18,20,150,9,"Non-TWC local line or LD traver ",s_blnNonTwc, TABSTOP GROUP
'TODO:  ltext    28,30,150,9,"(7, 10 and 1+10 digits)"
'TODO:  radiobutton 18,45,150,9,"Non-TWC local line traver to term number",s_blnNonTwcLrn
'TODO:  ltext    28,55,150,12,"with associated LRN"
'TODO:  radiobutton 18,70,150,9,"Non-TWC traver to Toll Free number",s_blnNonTwcTf
'TODO:  radiobutton 18,85,150,9,"TWC local trunk group traver",s_blnTwcLocTgp
'TODO:  radiobutton 18,100,160,9,"TWC local trunk group traver to term number",s_blnTwcLocTgpLrn
'TODO:  ltext    28,110,150,12,"with associated LRN (10 digit only)"
'TODO:  radiobutton 18,125,150,9,"TWC long distance traver", s_blnTwcLd
'TODO:  radiobutton 18,140,150,9,"TWC international traver", s_blnTwcIntl
'TODO:  radiobutton 18,155,150,9,"TWC Toll Free traver", s_blnTwcTf
'TODO:  DEFPUSHBUTTON 58,175,36,14,"OK", OK TABSTOP
'TODO:  PUSHBUTTON    108,175,36,14,"Cancel",CANCEL TABSTOP
'TODO:ENDDIALOG
'TODO:if choice = 2 then end


'TODO: Was this all one radio group?  Or could you make multiple selections?
	Do 
		blnTryAgain = False 
		With s_objDialogBox
			.Clear
			.Height = 370
			.Width = 600
			.LabelWidths = "10px"
			.Title = "Outbound CS2K TRAVER"
			.AddItem Array("type=html", "<u>S</u>elect Type of TRAVER to be performed:<br>")
			.AddItem Array("type=radio", "name=strChoice", "label=", "accesskey=S", "value=1", _
					"choices=" & _
					"Non-TWC local line or LD traver (7, 10 and 1+10 digits)|" & _
					"Non-TWC local line traver to term number with associated LRN|" & _
					"Non-TWC traver to Toll Free number|" & _
					"TWC local trunk group traver|" & _
					"TWC local trunk group traver to term number with associated LRN (10 digit only)|" & _
					"TWC long distance traver|" & _
					"TWC international traver|" & _
					"TWC Toll Free traver", _		
					"terminate=newline", "termination=newline")
			.AddItem Array("type=paragraph", "name=label21", "value=Note: To perform an Inbound Traver, you will need to run the outbound traver script first to determine the trunk group the call may originate over.", "terminate=newline")
			.AddItem Array("type=paragraph", "name=label22", "value=Make note of the trunk group and re-run the traver script and choose inbound.", "terminate=newline")
			.Show
			If .Status <> vbOK Then 
				get_outbound_traver_type = False 
				Exit Function  
			End If 
			strChoice = .Responses.Item("strChoice")        
		End With
		
		Select Case strChoice
		    Case "Non-TWC local line or LD traver (7, 10 and 1+10 digits)"
				get_outbound_traver_type = NON_TWC_LOCAL_OR_LD_TRAVER()
		    Case "Non-TWC local line traver to term number with associated LRN"
		    	get_outbound_traver_type = NON_TWC_LOCAL_LINE_TRAVER_WITH_LRN()
		    Case "Non-TWC traver to Toll Free number"
		    	get_outbound_traver_type = NON_TWC_TOLLFREE_TRAVER()
		    Case "TWC local trunk group traver"
		    	get_outbound_traver_type = TWC_LOCAL_TRK_GRP_TRAVER()
		    Case "TWC local trunk group traver to term number with associated LRN (10 digit only)"
		    	get_outbound_traver_type = TWC_LOCAL_TRK_GRP_TRAVER_WITH_LRN() 
		    Case "TWC long distance traver"
		    	get_outbound_traver_type = TWC_LONG_DISTANCE_TRAVER()
		    Case "TWC international traver"
		    	get_outbound_traver_type = TWC_INTERNATIONAL_TRAVER() 
		    Case "TWC Toll Free traver"
		    	get_outbound_traver_type = TWC_TOLLFREE_TRAVER()
		    Case Else
				crt.Dialog.MessageBox "Please choose a radio button, or click the Cancel button.", "MISSING DATA", vbExclamation + vbOKOnly
				blnTryAgain = True
		End Select
	Loop While blnTryAgain
	get_outbound_traver_type = True 
End Function 'get_outbound_traver_type

'********'*********'*********'*********'*********'*********'*********'*********'
'NON_TWC_LOCAL_OR_LD_TRAVER
'********'*********'*********'*********'*********'*********'*********'*********'
Function NON_TWC_LOCAL_OR_LD_TRAVER()
	Dim blnTryAgain
	Dim strChoice
	Dim strNonTwcTermOne
	Dim strNonTwcTermLocal

	Do 
		blnTryAgain = False 
		With s_objDialogBox
			.Clear
			.Height = 150
			.Width = 350
			.Title = "NON-TWC Local or LD Traver"
			.LabelWidths = "120px"
			.AddItem Array("type=text", "name=s_strNonTwcOrig", "label=Enter ORIG #:", "accesskey=R", "value=" & s_strNonTwcOrig, "size=13", "terminate=newline")
			.AddItem Array("type=text", "name=s_strNonTwcTerm", "label=Enter TERM #:", "accesskey=T", "value=" & s_strNonTwcTerm, "size=13", "terminate=newline")
			.Show
			If .Status <> vbOK Then 
				NON_TWC_LOCAL_OR_LD_TRAVER = False 
				Exit Function  
			End If 
			s_strNonTwcOrig = s_objUtils.stripTN(.Responses.Item("s_strNonTwcOrig"))           
			s_strNonTwcTerm = s_objUtils.stripTN(.Responses.Item("s_strNonTwcTerm"))          
		End With
		
		If (Len(s_strNonTwcOrig) = 0 Or Len(s_strNonTwcTerm) = 0) Then
			crt.Dialog.MessageBox "Please enter values for all fields, or click the Cancel button.", "MISSING DATA", vbExclamation + vbOKOnly
			blnTryAgain = True 
		End If
	Loop While blnTryAgain

	strNonTwcTermOne = "1" & s_strNonTwcTerm 
	strNonTwcTermLocal = Right(s_strNonTwcTerm, 7)

	s_objUtils.SendAndWaitFor "TRAVER L " & s_strNonTwcOrig & " " _
			& s_strNonTwcTerm & " NT", _
			">", s_intTimeout
	If (Len(s_strNonTwcTerm) = 10) then 
		s_objUtils.SendAndWaitFor "TRAVER L " & s_strNonTwcOrig & " " _
				& strNonTwcTermOne & " NT", _
				">", s_intTimeout
		s_objUtils.SendAndWaitFor "TRAVER L " & s_strNonTwcOrig & " " _
				& strNonTwcTermLocal & " NT", _
				">", s_intTimeout
	End If 
	NON_TWC_LOCAL_OR_LD_TRAVER = True 
End Function 'NON_TWC_LOCAL_OR_LD_TRAVER

'********'*********'*********'*********'*********'*********'*********'*********'
'NON_TWC_LOCAL_LINE_TRAVER_WITH_LRN
'********'*********'*********'*********'*********'*********'*********'*********'
Function NON_TWC_LOCAL_LINE_TRAVER_WITH_LRN()
	Dim blnTryAgain
	Dim strNextLine
	Dim strLRNCapture
	Dim strNonTwcLrnTerm
	Dim strNonTwcLrnOrig
	Dim strNonTwcLrn
	
'	strNonTwcLrnTerm = crt.Dialog.Prompt("Enter TERM #:", "NON-TWC Local or LD Traver")
'	strNonTwcLrnTerm = s_objUtils.stripTN(strNonTwcLrnTerm)
'	If (Len(strNonTwcLrnTerm) = 0) Then
'		NON_TWC_LOCAL_LINE_TRAVER_WITH_LRN = False
'		Exit Function 
'	End If
	
	Do 
		blnTryAgain = False 
		With s_objDialogBox
			.Clear
			.Height = 150
			.Width = 350
			.LabelWidths = "120px"
			.Title = "NON-TWC Local or LD Traver"
			.AddItem Array("type=text", "name=strNonTwcLrnTerm", "label=Enter TERM #:", "accesskey=T", "value=" & strNonTwcLrnTerm, "size=13", "terminate=newline")
			.Show
			If .Status <> vbOK Then 
				NON_TWC_LOCAL_LINE_TRAVER_WITH_LRN = False 
				Exit Function  
			End If 
			strNonTwcLrnTerm = s_objUtils.stripTN(.Responses.Item("strNonTwcLrnTerm"))           
		End With
		If (Trim(Len(strNonTwcLrnTerm)) = 0) Then
			crt.Dialog.MessageBox "Please enter values for all fields, or click the Cancel button.", "MISSING DATA", vbExclamation + vbOKOnly
			blnTryAgain = True 
		End If
	Loop While blnTryAgain

	strLRNCapture = getLRNCapture(strNonTwcLrnTerm)
	If (Len(strLRNCapture) = 0) Then
		NON_TWC_LOCAL_LINE_TRAVER_WITH_LRN = False 
		Exit Function  
	End If
	
	Do 
		blnTryAgain = False 
		With s_objDialogBox
			.Clear
			.Height = 150
			.Width = 350
			.LabelWidths = "120px"
			.Title = "NON-TWC Local or LD Traver"
			.AddItem Array("type=text", "name=strNonTwcLrnOrig", "label=Enter ORIG #:", "accesskey=R", "value=" & strNonTwcLrnOrig, "size=13", "terminate=newline")
			.AddItem Array("type=text", "name=strNonTwcLrnTerm", "label=Enter TERM #:", "accesskey=T", "value=" & strNonTwcLrnTerm, "size=13", "terminate=newline")
			.AddItem Array("type=text", "name=strNonTwcLrn", "label=Enter LRN #:", "accesskey=L", "value=" & strLRNCapture, "size=13", "terminate=newline")
			.Show
			If .Status <> vbOK Then 
				NON_TWC_LOCAL_LINE_TRAVER_WITH_LRN = False 
				Exit Function  
			End If 
			strNonTwcLrnOrig = s_objUtils.stripTN(.Responses.Item("strNonTwcLrnOrig"))           
			strNonTwcLrnTerm = s_objUtils.stripTN(.Responses.Item("strNonTwcLrnTerm"))           
			strNonTwcLrn = .Responses.Item("strNonTwcLrn")           
		End With

		If (Len(strNonTwcLrnOrig) = 0 Or Len(strNonTwcLrnTerm) = 0 Or Trim(Len(strNonTwcLrn) = 0)) Then
			crt.Dialog.MessageBox "Please enter values for all fields, or click the Cancel button.", "MISSING DATA", vbExclamation + vbOKOnly
			blnTryAgain = True 
		End If
	Loop While blnTryAgain

	s_objUtils.SendAndWaitFor "TRAVER L " & strNonTwcLrnOrig & " N CDN NA " _
			& strNonTwcLrn & " AINRES R01 LNPAR " & strNonTwcLrnTerm & " NT", _
			">", s_intTimeout
	
	NON_TWC_LOCAL_LINE_TRAVER_WITH_LRN = True 
End Function 'NON_TWC_LOCAL_LINE_TRAVER_WITH_LRN

'********'*********'*********'*********'*********'*********'*********'*********'
'TWC_LOCAL_TRK_GRP_TRAVER
'********'*********'*********'*********'*********'*********'*********'*********'
Function TWC_LOCAL_TRK_GRP_TRAVER()
	Dim blnTryAgain
	Dim strMsg
	
	If (Not DETERMINE_TWC_ORIG_TRUNK_GROUP()) Then
		TWC_LOCAL_TRK_GRP_TRAVER = False
		Exit Function
	End If 
	
	strMsg = "Enter Trunk Group from the results of the last traver that was "
	strMsg = strMsg & "was performed on your screen.  The script will then "
	strMsg = strMsg & "perform the Local Trunk Group Traver with LRN."
	Do 
		blnTryAgain = False 
		With s_objDialogBox
			.Clear
			.Height = 200
			.Width = 350
			.LabelWidths = "120px"
			.Title = "NON-TWC Local or LD Traver"
			.AddItem Array("type=paragraph", "name=label1", "value=" & strMsg, "terminate=newline")
			.AddItem Array("type=text", "name=s_strOrigTrkGrp", "label=ORIG TRKGRP:", "accesskey=R", "value=" & s_strOrigTrkGrp, "size=13", "terminate=newline")
			.Show
			If .Status <> vbOK Then 
				TWC_LOCAL_TRK_GRP_TRAVER = False 
				Exit Function  
			End If 
			s_strOrigTrkGrp = Replace(.Responses.Item("s_strOrigTrkGrp"), " ", "")           
		End With
		
		If (Len(s_strOrigTrkGrp) = 0) Then
			crt.Dialog.MessageBox "Please enter values for all fields, or click the Cancel button.", "MISSING DATA", vbExclamation + vbOKOnly
			blnTryAgain = True 
		End If
	Loop While blnTryAgain

	s_objUtils.SendAndWaitFor "TRAVER TR " & s_strOrigTrkGrp _
			& " " & s_strTwcTerm & " NT", _
			">", s_intTimeout
	
	TWC_LOCAL_TRK_GRP_TRAVER = True 
End Function 'TWC_LOCAL_TRK_GRP_TRAVER

'********'*********'*********'*********'*********'*********'*********'*********'
'TWC_LOCAL_TRK_GRP_TRAVER_WITH_LRN
'********'*********'*********'*********'*********'*********'*********'*********'
Function TWC_LOCAL_TRK_GRP_TRAVER_WITH_LRN()
	Dim blnTryAgain
	Dim strNextLine
	Dim strLRNCapture
	Dim strMsg
	
	If (Not DETERMINE_TWC_ORIG_TRUNK_GROUP()) Then
		TWC_LOCAL_TRK_GRP_TRAVER_WITH_LRN = False
		Exit Function
	End If 
	
	strMsg = "Enter Trunk Group from the results of the last traver that was "
	strMsg = strMsg & "performed on your screen.  The script will then "
	strMsg = strMsg & "perform the Local Trunk Group Traver with LRN."
	Do 
		blnTryAgain = False
		With s_objDialogBox
			.Clear
			.Height = 200
			.Width = 350
			.LabelWidths = "120px"
			.Title = "Local TWC Trunk Group Traver with LRN"
			.AddItem Array("type=paragraph", "name=label1", "value=" & strMsg, "terminate=newline")
			.AddItem Array("type=text", "name=s_strOrigTrkGrp", "label=ORIG TRKGRP:", "accesskey=R", "value=" & s_strOrigTrkGrp, "size=13", "terminate=newline")
			.Show
			If .Status <> vbOK Then 
				TWC_LOCAL_TRK_GRP_TRAVER_WITH_LRN = False 
				Exit Function  
			End If 
			s_strOrigTrkGrp = Replace(.Responses.Item("s_strOrigTrkGrp")," ", "")           
		End With
		
		If (Len(s_strOrigTrkGrp) = 0) Then
			crt.Dialog.MessageBox "Please enter values for all fields, or click the Cancel button.", "MISSING DATA", vbExclamation + vbOKOnly
			blnTryAgain = True 
		End If
	Loop While blnTryAgain
	
	strLRNCapture = getLRNCapture(s_strTwcTerm)
	If (Len(strLRNCapture) = 0) Then
		TWC_LOCAL_TRK_GRP_TRAVER_WITH_LRN = False 
		Exit Function  
	End If

	s_objUtils.SendAndWaitFor "TRAVER TR " & s_strOrigTrkGrp & " N CDN NA " _
			& strLRNCapture & " AINRES R01 LNPAR " & s_strTwcTerm & " NT", _
			">", s_intTimeout
	
	TWC_LOCAL_TRK_GRP_TRAVER_WITH_LRN = True 
End Function 'TWC_LOCAL_TRK_GRP_TRAVER_WITH_LRN

'********'*********'*********'*********'*********'*********'*********'*********'
'TWC_LONG_DISTANCE_TRAVER
'********'*********'*********'*********'*********'*********'*********'*********'
Function TWC_LONG_DISTANCE_TRAVER()
	Dim blnTryAgain
	Dim strCarCode
	Dim strMsg
	
	If (Not DETERMINE_TWC_ORIG_TRUNK_GROUP()) Then
		TWC_LONG_DISTANCE_TRAVER = False
		Exit Function
	End If 

	strMsg = "Enter Trunk Group from the results of the last traver that was "	
	strMsg = strMsg & "performed on your screen.  The script will then "	
	strMsg = strMsg & "perform the TWC Long Distance Traver."	
	Do 
		blnTryAgain = False 
		With s_objDialogBox
			.Clear
			.Height = 200
			.Width = 350
			.LabelWidths = "120px"
			.Title = "TWC Long Distance Traver"
			.AddItem Array("type=paragraph", "name=label1", "value=" & strMsg, "terminate=newline")
			.AddItem Array("type=text", "name=s_strOrigTrkGrp", "label=ORIG TRKGRP:", "accesskey=R", "value=" & s_strOrigTrkGrp, "size=13", "terminate=newline")
			.AddItem Array("type=text", "name=strCarCode", "label=Carrier Code:", "accesskey=a", "value=0333", "size=13", "terminate=newline")
			.Show
			If .Status <> vbOK Then 
				TWC_LONG_DISTANCE_TRAVER = False 
				Exit Function  
			End If 
			s_strOrigTrkGrp = Replace(.Responses.Item("s_strOrigTrkGrp")," ", "")           
			strCarCode = .Responses.Item("strCarCode")           
		End With
		
		If (Len(s_strOrigTrkGrp) = 0 Or Len(strCarCode) = 0) Then
			crt.Dialog.MessageBox "Please enter values for all fields, or click the Cancel button.", "MISSING DATA", vbExclamation + vbOKOnly
			blnTryAgain = True 
		End If
	Loop While blnTryAgain

	s_objUtils.SendAndWaitFor "table cktdigit;pos "  & strCarCode _
			& " 1;lis 10;leave all", _
			">", s_intTimeout

	strMsg = strCarCode & " 8 is for 1+ Domestic Direct Dialed / Toll Free<br>" & vbcrlf	
	strMsg = strMsg & strCarCode & " 9 is for Operator Assisted Domestic LD Call"	
	Do 
		blnTryAgain = False 
		With s_objDialogBox
			.Clear
			.Height = 180
			.Width = 350
			.LabelWidths = "120px"
			.Title = "Choose OZZ Code"
			.AddItem Array("type=paragraph", "name=label1", "value=" & strMsg, "terminate=newline")
			.AddItem Array("type=text", "name=s_strCktDigit", "label=Enter OZZ Code:", "accesskey=Z", "value=" & s_strCktDigit, "size=13", "terminate=newline")
			.Show
			If .Status <> vbOK Then 
				TWC_LONG_DISTANCE_TRAVER = False 
				Exit Function  
			End If 
			s_strCktDigit = Replace(.Responses.Item("s_strCktDigit")," ", "")           
		End With
		
		If (Len(s_strCktDigit) = 0) Then
			crt.Dialog.MessageBox "Please enter values for all fields, or click the Cancel button.", "MISSING DATA", vbExclamation + vbOKOnly
			blnTryAgain = True 
		End If
	Loop While blnTryAgain

	s_objUtils.SendAndWaitFor "TRAVER TR " & s_strOrigTrkGrp & " " _
			& s_strCktDigit & strCarCode & s_strTwcTerm & " NT", _
			">", s_intTimeout
	
	TWC_LONG_DISTANCE_TRAVER = True 
End Function 'TWC_LONG_DISTANCE_TRAVER

'********'*********'*********'*********'*********'*********'*********'*********'
'TWC_INTERNATIONAL_TRAVER
'********'*********'*********'*********'*********'*********'*********'*********'
Function TWC_INTERNATIONAL_TRAVER()
	Dim blnTryAgain
	Dim strCarCode
	Dim strMsg

	If (Not DETERMINE_TWC_ORIG_TRUNK_GROUP()) Then
		TWC_INTERNATIONAL_TRAVER = False
		Exit Function
	End If 
	
	strMsg = "Enter Trunk Group from the results of the last traver that was "
	strMsg = strMsg & "performed on your screen.  The script will then "
	strMsg = strMsg & "perform the TWC Long Distance Traver."
	Do 
		blnTryAgain = False 
		With s_objDialogBox
			.Clear
			.Height = 200
			.Width = 350
			.LabelWidths = "120px"
			.Title = "TWC Long Distance International Traver"
			.AddItem Array("type=paragraph", "name=label1", "value=" & strMsg, "terminate=newline")
			.AddItem Array("type=text", "name=s_strOrigTrkGrp", "label=ORIG TRKGRP:", "accesskey=R", "value=" & s_strOrigTrkGrp, "size=13", "terminate=newline")
			.AddItem Array("type=text", "name=strCarCode", "label=Carrier Code:", "accesskey=a", "value=0333", "size=13", "terminate=newline")
			.Show
			If .Status <> vbOK Then 
				TWC_INTERNATIONAL_TRAVER = False 
				Exit Function  
			End If 
			s_strOrigTrkGrp = Replace(.Responses.Item("s_strOrigTrkGrp")," ", "")           
			strCarCode = .Responses.Item("strCarCode")           
		End With
		
		If (Len(s_strOrigTrkGrp) = 0 Or Len(strCarCode) = 0) Then
			crt.Dialog.MessageBox "Please enter values for all fields, or click the Cancel button.", "MISSING DATA", vbExclamation + vbOKOnly
			blnTryAgain = True 
		End If
	Loop While blnTryAgain

	s_objUtils.SendAndWaitFor "table cktdigit;pos "  _
			& strCarCode & " 1;lis 10;leave all", _
			">", s_intTimeout

	Do 
		blnTryAgain = False 
		With s_objDialogBox
			.Clear
			.Height = 280
			.Width = 350
			.LabelWidths = "120px"
			.Title = "Choose OZZ Code"
			.AddItem Array("type=paragraph", "name=label1", "value=" & strCarCode & " 1 is for International Direct Dialed - 011<br>" & vbcrlf & strCarCode + " 2 is for Operator Assisted Interational - 01", "terminate=newline")
			.AddItem Array("type=text", "name=s_strCktDigit", "label=Enter OZZ Code:", "accesskey=Z", "value=" & s_strCktDigit, "size=13", "terminate=newline")
			.Show
			If .Status <> vbOK Then 
				TWC_INTERNATIONAL_TRAVER = False 
				Exit Function  
			End If 
			s_strCktDigit = Replace(.Responses.Item("s_strCktDigit")," ", "")           
		End With
		
		If (Len(s_strCktDigit) = 0) Then
			crt.Dialog.MessageBox "Please enter values for all fields, or click the Cancel button.", "MISSING DATA", vbExclamation + vbOKOnly
			blnTryAgain = True 
		End If
	Loop While blnTryAgain

	s_objUtils.SendAndWaitFor "TRAVER TR " & s_strOrigTrkGrp & " " _
			& s_strCktDigit & strCarCode & "011" & s_strTwcTerm & " NT", _
			">", s_intTimeout

	TWC_INTERNATIONAL_TRAVER = True 
End Function 'TWC_INTERNATIONAL_TRAVER

'********'*********'*********'*********'*********'*********'*********'*********'
'INBOUND_NON_TWC_TRAVER
'********'*********'*********'*********'*********'*********'*********'*********'
Function INBOUND_NON_TWC_TRAVER()
	Dim blnTryAgain
	Dim strNextLine
	Dim strLRNCapture
	Dim i
	Dim x

	Do 
		blnTryAgain = False 
		With s_objDialogBox
			.Clear
			.Height = 180
			.Width = 350
			.LabelWidths = "120px"
			.Title = "Inbound Non-TWC Traver"
			.AddItem Array("type=paragraph", "name=label1", "value=Enter Trunk Group from the results of the outbound traver.", "terminate=newline")
			.AddItem Array("type=text", "name=s_strOrigTrkGrp", "label=ORIG TRKGRP:", "accesskey=R", "value=" & s_strOrigTrkGrp, "size=13", "terminate=newline")
			.AddItem Array("type=text", "name=s_strNonTwcTerm", "label=TERM #:", "accesskey=T", "value=" & s_strNonTwcTerm, "size=13", "terminate=newline")
			.Show
			If .Status <> vbOK Then 
				INBOUND_NON_TWC_TRAVER = False 
				Exit Function  
			End If 
			s_strOrigTrkGrp = Replace(.Responses.Item("s_strOrigTrkGrp") ," ", "")          
			s_strNonTwcTerm = s_objUtils.stripTN(.Responses.Item("s_strNonTwcTerm"))         
		End With
		
		If (Len(s_strOrigTrkGrp) = 0 Or Len(s_strNonTwcTerm) = 0) Then
			crt.Dialog.MessageBox "Please enter values for all fields, or click the Cancel button.", "MISSING DATA", vbExclamation + vbOKOnly
			blnTryAgain = True 
		End If
	Loop While blnTryAgain

	strLRNCapture = getLRNCapture(s_strNonTwcTerm)
	If (Len(strLRNCapture) = 0) Then
		INBOUND_NON_TWC_TRAVER = False 
		Exit Function  
	End If
	
	s_objUtils.SendAndWaitFor "traver tr "  & s_strOrigTrkGrp & " " _
			& strLRNCapture & " tcni " & s_strNonTwcTerm & " nt", _
			">", s_intTimeout
	INBOUND_NON_TWC_TRAVER = True 
End Function 'INBOUND_NON_TWC_TRAVER

'********'*********'*********'*********'*********'*********'*********'*********'
'INBOUND_TWC_TRAVER
'********'*********'*********'*********'*********'*********'*********'*********'
Function INBOUND_TWC_TRAVER()
	Dim blnTryAgain
	Dim strNextLine
	Dim strLRNCapture

	Do 
		blnTryAgain = False 
		With s_objDialogBox
			.Clear
			.Height = 180
			.Width = 350
			.LabelWidths = "120px"
			.Title = "Inbound TWC Traver"
			.AddItem Array("type=paragraph", "name=label1", "value=Enter Trunk Group from the results of the outbound traver.", "terminate=newline")
			.AddItem Array("type=text", "name=s_strOrigTrkGrp", "label=ORIG TRKGRP:", "accesskey=R", "value=" & s_strOrigTrkGrp, "size=13", "terminate=newline")
			.AddItem Array("type=text", "name=s_strTwcTerm", "label=TERM #:", "accesskey=T", "value=" & s_strTwcTerm, "size=13", "terminate=newline")
			.Show
			If .Status <> vbOK Then 
				INBOUND_TWC_TRAVER = False 
				Exit Function  
			End If 
			s_strOrigTrkGrp = Replace(.Responses.Item("s_strOrigTrkGrp")," ", "")          
			s_strTwcTerm = s_objUtils.stripTN(.Responses.Item("s_strTwcTerm"))           
		End With
		
		If (Len(s_strOrigTrkGrp) = 0 Or Len(s_strTwcTerm) = 0) Then
			crt.Dialog.MessageBox "Please enter values for all fields, or click the Cancel button.", "MISSING DATA", vbExclamation + vbOKOnly
			blnTryAgain = True 
		End If
	Loop While blnTryAgain
	
	strLRNCapture = getLRNCapture(s_strTwcTerm)
	If (Len(strLRNCapture) = 0) Then
		INBOUND_TWC_TRAVER = False 
		Exit Function  
	End If

'FIXME: SHouldn't this be a SendAndWaitFor?
	s_objUtils.SendAndWaitFor "traver tr "  & s_strOrigTrkGrp & " n cdn na " _
			& strLRNCapture & " ainres r01 lnpar " & s_strTwcTerm & " nt", _
			">", s_intTimeout

	INBOUND_TWC_TRAVER = true
End Function 'INBOUND_TWC_TRAVER

'********'*********'*********'*********'*********'*********'*********'*********'
'DETERMINE_TWC_ORIG_TRUNK_GROUP
'********'*********'*********'*********'*********'*********'*********'*********'
Function DETERMINE_TWC_ORIG_TRUNK_GROUP()
	Dim blnTryAgain
	Dim strNextLine
	Dim strLRNCapture
	Dim arrScreenData1
	Dim arrScreenData2
	Dim strMsg

	Do 
		blnTryAgain = False 
		With s_objDialogBox
			.Clear
			.Height = 160
			.Width = 350
			.LabelWidths = "120px"
			.Title = "Determine ORIG Trunk Group"
			.AddItem Array("type=text", "name=s_strTwcOrig", "label=Enter ORIG #:", "accesskey=R", "value=" & s_strTwcOrig, "size=13", "terminate=newline")
			.AddItem Array("type=text", "name=s_strLata", "label=LATA for ORIG #:", "accesskey=L", "value=" & s_strLata, "size=13", "terminate=newline")
			.AddItem Array("type=text", "name=s_strTwcTerm", "label=Enter TERM #", "accesskey=T", "value=" & s_strTwcTerm, "size=13", "terminate=newline")
			.Show
			If .Status <> vbOK Then 
				DETERMINE_TWC_ORIG_TRUNK_GROUP = False 
				Exit Function  
			End If 
			s_strTwcOrig = s_objUtils.stripTN(.Responses.Item("s_strTwcOrig"))           
			s_strLata = .Responses.Item("s_strLata")           
			s_strTwcTerm = Replace(s_objUtils.stripTN(.Responses.Item("s_strTwcTerm")),"+", "")           
		End With

		If (Len(s_strTwcOrig) = 0 Or Len(s_strLata) = 0 Or Len(s_strTwcTerm) = 0) Then
			crt.Dialog.MessageBox "Please enter values for all fields, or click the Cancel button.", "MISSING DATA", vbExclamation + vbOKOnly
			blnTryAgain = True 
		End If
	Loop While blnTryAgain
	
	strLRNCapture = getLRNCapture(s_strTwcOrig)
	If (Len(strLRNCapture) = 0) Then
		DETERMINE_TWC_ORIG_TRUNK_GROUP = False 
		Exit Function  
	End If

	arrScreenData1 = s_objUtils.SendAndReadStrings("table ssptkinf;lis all(2 eq "  _
			& s_strLata & ");leave all", _
			">", s_intTimeout)
	arrScreenData2 = s_objUtils.SendAndReadStrings("", _
			">", s_intTimeout)
	s_objArrays.WriteFile s_objArrays.AppendArray(arrScreenData1, arrScreenData2), DIR_NGVN_DATA & "TF_TRAVER.txt"
	s_objUtils.OpenFileImmediately = True
	s_objUtils.OpenFile DIR_NGVN_DATA & "TF_TRAVER.txt"

	strMsg = "Enter Trunk Group with the lowest number for area customer works out "
	strMsg = strMsg & "of to ensure routing is correct and trunk has been provisioned."
	Do 
		blnTryAgain = False 
		With s_objDialogBox
			.Clear
			.Height = 180
			.Width = 350
			.LabelWidths = "120px"
			.Title = "Choose Trunk Group"
			.AddItem Array("type=paragraph", "name=label1", "value=" & strMsg, "terminate=newline")
			.AddItem Array("type=text", "name=s_strSsptkTrkGrp", "label=Trunk Group:", "accesskey=T", "value=" & s_strSsptkTrkGrp, "size=13", "terminate=newline")
			.Show
			If .Status <> vbOK Then 
				DETERMINE_TWC_ORIG_TRUNK_GROUP = False 
				Exit Function  
			End If 
			s_strSsptkTrkGrp = Replace(.Responses.Item("s_strSsptkTrkGrp")," ", "")           
		End With
		
		If (Len(s_strSsptkTrkGrp) = 0) Then
			crt.Dialog.MessageBox "Please enter values for all fields, or click the Cancel button.", "MISSING DATA", vbExclamation + vbOKOnly
			blnTryAgain = True 
		End If
	Loop While blnTryAgain
	
	s_objUtils.SendAndWaitFor "table cktdigit;pos 0333 1;lis 10;leave all", ">", s_intTimeout

	Do 
		blnTryAgain = False 
		With s_objDialogBox
			.Clear
			.Height = 150
			.Width = 359
			.LabelWidths = "120px"
			.Title = "Choose OZZ Code"
			.AddItem Array("type=paragraph", "name=label1", "value=Enter OZZ Code for 0333 8.", "terminate=newline")
			.AddItem Array("type=text", "name=s_strCktDigit", "label=OZZ Code:", "accesskey=Z", "value=" & s_strCktDigit, "size=13", "terminate=newline")
			.Show
			If .Status <> vbOK Then 
				DETERMINE_TWC_ORIG_TRUNK_GROUP = False 
				Exit Function  
			End If 
			s_strCktDigit = Replace(.Responses.Item("s_strCktDigit")," ", "")           
		End With
		
		If (Len(s_strSsptkTrkGrp) = 0) Then
			crt.Dialog.MessageBox "Please enter values for all fields, or click the Cancel button.", "MISSING DATA", vbExclamation + vbOKOnly
			blnTryAgain = True 
		End If
	Loop While blnTryAgain

	s_objUtils.SendAndWaitFor "TRAVER TR " & s_strSsptkTrkGrp & " " & s_strCktDigit & "0333" & s_strTwcTerm & " NT", ">", s_intTimeout

	strMsg = "Now we will traver back to the LRN to determine the trunk group CLLI "
	strMsg = strMsg & "that will be used for the final traver.  Enter first route "
	strMsg = strMsg & "choice of last traver currently on the screen."
	Do 
		blnTryAgain = False 
		With s_objDialogBox
			.Clear
			.Height = 200
			.Width = 350
			.LabelWidths = "120px"
			.Title = "TRAVER Back to LRN"
			.AddItem Array("type=paragraph", "name=label1", "value=" & strMsg, "terminate=newline")
			.AddItem Array("type=text", "name=s_strTraverToOrigLrnTrkgrp", "label=LRN Trunk Group:", "accesskey=L", "value=" & s_strTraverToOrigLrnTrkgrp, "size=13", "terminate=newline")
			.Show
			If .Status <> vbOK Then 
				DETERMINE_TWC_ORIG_TRUNK_GROUP = False 
				Exit Function  
			End If 
			s_strTraverToOrigLrnTrkgrp = Replace(.Responses.Item("s_strTraverToOrigLrnTrkgrp")," ", "")           
		End With
		
		If (Len(s_strSsptkTrkGrp) = 0) Then
			crt.Dialog.MessageBox "Please enter values for all fields, or click the Cancel button.", "MISSING DATA", vbExclamation + vbOKOnly
			blnTryAgain = True 
		End If
	Loop While blnTryAgain
	
	s_objUtils.SendAndWaitFor "TRAVER TR " & s_strTraverToOrigLrnTrkgrp & " " & strLRNCapture & " NT", ">", s_intTimeout

	strMsg = "The trunk group that appeared from the results of the last Traver "
	strMsg = strMsg & "on your screen, is the trunk group that the customer will "
	strMsg = strMsg & "originate from.  Click ok to proceed to the next step."
	With s_objDialogBox
		.Clear
		.Height = 170
		.Width = 350
		.LabelWidths = "120px"
		.Title = "ORIG Trunk Group Results"
		.AddItem Array("type=paragraph", "name=label1", "value=" & strMsg, "terminate=newline")
		.Show
		If .Status <> vbOK Then 
			DETERMINE_TWC_ORIG_TRUNK_GROUP = False 
			Exit Function  
		End If 
	End With

	DETERMINE_TWC_ORIG_TRUNK_GROUP = True 
End Function 'DETERMINE_TWC_ORIG_TRUNK_GROUP

'********'*********'*********'*********'*********'*********'*********'*********'
'getLRNCapture
'********'*********'*********'*********'*********'*********'*********'*********'
Function getLRNCapture(strNumber)
	Dim arrScreenData
	Dim strScreenRow
	Dim intPos
	Dim strLRNCapture
	
	arrScreenData = s_objUtils.SendAndReadStrings("QLRN  " & strNumber, ">", s_intTimeout)
	s_objUtils.SendAndWaitFor "", ">", s_intTimeout

	For Each strScreenRow In arrScreenData
		intPos = InStr(strScreenRow, "Routing number: ")
		If (intPos > 0) Then 	'  capture LRN  
			strLRNCapture = Mid(strScreenRow, intPos + Len("Routing number: ") ,10)
		End If 	
	Next
	If (Len(Trim(strLRNCapture)) = 0) Then
		crt.Dialog.MessageBox "Can not find ""Routing number: "" on the screen.  Can not extract the LRN.  Execution terminating.", "MISSING DATA", vbCritical + vbOKOnly
	End If
	
	getLRNCapture = strLRNCapture
End Function 'getLRNCapture

'********'*********'*********'*********'*********'*********'*********'*********'
'NON_TWC_TOLLFREE_TRAVER
'********'*********'*********'*********'*********'*********'*********'*********'
Function NON_TWC_TOLLFREE_TRAVER()
	Dim blnTryAgain
	Dim strNonTwcTermOne
	Dim arrScreenData
	Dim intReturn
	Dim intMBReturn
	
	Do 
		blnTryAgain = False 
		With s_objDialogBox
			.Clear
			.Height = 150
			.Width = 350
			.LabelWidths = "120px"
			.Title = "NON_TWC_TOLLFREE_TRAVER"
			.AddItem Array("type=text", "name=s_strNonTwcOrig", "label=Enter ORIG #:", "accesskey=R", "value=" & s_strNonTwcOrig, "size=13", "terminate=newline")
			.AddItem Array("type=text", "name=s_strNonTwcTf", "label=Enter TF #:", "accesskey=T", "value=" & s_strNonTwcTf, "size=13", "terminate=newline")
			.Show
			If .Status <> vbOK Then 
				NON_TWC_TOLLFREE_TRAVER = False 
				Exit Function  
			End If 
			s_strNonTwcOrig = s_objUtils.stripTN(.Responses.Item("s_strNonTwcOrig"))           
			s_strNonTwcTf = s_objUtils.stripTN(.Responses.Item("s_strNonTwcTf"))           
		End With
		
		If (Len(s_strNonTwcOrig) = 0 Or Len(s_strNonTwcTf) = 0) Then
			crt.Dialog.MessageBox "Please enter values for all fields, or click the Cancel button.", "MISSING DATA", vbExclamation + vbOKOnly
			blnTryAgain = True 
		End If
	Loop While blnTryAgain

	strNonTwcTermOne = "1" & s_strNonTwcTf 

	arrScreenData = s_objUtils.SendAndReadStrings("TRAVER L " _
			& s_strNonTwcOrig & " " & strNonTwcTermOne & " B" _
			& DEFAULT_LINE_TERMINATOR, _
			">", s_intTimeout)
	s_objArrays.WriteFile arrScreenData, DIR_NGVN_DATA & "TF_TRAVER.txt" 
	s_objUtils.OpenFileImmediately = True
	s_objUtils.OpenFile DIR_NGVN_DATA & "TF_TRAVER.txt"
	
	Do 
		blnTryAgain = False 
		With s_objDialogBox
			.Clear
			.Height = 150
			.Width = 350
			.LabelWidths = "120px"
			.Title = "Enter Lata Name"
			.AddItem Array("type=paragraph", "name=label1", "value=Scroll through the Traver results in Notepad to find the entry for TABLE RATEAREA.", "terminate=newline")
			.AddItem Array("type=text", "name=s_strLataName", "label=LATA Name:", "accesskey=L", "value=" & s_strLataName, "size=13", "terminate=newline")
			.Show
			If .Status <> vbOK Then 
				NON_TWC_TOLLFREE_TRAVER = False 
				Exit Function  
			End If 
			s_strLataName = Replace(.Responses.Item("s_strLataName"), " ", "")          
		End With
		
		If (Len(s_strLataName) = 0) Then
			crt.Dialog.MessageBox "Please enter values for all fields, or click the Cancel button.", "MISSING DATA", vbExclamation + vbOKOnly
			blnTryAgain = True 
		End If
	Loop While blnTryAgain

	blnTryAgain = False 
	Do 
		intReturn = CInt(s_objUtils.SendAndWaitFor("TABLE LATANAME;LIS ALL (1 eq "  _
				& s_strLataName & ")", _
				Array ("BOTTOM", ">"), s_intTimeout))
		Select Case intReturn
			Case 0
				intMBReturn = crt.Dialog.MessageBox("The script timed out before the expected screen data was found.  Please check your connection and review the screen data  and decide to Abort (exit immediately), Retry (issue the previous command again) or Ignore (continue to the next screen).", "TIMEOUT!", vbAbortRetryIgnore + vbCritical)
				Select Case intMBReturn
					Case vbAbort
						NON_TWC_TOLLFREE_TRAVER = False 
						Exit Function  
					Case vbRetry
						blnTryAgain = True  
					Case vbIgnore
						blnTryAgain = False 
				End Select 
			Case 1
				s_objUtils.WaitForString ">", s_intTimeout
				blnTryAgain = False
			Case 2
				intMBReturn = crt.Dialog.MessageBox("The key word 'BOTTOM' was not found on the screen as was expected.  Please review the results of the previous command on the screen and decide to Abort (exit immediately), Retry (issue the previous command again) or Ignore (continue to the next screen).", "KEY WORD NOT FOUND!", vbAbortRetryIgnore + vbCritical)
				Select Case intMBReturn
					Case vbAbort
						NON_TWC_TOLLFREE_TRAVER = False 
						Exit Function  
					Case vbRetry
						blnTryAgain = True  
					Case vbIgnore
						blnTryAgain = False 
				End Select 
			End Select 
	Loop While blnTryAgain
	
	Do 
		blnTryAgain = False 
		With s_objDialogBox
			.Clear
			.Height = 150
			.Width = 350
			.LabelWidths = "120px"
			.Title = "Enter Lata Number"
			.AddItem Array("type=paragraph", "name=label1", "value=Enter the Lata Number.", "terminate=newline")
			.AddItem Array("type=text", "name=s_strLataNum", "label=LATA Number:", "accesskey=L", "value=" & s_strLataNum, "size=13", "terminate=newline")
			.Show
			If .Status <> vbOK Then 
				NON_TWC_TOLLFREE_TRAVER = False 
				Exit Function  
			End If 
			s_strLataNum = Replace(.Responses.Item("s_strLataNum"), " ", "")           
		End With

		If (Len(s_strLataNum) = 0) Then
			crt.Dialog.MessageBox "Please enter values for all fields, or click the Cancel button.", "MISSING DATA", vbExclamation + vbOKOnly
			blnTryAgain = True 
		End If
	Loop While blnTryAgain

	s_objUtils.SendAndWaitFor "e800ver " & s_strNonTwcOrig & " " _
			& s_strLataNum & " " & s_strNonTwcTf, _
			">", s_intTimeout

	Do 
		blnTryAgain = False 
		With s_objDialogBox
			.Clear
			.Height = 150
			.Width = 350
			.LabelWidths = "120px"
			.Title = "Enter Carrier Number"
			.AddItem Array("type=paragraph", "name=label1", "value=Review the results of the e800ver and enter the Carrier Number.  (e.g. 0333; enter 4 digits only)", "terminate=newline")
			.AddItem Array("type=text", "name=s_strCarrierNum", "label=Carrier Number:", "accesskey=N", "value=" & s_strCarrierNum, "size=13", "terminate=newline")
			.Show
			If .Status <> vbOK Then 
				NON_TWC_TOLLFREE_TRAVER = False 
				Exit Function  
			End If 
			s_strCarrierNum = Replace(.Responses.Item("s_strCarrierNum"), " ", "")           
		End With
		
		blnTryAgain = Not s_objUtils.TestLength(s_strCarrierNum, "Carrier Number", 4)
	Loop While blnTryAgain

	s_objUtils.SendAndWaitFor "TRAVER L "  & s_strNonTwcOrig & " 101" _
			& s_strCarrierNum & strNonTwcTermOne & " B", _
			">", s_intTimeout	
	crt.Dialog.MessageBox "Scroll up until you find the entry for table OFR4 to find the route destination.", "FIND DATA", vbInformation + vbOKOnly
	
	NON_TWC_TOLLFREE_TRAVER = true
End Function 'NON_TWC_TOLLFREE_TRAVER

'********'*********'*********'*********'*********'*********'*********'*********'
'TWC_TOLLFREE_TRAVER
'********'*********'*********'*********'*********'*********'*********'*********'
Function TWC_TOLLFREE_TRAVER()
	Dim blnTryAgain
	Dim strNextLine
	Dim strLRNCapture
	Dim strOzzCode
	Dim arrScreenData1 
	Dim arrScreenData2
	
	Do 
		blnTryAgain = False 
		With s_objDialogBox
			.Clear
			.Height = 150
			.Width = 350
			.LabelWidths = "120px"
			.Title = "Determine ORIG Trunk Group"
			.AddItem Array("type=text", "name=s_strTwcOrig", "label=Enter ORIG #:", "accesskey=R", "value=" & s_strTwcOrig, "size=13", "terminate=newline")
			.AddItem Array("type=text", "name=s_strLata", "label=LATA for ORIG #:", "accesskey=L", "value=" & s_strLata, "size=13", "terminate=newline")
			.AddItem Array("type=text", "name=s_strTwcTf", "label=Enter TF #:", "accesskey=T", "value=" & s_strTwcTf, "size=13", "terminate=newline")
			.Show
			If .Status <> vbOK Then 
				TWC_TOLLFREE_TRAVER = False 
				Exit Function  
			End If 
			s_strTwcOrig = s_objUtils.stripTN(.Responses.Item("s_strTwcOrig"))           
			s_strLata = .Responses.Item("s_strLata")           
			s_strTwcTf = Replace(s_objUtils.stripTN(.Responses.Item("s_strTwcTf")), "+", "")  'TODO: This is set, but then never used.  Why?  Can it be removed?         
		End With
		
		If (Len(s_strTwcOrig) = 0 Or Len(s_strLata) = 0 Or Len(s_strTwcTf) = 0) Then
			crt.Dialog.MessageBox "Please enter values for all fields, or click the Cancel button.", "MISSING DATA", vbExclamation + vbOKOnly
			blnTryAgain = True 
		End If
	Loop While blnTryAgain
	s_strTwcTerm = s_strTwcTf

	strLRNCapture = getLRNCapture(s_strTwcOrig)
	If (Len(strLRNCapture) = 0) Then
		TWC_TOLLFREE_TRAVER = False 
		Exit Function  
	End If
	
	arrScreenData1 = s_objUtils.SendAndReadStrings("table ssptkinf;lis all(2 eq " _
			& s_strLata & ");leave all", _
			">", s_intTimeout)
	arrScreenData2 = s_objUtils.SendAndReadStrings("", _
			">", s_intTimeout)
	s_objArrays.WriteFile s_objArrays.AppendArray(arrScreenData1, arrScreenData2), DIR_NGVN_DATA & "TF_TRAVER.txt"
	s_objUtils.OpenFileImmediately = True
	s_objUtils.OpenFile DIR_NGVN_DATA & "TF_TRAVER.txt"

	Do 
		blnTryAgain = False 
		With s_objDialogBox
			.Clear
			.Height = 170
			.Width = 350
			.LabelWidths = "120px"
			.Title = "Choose Trunk Group"
			.AddItem Array("type=paragraph", "name=label1", "value=Enter Trunk Group with the lowest number for area customer works out of to ensure routing is correct and trunk has been provisioned.", "terminate=newline")
			.AddItem Array("type=text", "name=s_strSsptkTrkGrp", "label=Trunk Group:", "accesskey=T", "value=" & s_strSsptkTrkGrp, "size=13", "terminate=newline")
			.Show
			If .Status <> vbOK Then 
				TWC_TOLLFREE_TRAVER = False 
				Exit Function  
			End If 
			s_strSsptkTrkGrp = Replace(.Responses.Item("s_strSsptkTrkGrp"), " ", "")           
		End With
		
		If (Len(s_strSsptkTrkGrp) = 0) Then
			crt.Dialog.MessageBox "Please enter values for all fields, or click the Cancel button.", "MISSING DATA", vbExclamation + vbOKOnly
			blnTryAgain = True 
		End If
	Loop While blnTryAgain

	s_objUtils.SendAndWaitFor "table cktdigit;pos 0333 1;lis 10;leave all", _
			">", s_intTimeout

	Do 
		blnTryAgain = False 
		With s_objDialogBox
			.Clear
			.Height = 150
			.Width = 350
			.Title = "Choose OZZ Code"
			.LabelWidths = "120px"
			.AddItem Array("type=paragraph", "name=label1", "value=Enter OZZ Code for 0333 8.", "terminate=newline")
			.AddItem Array("type=text", "name=s_strCktDigit", "label=OZZ Code:", "accesskey=Z", "value=" & s_strCktDigit, "size=13", "terminate=newline")
			.Show
			If .Status <> vbOK Then 
				TWC_TOLLFREE_TRAVER = False 
				Exit Function  
			End If 
			s_strCktDigit = Replace(.Responses.Item("s_strCktDigit"), " ", "")           
		End With
		
		If (Len(s_strCktDigit) = 0) Then
			crt.Dialog.MessageBox "Please enter values for all fields, or click the Cancel button.", "MISSING DATA", vbExclamation + vbOKOnly
			blnTryAgain = True 
		End If
	Loop While blnTryAgain

	s_objUtils.SendAndWaitFor "TRAVER TR " & s_strSsptkTrkGrp & " " _
			& s_strCktDigit & "0333" & s_strTwcTerm & " NT", _
			">", s_intTimeout

	Do 
		blnTryAgain = False 
		With s_objDialogBox
			.Clear
			.Height = 200
			.Width = 350
			.LabelWidths = "120px"
			.Title = "TRAVER Back to LRN"
			.AddItem Array("type=paragraph", "name=label1", "value=Now we will traver back to the LRN to determine the trunk group CLLI that will be used for the final traver.  Enter first route choice of last traver currently on the screen.", "terminate=newline")
			.AddItem Array("type=text", "name=s_strTraverToOrigLrnTrkgrp", "label=LRN Trunk Group:", "accesskey=L", "value=" & s_strTraverToOrigLrnTrkgrp, "size=13", "terminate=newline")
			.Show
			If .Status <> vbOK Then 
				TWC_TOLLFREE_TRAVER = False 
				Exit Function  
			End If 
			s_strTraverToOrigLrnTrkgrp = Replace(.Responses.Item("s_strTraverToOrigLrnTrkgrp"), " ", "")           
		End With
		
		If (Len(s_strTraverToOrigLrnTrkgrp) = 0) Then
			crt.Dialog.MessageBox "Please enter values for all fields, or click the Cancel button.", "MISSING DATA", vbExclamation + vbOKOnly
			blnTryAgain = True 
		End If
	Loop While blnTryAgain

	s_objUtils.SendAndWaitFor "TRAVER TR " & s_strTraverToOrigLrnTrkgrp _
			& " " & strLRNCapture & " NT", _
			">", s_intTimeout

	Do 
		blnTryAgain = False 
		With s_objDialogBox
			.Clear
			.Height = 200
			.Width = 350
			.LabelWidths = "130px"
			.Title = "ORIG Trunk Group Results"
			.AddItem Array("type=paragraph", "name=label1", "value=The trunk group that appeared from the results of the last Traver on your screen, is the trunk group that the customer will originate from.  Enter the trunk group value here.", "terminate=newline")
			.AddItem Array("type=text", "name=s_strOrigTwctfTrkGrp", "label=ORIG Trunk Group:", "accesskey=T", "value=" & s_strOrigTwctfTrkGrp, "size=13", "terminate=newline")
			.Show
			If .Status <> vbOK Then 
				TWC_TOLLFREE_TRAVER = False 
				Exit Function  
			End If 
			s_strOrigTwctfTrkGrp = .Responses.Item("s_strOrigTwctfTrkGrp")           
		End With
		
		If (Len(s_strOrigTwctfTrkGrp) = 0) Then
			crt.Dialog.MessageBox "Please enter values for all fields, or click the Cancel button.", "MISSING DATA", vbExclamation + vbOKOnly
			blnTryAgain = True 
		End If
	Loop While blnTryAgain

	s_objUtils.SendAndWaitFor "e800ver " & s_strTwcOrig & " " _
			& s_strLata & " " & s_strTwcTerm, _
			">", s_intTimeout

	Do 
		blnTryAgain = False 
		With s_objDialogBox
			.Clear
			.Height = 150
			.Width = 350
			.LabelWidths = "120px"
			.Title = "Enter Carrier Number"
			.AddItem Array("type=paragraph", "name=label1", "value=Enter the Carrier Number (e.g. 0222, 0333, etc)", "terminate=newline")
			.AddItem Array("type=text", "name=s_strCarrierNumber", "label=Carrier Number:", "accesskey=N", "value=" & s_strCarrierNumber, "size=13", "terminate=newline")
			.Show
			If .Status <> vbOK Then 
				TWC_TOLLFREE_TRAVER = False 
				Exit Function  
			End If 
			s_strCarrierNumber = Replace(.Responses.Item("s_strCarrierNumber"), " ", "")           
		End With
		
		If (Len(s_strCarrierNumber) = 0) Then
			crt.Dialog.MessageBox "Please enter values for all fields, or click the Cancel button.", "MISSING DATA", vbExclamation + vbOKOnly
			blnTryAgain = True 
		End If
	Loop While blnTryAgain

	s_objUtils.SendAndWaitFor "table nscdefs;lis all;leave all", _
			">", s_intTimeout

	Do 
		blnTryAgain = False 
		With s_objDialogBox
			.Clear
			.Height = 150
			.Width = 350
			.LabelWidths = "120px"
			.Title = "Enter NSCOZZ Code"
			.AddItem Array("type=paragraph", "name=label1", "value=Enter the numeric value after NSCOZZ (e.g. 051, 080, etc)", "terminate=newline")
			.AddItem Array("type=text", "name=strOzzCode", "label=NSCOZZ Code:", "accesskey=N", "value=" & strOzzCode, "size=13", "terminate=newline")
			.Show
			If .Status <> vbOK Then 
				TWC_TOLLFREE_TRAVER = False 
				Exit Function  
			End If 
			strOzzCode = Replace(.Responses.Item("strOzzCode"), " ", "")           
		End With
		
		If (Len(strOzzCode) = 0) Then
			crt.Dialog.MessageBox "Please enter values for all fields, or click the Cancel button.", "MISSING DATA", vbExclamation + vbOKOnly
			blnTryAgain = True 
		End If
	Loop While blnTryAgain

	s_objUtils.SendAndWaitFor "traver tr "  & s_strOrigTwctfTrkGrp & " " _
			& strOzzCode & s_strCarrierNumber & s_strTwcTerm & " NT", _
			">", s_intTimeout
	
	TWC_TOLLFREE_TRAVER = true
End Function 'TWC_TOLLFREE_TRAVER

'********'*********'*********'*********'*********'*********'*********'*********'
'display_clli_info
'********'*********'*********'*********'*********'*********'*********'*********'
Function display_clli_info()
	
End Function 'display_clli_info


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


