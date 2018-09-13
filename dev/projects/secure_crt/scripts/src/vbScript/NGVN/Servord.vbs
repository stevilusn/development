'#$language = "VBScript"
'#$interface = "1.0"
'Name: Servord.vbs

Option Explicit

Dim s_objScriptTab
Dim s_objUtils
Dim s_objDialogBox
Dim s_objStrings
Dim s_objArrays
Dim s_intTimeout

Dim s_objQDN
Dim s_dicstrCommands

Dim s_strO_Tele
Dim s_strVM_Number
Dim s_strSeconds
Dim s_strProviderLspaoCode
Dim s_strDeleteOrAdd
Dim s_strCfw_Number
Dim s_strPicValue
Dim s_strLPicValue
Dim s_strIntPicValue
Dim s_strMtzvalue
Dim s_strEnduserstn

Dim s_blnDisplayResults

'********'*********'*********'*********'*********'*********'*********'*********'
'Main
'********'*********'*********'*********'*********'*********'*********'*********'
Sub Main
	Dim blnRet
	Dim objImport
	
	If (Not Initialize) Then
		Exit Sub
	End If
	s_objUtils.ScriptStart
	
	Set objImport = New clsImport
	objImport.File(DIR_SECURECRT & "clsQDN.vbs")
	Set objImport = Nothing 
	Set s_objQDN = New clsQDN

	
	Set s_dicstrCommands = CreateObject("scripting.dictionary")
		
	blnRet = s_objUtils.SendAndWaitFor("leave all", ">", s_intTimeout) 
	If (blnRet) Then 'This tests if you have access to the screen.
		If (get_user_data = vbOK) Then
			If (s_objQDN.GetData(s_strO_Tele)) Then 
				If (s_blnDisplayResults) Then
					display_results(s_objQDN)
				End If
				If (get_servord_data() = vbOK) Then
					AddOrDeleteFeatures s_strDeleteOrAdd
					s_objUtils.WriteMwqDn(s_strO_Tele)
				End If 
			End If 
		End If 
	Else
		s_objUtils.CancelMsgBox
	End If 
	
	s_objUtils.ScriptEnd
	Set s_objQDN = Nothing
	Set s_dicstrCommands = Nothing
	
	Set s_objScriptTab = Nothing
	Set s_objUtils = Nothing
	Set s_objDialogBox = Nothing
	Set s_objStrings = Nothing 
	Set s_objArrays = Nothing
End Sub 'Main

'********'*********'*********'*********'*********'*********'*********'*********'
'get_user_data
'********'*********'*********'*********'*********'*********'*********'*********'
Function get_user_data()
	Dim blnTryAgain
	
	s_strO_Tele = s_objUtils.ReadMwqDn()
	s_blnDisplayResults = False 
	
	blnTryAgain = True 
	Do 
		With s_objDialogBox
			.Clear
			.Height = 200
			.Width = 400
			.LabelWidths = "50px"
			.Title = "What is the DN"
			.AddItem Array("type=paragraph", "name=label1", "value=" & s_objUtils.CleanTrashMsg("DN"), "terminate=newline")
			.AddItem Array("type=text", "name=s_strO_Tele", "label=DN:", "accesskey=D", "value=" & s_strO_Tele, "size=10", "terminate=newline")
			.AddItem Array("type=html", "</br>")
			.AddItem Array("type=checkbox", "name=s_blnDisplayResults", _
					"label=Display QDN command Results?", "accesskey=Q", _
					"alignlabel=right", "labelwidth=230px", "value=" & s_blnDisplayResults, "terminate=newline")
			.Show
			get_user_data = .Status
			If .Status <> vbOK Then 
				Exit Function  
			End If 
			s_strO_Tele = s_objStrings.RemoveNonNumericCharacters(.Responses.Item("s_strO_Tele"))
			s_blnDisplayResults = CBool(.Responses.Item("s_blnDisplayResults"))
		End With
		'clean the spaces, underlines and dashes from the orig and term teles...   
		If (s_objUtils.TestLength(s_strO_Tele, "DN", 10)) Then 
			blnTryAgain = false
		End if
	Loop While blnTryAgain
	get_user_data = vbOK 
End Function 'get_user_data

'********'*********'*********'*********'*********'*********'*********'*********'
'get_servord_data
'********'*********'*********'*********'*********'*********'*********'*********'
Function get_servord_data()
	Dim blnTryAgain
	Dim strCompany
	Dim strCheckboxLabelWidth
	Dim strCheckboxAlignlabel
	
	strCheckboxLabelWidth = "labelwidth=60px"
	strCheckboxAlignlabel = "alignlabel=right"

	s_objUtils.SendAndWaitFor "leave all", ">", s_intTimeout 
	
	Do 
		blnTryAgain = False 
		With s_objDialogBox
			.Clear
			.Height = 500
			.Width = 550
			.LabelWidths = "100px"
			.Title = "Servord"
			.AddItem Array("type=html", "<b><a href='http://doccenter.nmcc.sprintspectrum.com/sites/Global_Voice/Resource%20Help/ResourceHelp_web/resourcehelp.htm' title='http://doccenter.nmcc.sprintspectrum.com/sites/Global_Voice/Resource%20Help/ResourceHelp_web/resourcehelp.htm'>SERVORD Help</a></b></br>")
			.AddItem Array("type=header", "name=header1", "label=ONLY USE THESE COMMANDS IF YOU HAVE RECEIVED TRAINING AND UNDERSTAND WHAT THEY ARE DOING !!!", "color=red")
			.AddItem Array("type=paragraph", "name=labelDN", "value=" & s_objUtils.CleanTrashMsg("DN"), "terminate=newline")
			.AddItem Array("type=text", "name=s_strO_Tele", "label=DN:", "accesskey=D", "value=" & s_strO_Tele, "size=10", "terminate=newline")
			.AddItem Array("type=radio", "name=strCompany", "label=Company:", "accesskey=p", "checked=" & strCompany, "choices=ADVR|ANTT|BAJA|BLRG|MDCM|MLNM|MSLN|MTCI|NPGC|SELCO|WAVE|WEHC|WOW", "terminate=newline")
			.AddItem Array("type=html", "<br>")
			.AddItem Array("type=radio", "name=s_strDeleteOrAdd", "label=Delete or Add:", "accesskey=A", "choices=Delete|Add", "checked=" & s_strDeleteOrAdd, "terminate=newline")
			.AddItem Array("type=html", "<div style='clear: both;'>Features: </div>")
			.AddItem Array("type=html", "<table width='100%'><tr><td>")
			.AddItem Array("type=checkbox", "name=bln3WC", strCheckboxAlignlabel, "label=3WC", strCheckboxLabelWidth, "checked=" & CStr(s_objQDN.Features.Item("3WC")))
			.AddItem Array("type=html", "</td><td>")
			.AddItem Array("type=checkbox", "name=blnACB", strCheckboxAlignlabel, 	"label=ACB", strCheckboxLabelWidth, "checked=" & CStr(s_objQDN.Features.Item("ACB")))
			.AddItem Array("type=html", "</td><td>")
			.AddItem Array("type=checkbox", "name=blnACRJ", strCheckboxAlignlabel, "label=ACRJ", strCheckboxLabelWidth, "checked=" & CStr(s_objQDN.Features.Item("ACRJ")))
			.AddItem Array("type=html", "</td><td>")
			.AddItem Array("type=checkbox", "name=blnAR", strCheckboxAlignlabel, 	"label=AR", strCheckboxLabelWidth, "checked=" & CStr(s_objQDN.Features.Item("AR")))
			.AddItem Array("type=html", "</td><td>")
			.AddItem Array("type=checkbox", "name=blnCCW", strCheckboxAlignlabel, 	"label=CCW", strCheckboxLabelWidth, "checked=" & CStr(s_objQDN.Features.Item("CCW")))
			.AddItem Array("type=html", "</td><td>")
			.AddItem Array("type=checkbox", "name=blnCFBL", strCheckboxAlignlabel, "label=CFBL", strCheckboxLabelWidth, "checked=" & CStr(s_objQDN.Features.Item("CFBL")))
			.AddItem Array("type=html", "</td></tr><tr><td>")
			.AddItem Array("type=checkbox", "name=blnCFDA", strCheckboxAlignlabel, "label=CFDA", strCheckboxLabelWidth, "checked=" & CStr(s_objQDN.Features.Item("CFDA")))
			.AddItem Array("type=html", "</td><td>")
			.AddItem Array("type=checkbox", "name=blnCFW", strCheckboxAlignlabel, 	"label=CFW", strCheckboxLabelWidth, "checked=" & CStr(s_objQDN.Features.Item("CFW")))
			.AddItem Array("type=html", "</td><td>")
			.AddItem Array("type=checkbox", "name=blnCFW-F", strCheckboxAlignlabel,"label=CFW-F", strCheckboxLabelWidth, "checked=" & CStr(s_objQDN.Features.Item("CFW-FIXED")))
			.AddItem Array("type=html", "</td><td>")
			.AddItem Array("type=checkbox", "name=blnCNAMD", strCheckboxAlignlabel,"label=CNAMD", strCheckboxLabelWidth, "checked=" & CStr(s_objQDN.Features.Item("CNAMD")))
			.AddItem Array("type=html", "</td><td>")
			.AddItem Array("type=checkbox", "name=blnCND", strCheckboxAlignlabel, 	"label=CND", strCheckboxLabelWidth, "checked=" & CStr(s_objQDN.Features.Item("CND")))
			.AddItem Array("type=html", "</td><td>")
			.AddItem Array("type=checkbox", "name=blnCNDB", strCheckboxAlignlabel, "label=CNDB", strCheckboxLabelWidth, "checked=" & CStr(s_objQDN.Features.Item("CNDB")))
			.AddItem Array("type=html", "</td></tr><tr><td>")
			.AddItem Array("type=checkbox", "name=blnCOT", strCheckboxAlignlabel, 	"label=COT", strCheckboxLabelWidth, "checked=" & CStr(s_objQDN.Features.Item("COT")))
			.AddItem Array("type=html", "</td><td>")
			.AddItem Array("type=checkbox", "name=blnCWT", strCheckboxAlignlabel, 	"label=CWT", strCheckboxLabelWidth, "checked=" & CStr(s_objQDN.Features.Item("CWT")))
			.AddItem Array("type=html", "</td><td>")
			.AddItem Array("type=checkbox", "name=blnDGT", strCheckboxAlignlabel, 	"label=DGT", strCheckboxLabelWidth, "checked=" & CStr(s_objQDN.Features.Item("DGT")))
			.AddItem Array("type=html", "</td><td>")
			.AddItem Array("type=checkbox", "name=blnINTPIC", strCheckboxAlignlabel,"label=INTPIC", strCheckboxLabelWidth, "checked=" & CStr(s_objQDN.Features.Item("INTPIC")))
			.AddItem Array("type=html", "</td><td>")
			.AddItem Array("type=checkbox", "name=blnLPIC", strCheckboxAlignlabel, "label=LPIC", strCheckboxLabelWidth, "checked=" & CStr(s_objQDN.Features.Item("LPIC")))
			.AddItem Array("type=html", "</td><td>")
			.AddItem Array("type=checkbox", "name=blnLSPAO", strCheckboxAlignlabel,"label=LSPAO", strCheckboxLabelWidth, "checked=" & CStr(s_objQDN.Features.Item("LSPAO")))
			.AddItem Array("type=html", "</td></tr><tr><td>")
			.AddItem Array("type=checkbox", "name=blnMTZ", strCheckboxAlignlabel, 	"label=MTZ", strCheckboxLabelWidth, "checked=" & CStr(s_objQDN.Features.Item("MTZ")))
			.AddItem Array("type=html", "</td><td>")
			.AddItem Array("type=checkbox", "name=blnMWT", strCheckboxAlignlabel, 	"label=MWT", strCheckboxLabelWidth, "checked=" & CStr(s_objQDN.Features.Item("MWT")))
			.AddItem Array("type=html", "</td><td>")
			.AddItem Array("type=checkbox", "name=blnPIC", strCheckboxAlignlabel, 	"label=PIC", strCheckboxLabelWidth, "checked=" & CStr(s_objQDN.Features.Item("PIC")))
			.AddItem Array("type=html", "</td><td>")
			.AddItem Array("type=checkbox", "name=blnPORT", strCheckboxAlignlabel, "label=PORT", strCheckboxLabelWidth, "checked=" & CStr(s_objQDN.Features.Item("PORT")))
			.AddItem Array("type=html", "</td><td>")
			.AddItem Array("type=checkbox", "name=blnSC1", strCheckboxAlignlabel, 	"label=SC1", strCheckboxLabelWidth, "checked=" & CStr(s_objQDN.Features.Item("SC1")))
			.AddItem Array("type=html", "</td><td>")
			.AddItem Array("type=checkbox", "name=blnSCA", strCheckboxAlignlabel, 	"label=SCA", strCheckboxLabelWidth, "checked=" & CStr(s_objQDN.Features.Item("SCA")))
			.AddItem Array("type=html", "</td></tr><tr><td>")
			.AddItem Array("type=checkbox", "name=blnSCF", strCheckboxAlignlabel, 	"label=SCF", strCheckboxLabelWidth, "checked=" & CStr(s_objQDN.Features.Item("SCF")))
			.AddItem Array("type=html", "</td><td>")
			.AddItem Array("type=checkbox", "name=blnSCRJ", strCheckboxAlignlabel, "label=SCRJ", strCheckboxLabelWidth, "checked=" & CStr(s_objQDN.Features.Item("SCRJ")))
			.AddItem Array("type=html", "</td><td>")
			.AddItem Array("type=checkbox", "name=blnSCWID", strCheckboxAlignlabel,"label=SCWID", strCheckboxLabelWidth, "checked=" & CStr(s_objQDN.Features.Item("SCWID")))
			.AddItem Array("type=html", "</td><td>")
			.AddItem Array("type=checkbox", "name=blnSPB", strCheckboxAlignlabel, 	"label=SPB", strCheckboxLabelWidth, "checked=" & CStr(s_objQDN.Features.Item("SPB")))			
			.AddItem Array("type=html", "</td> <td>")
			.AddItem Array("type=html", "</td> <td>")
			.AddItem Array("type=html", "</td></tr></table>")
			.AddItem Array("type=html", vbcrlf)
			.AddItem Array("type=text", "name=s_strCfw_Number", "label=CFW Fixed#:", "value=" & s_strCfw_Number, "size=13", "accesskey=F", "terminate=newline")
			.Show
			
			get_servord_data = .Status
			If .Status <> vbOK Then 
				Exit Function  
			End If 
			
			'*****  clean the spaces, underlines and dashes  
			s_strO_Tele = s_objStrings.RemoveNonNumericCharacters(.Responses.Item("s_strO_Tele"))	
			strCompany = .Responses.Item("strCompany")
			s_strDeleteOrAdd = .Responses.Item("s_strDeleteOrAdd")
			
			s_objQDN.Features.Item("3WC") = CBool(.Responses.Item("bln3WC")) 
			s_objQDN.Features.Item("ACB") = CBool(.Responses.Item("blnACB"))  
			s_objQDN.Features.Item("ACRJ") = CBool(.Responses.Item("blnACRJ"))  
			s_objQDN.Features.Item("AR") = CBool(.Responses.Item("blnAR"))  
			s_objQDN.Features.Item("CCW") = CBool(.Responses.Item("blnCCW"))  
			s_objQDN.Features.Item("CFBL") = CBool(.Responses.Item("blnCFBL"))  
			s_objQDN.Features.Item("CFDA") = CBool(.Responses.Item("blnCFDA"))  
			s_objQDN.Features.Item("CFW") = CBool(.Responses.Item("blnCFW"))  
			s_objQDN.Features.Item("CFW-FIXED") = CBool(.Responses.Item("blnCFW-F"))  
			s_objQDN.Features.Item("CNAMD") = CBool(.Responses.Item("blnCNAMD"))  
			s_objQDN.Features.Item("CND") = CBool(.Responses.Item("blnCND"))  
			s_objQDN.Features.Item("CNDB") = CBool(.Responses.Item("blnCNDB"))  
			s_objQDN.Features.Item("COT") = CBool(.Responses.Item("blnCOT"))  
			s_objQDN.Features.Item("CWT") = CBool(.Responses.Item("blnCWT"))  
			s_objQDN.Features.Item("DGT") = CBool(.Responses.Item("blnDGT"))  
			s_objQDN.Features.Item("INTPIC") = CBool(.Responses.Item("blnINTPIC"))  
			s_objQDN.Features.Item("LPIC") = CBool(.Responses.Item("blnLPIC"))  
			s_objQDN.Features.Item("LSPAO") = CBool(.Responses.Item("blnLSPAO"))  
			s_objQDN.Features.Item("MTZ") = CBool(.Responses.Item("blnMTZ"))  
			s_objQDN.Features.Item("MWT") = CBool(.Responses.Item("blnMWT"))  
			s_objQDN.Features.Item("PIC") = CBool(.Responses.Item("blnPIC"))  
			s_objQDN.Features.Item("PORT") = CBool(.Responses.Item("blnPORT")) 
			s_objQDN.Features.Item("SC1") = CBool(.Responses.Item("blnSC1"))  
			s_objQDN.Features.Item("SCA") = CBool(.Responses.Item("blnSCA"))  
			s_objQDN.Features.Item("SCF") = CBool(.Responses.Item("blnSCF"))  
			s_objQDN.Features.Item("SCRJ") = CBool(.Responses.Item("blnSCRJ"))  
			s_objQDN.Features.Item("SCWID") = CBool(.Responses.Item("blnSCWID"))
			s_objQDN.Features.Item("SPB") = CBool(.Responses.Item("blnSPB")) 
			
			s_strCfw_Number = s_objStrings.RemoveNonNumericCharacters(.Responses.Item("s_strCfw_Number"))
	
		End With
		
		If (Not s_objUtils.TestLength(s_strO_Tele, "DN", 10)) Then 
			blnTryAgain = True 
		End If 
		
		s_strVM_Number = setVMNumber(strCompany)
		If (s_strVM_Number = "") Then 
			crt.Dialog.MessageBox "Please select a Company.", "INVALID COMPANY", vbOKOnly + vbExclamation
			blnTryAgain = True 
		End If 
		
		If (s_objQDN.Features.Item("CFW") And s_objQDN.Features.Item("CFW-FIXED")) Then 
			crt.Dialog.MessageBox "Unable to add both cfw and cfw fixed feature to line", "CFW AND CFW-FIXED NOT ALLOWED", vbOKOnly + vbExclamation
			blnTryAgain = True 
		End If 
		
		If (s_strDeleteOrAdd <> "Add" And s_strDeleteOrAdd <> "Delete") Then
			crt.Dialog.MessageBox "Please select Add or Delete", "ADD OR DELETE", vbOKOnly + vbExclamation
			blnTryAgain = True 
		End If
	
		If (s_strDeleteOrAdd = "Add" And Len(s_strCfw_Number) = 0) And _
				(Not s_objQDN.Flags.Item("CFW-FIXED") And s_objQDN.Features.Item("CFW-FIXED")) Then 
			crt.Dialog.MessageBox "Please enter the phone number you want to forward the call to in the ""CFW Fixed#"" field.", "ENTER PHONE NUMBER", vbOKOnly + vbExclamation
			blnTryAgain = True 
		End If 
	Loop While blnTryAgain
	
	If (s_strDeleteOrAdd = "Add") Then 
		If (Not s_objQDN.Flags.Item("MTZ") And s_objQDN.Features.Item("MTZ")) Then
			get_servord_data = selectmtz()
			If (get_servord_data <> vbOK) Then
				Exit Function 
			End If 
		End If  
		
		If ((Not s_objQDN.Flags.Item("INTPIC") And s_objQDN.Features.Item("INTPIC")) _
				Or (Not s_objQDN.Flags.Item("PIC") And s_objQDN.Features.Item("PIC")) _
				Or (Not s_objQDN.Flags.Item("LPIC") And s_objQDN.Features.Item("LPIC"))) Then 
			get_servord_data = selectpic()
			If (get_servord_data <> vbOK) Then
				Exit Function 
			End If 
		End If
		
		If (Not s_objQDN.Flags.Item("CFDA") And s_objQDN.Features.Item("CFDA")) Then 
			get_servord_data = selectrings()
			If (get_servord_data <> vbOK) Then
				Exit Function 
			End If 
		End If
		
		If (Not s_objQDN.Flags.Item("SPB") And s_objQDN.Features.Item("SPB")) Then 
			get_servord_data = EnterEndUserTN()
			If (get_servord_data <> vbOK) Then
				Exit Function 
			End If 
		End If
			
		s_strProviderLspaoCode = setProviderLspaoCode(strCompany)	
		Call Create_Servord_Commands
	End If

End Function 'get_servord_data

'********'*********'*********'*********'*********'*********'*********'*********'
'setProviderLspaoCode
'********'*********'*********'*********'*********'*********'*********'*********'
Function setProviderLspaoCode(strCompany)
	Dim strProviderLspaoCode
	Select Case strCompany
	    Case "ADVR"
	    	strProviderLspaoCode = "AD00"
	    Case "ANTT"
	    	strProviderLspaoCode = "AT00"
	    Case "BAJA"
	    	strProviderLspaoCode = "BJ00"
	    Case "BLRG"
	    	strProviderLspaoCode = "BR00"
	    Case "MDCM"
	    	strProviderLspaoCode = "ME00"
	    Case "MLNM"
	    	strProviderLspaoCode = "ML00"
	    Case "MSLN"
	    	strProviderLspaoCode = "MA00"
	    Case "MTCI"
	    	strProviderLspaoCode = "MT00"
	    Case "NPGC"
	    	strProviderLspaoCode = "NP00"
	    Case "SELCO"
	    	strProviderLspaoCode = "SH00"
	    Case "WAVE"
	    	strProviderLspaoCode = "WB00"
	    Case "WEHC"
	    	strProviderLspaoCode = "WE00"
	    Case "WOW"	
	    	strProviderLspaoCode = "WW00"
	    Case Else
	    	strProviderLspaoCode = ""
	End Select
	setProviderLspaoCode = strProviderLspaoCode
End Function 'setProviderLspaoCode

'********'*********'*********'*********'*********'*********'*********'*********'
'setVM_Number
'********'*********'*********'*********'*********'*********'*********'*********'
Function setVMNumber(strCompany)
	Dim strVMNumber
	Select Case strCompany
	    Case "ADVR"
	    	strVMNumber = "18663076655"
	    Case "ANTT"
	    	strVMNumber = "18772687676"
	    Case "BAJA"
	    	strVMNumber = "18666283161"
	    Case "BLRG"
	    	strVMNumber = "18662935180"
	    Case "MDCM"
	    	strVMNumber = "18662732012"
	    Case "MLNM"
	    	strVMNumber = "18885801667"
	    Case "MSLN"
	    	strVMNumber = "18663310491"
	    Case "MTCI"
	    	strVMNumber = "18776828699"
	    Case "NPGC"
	    	strVMNumber = "18662732012"
	    Case "SELCO"
	    	strVMNumber = "18662384777"
	    Case "WAVE"
	    	strVMNumber = "18664481716"
	    Case "WEHC"
	    	strVMNumber = "18668978259"
	    Case "WOW"	
	    	strVMNumber = "18662426019"
	    Case Else
	    	strVMNumber = ""
	End Select
	setVMNumber = strVMNumber
End Function 'setVMNumber

'********'*********'*********'*********'*********'*********'*********'*********'
'selectmtz
'Called from get_servord_data.
'********'*********'*********'*********'*********'*********'*********'*********'
Function selectmtz() 
	Dim blnTryAgain

	Do 
		blnTryAgain = False 
		With s_objDialogBox
			.Clear
			.Height = 150
			.Width = 300
			.LabelWidths = "75px"
			.Title = "MTZ Selection"
			.AddItem Array("type=select", "name=s_strMtzvalue", "label=PIC:" ,"accesskey=P", "choices= |LOC|EST|CNT|MNT|PAC|AZT", "size=1", "multiple=False", "terminate=newline")
			.Show
			selectmtz = .Status
			If .Status <> vbOK Then 
				Exit Function  
			End If 
			s_strMtzvalue = .Responses.Item("s_strMtzvalue")           
		End With

		If (len(Trim(s_strMtzvalue)) = 0) Then 
			crt.Dialog.MessageBox "Please make a selection for PIC.", "MAKE A SELECTION", vbExclamation + vbOKOnly
			blnTryAgain = True 
		End If 
	Loop While blnTryAgain
End Function 'selectmtz

'********'*********'*********'*********'*********'*********'*********'*********'
'selectpic
'Called from get_servord_data.
'********'*********'*********'*********'*********'*********'*********'*********'
'TODO: Show only the fields that need to be updated.
Function selectpic() 
	Dim blnTryAgain

	s_strPicValue = "0333"
	s_strLPicValue = "0333"
	s_strIntPicValue = "0333"
	Do 
		blnTryAgain = False 
		With s_objDialogBox
			.Clear
			.Height = 150
			.Width = 300
			.LabelWidths = "75px"
			.Title = "PIC Selection"
			.AddItem Array("type=text", "name=s_strPicValue", "label=PIC:", "accesskey=P", "value=" & s_strPicValue, "size=10", "terminate=newline")
			.AddItem Array("type=text", "name=s_strLPicValue", "label=LPIC:", "accesskey=L", "value=" & s_strLPicValue, "size=10", "terminate=newline")
			.AddItem Array("type=text", "name=s_strIntPicValue", "label=INTPIC:", "accesskey=I", "value=" & s_strIntPicValue, "size=10", "terminate=newline")
			.Show
			selectpic = .Status
			If .Status <> vbOK Then 
				Exit Function  
			End If 
			s_strPicValue = .Responses.Item("s_strPicValue")
			s_strLPicValue = .Responses.Item("s_strLPicValue")
			s_strIntPicValue = .Responses.Item("s_strIntPicValue")
		End With

		If (len(Trim(s_strPicValue)) = 0 Or len(Trim(s_strLPicValue)) = 0 Or len(Trim(s_strIntPicValue)) = 0) Then 
			crt.Dialog.MessageBox "Please enter a value for all fields.", "MISSING DATA", vbExclamation + vbOKOnly
			blnTryAgain = True 
		End If 
	Loop While blnTryAgain

End Function 'selectpic

'********'*********'*********'*********'*********'*********'*********'*********'
'selectrings
'Called from get_servord_data.
'********'*********'*********'*********'*********'*********'*********'*********'
Function selectrings() 
	Dim blnTryAgain
	Dim strSeconds

	Do 
		blnTryAgain = False 
		With s_objDialogBox
			.Clear
			.Height = 310
			.Width = 300
			.LabelWidths = "25px"
			.Title = "Quantity of Rings Selection"
			.AddItem Array("type=html", "<u>S</u>elect # Rings for CFDA:<br>")
			.AddItem Array("type=radio", "name=strSeconds", "label=", "accesskey=S", "value=4 Rings - 24 Seconds", _
					"choices=2 Rings - 12 Seconds|3 Rings - 18 Seconds|4 Rings - 24 Seconds|5 Rings - 30 Seconds|6 Rings - 36 Seconds|7 Rings - 42 Seconds|8 Rings - 48 Seconds|9 Rings - 54 Seconds|10 Rings - 60 Seconds", _
					"termination=newline")
			.AddItem Array("type=html", "<div style='clear: both;'></div>")
			.Show
			selectrings = .Status
			If .Status <> vbOK Then 
				Exit Function  
			End If 
			strSeconds = .Responses.Item("strSeconds")
		End With
		If (len(Trim(strSeconds)) = 0) Then 
			crt.Dialog.MessageBox "Please make a selection.", "MISSING DATA", vbExclamation + vbOKOnly
			blnTryAgain = True 
		End If 
	Loop While blnTryAgain

	If (strSeconds = "2 Rings - 12 Seconds") Then s_strSeconds = "12"
	If (strSeconds = "3 Rings - 18 Seconds") Then s_strSeconds = "18"
	If (strSeconds = "4 Rings - 24 Seconds") Then s_strSeconds = "24"
	If (strSeconds = "5 Rings - 30 Seconds") Then s_strSeconds = "30"
	If (strSeconds = "6 Rings - 36 Seconds") Then s_strSeconds = "36"
	If (strSeconds = "7 Rings - 42 Seconds") Then s_strSeconds = "42"
	If (strSeconds = "8 Rings - 48 Seconds") Then s_strSeconds = "48"
	If (strSeconds = "9 Rings - 54 Seconds") Then s_strSeconds = "56"
	If (strSeconds = "10 Rings - 60 Seconds") Then s_strSeconds = "60"

End Function 'selectrings

'********'*********'*********'*********'*********'*********'*********'*********'
'EnterEndUserTN
'Called from get_servord_data.
'********'*********'*********'*********'*********'*********'*********'*********'
Function EnterEndUserTN() 
	Dim blnTryAgain

	Do 
		blnTryAgain = False 
		With s_objDialogBox
			.Clear
			.Height = 150
			.Width = 300
			.LabelWidths = "100px"
			.Title = "SPB End User Number"
			.AddItem Array("type=text", "name=s_strEnduserstn", "label=End Users TN:", "accesskey=T", "value=" & s_strEnduserstn, "size=10", "terminate=newline")
			.Show
			EnterEndUserTN = .Status
			If .Status <> vbOK Then 
				Exit Function  
			End If 
			s_strEnduserstn = .Responses.Item("s_strEnduserstn")
		End With

		If (len(Trim(s_strEnduserstn)) = 0) Then 
			crt.Dialog.MessageBox "Please enter a value for TN.", "MISSING DATA", vbExclamation + vbOKOnly
			blnTryAgain = True 
		End If 
	Loop While blnTryAgain
End Function 'EnterEndUserTN


'********'*********'*********'*********'*********'*********'*********'*********'
'Create_Servord_Commands
'********'*********'*********'*********'*********'*********'*********'*********'
Sub Create_Servord_Commands()	
	s_dicstrCommands("3WC") = 		"3WC"
	s_dicstrCommands("ACB") = 		"ACB NOAMA"
	s_dicstrCommands("ACRJ") = 	"ACRJ INACT"
	s_dicstrCommands("AR") = 		"AR NOAMA"
	s_dicstrCommands("CCW") = 		"ccw"
	s_dicstrCommands("CFBL") = 	"cfbl n nscr 10 " + s_strVM_Number
	s_dicstrCommands("CFDA") = 	"cfda n nscr 10 " + s_strSeconds + " fixring " + s_strVM_Number
	s_dicstrCommands("CFW") = 		"CFW C NSCR 1"
	s_dicstrCommands("CFW-FIXED") = 	"cfw f nscr 1 A " + s_strCfw_Number 
	s_dicstrCommands("CNAMD") = 	"CNAMD NOAMA"
	s_dicstrCommands("CND") = 		"CND NOAMA"
	s_dicstrCommands("CNDB") = 	"cndb noama"
	s_dicstrCommands("COT") = 		"cot noama"
	s_dicstrCommands("CWT") = 		"CWT"
	s_dicstrCommands("DGT") = 		"DGT"
	s_dicstrCommands("INTPIC") = 	"intpic " + s_strIntPicValue + " y"
	s_dicstrCommands("LPIC") = 	"lpic " + s_strLPicValue + " y"
	s_dicstrCommands("LSPAO") = 	"lspao " + s_strProviderLspaoCode + " U"
	s_dicstrCommands("MTZ") = 		"MTZ " + s_strMtzvalue
	s_dicstrCommands("MWT") = 		"mwt cmwi y n n y all y"
	s_dicstrCommands("PIC") = 		"pic " + s_strPicValue + " y"
	s_dicstrCommands("PORT") = 	"port "
	s_dicstrCommands("SC1") = 		"sc1"
	s_dicstrCommands("SCA") = 		"sca noama inact $ $ $"
	s_dicstrCommands("SCF") = 		"SCF NOAMA INACT $ 1 NSCR 2 NORING"
	s_dicstrCommands("SCRJ") = 	"SCRJ NOAMA INACT $"
	s_dicstrCommands("SCWID") = 	"scwid"
	s_dicstrCommands("SPB") = 		"spb " + s_strEndUsersTn	
End Sub	'Create_Servord_Commands

'********'*********'*********'*********'*********'*********'*********'*********'
'AddOrDeleteFeatures
'********'*********'*********'*********'*********'*********'*********'*********'
'TODO:  Add logic to prompt again if no features were added.
Sub AddOrDeleteFeatures(strDeleteOrAdd)
	If (strDeleteOrAdd = "Add") Then
		s_objScriptTab.Screen.Send "Servord; ado $ "  & s_strO_Tele & " "
		SendFeatures("Add")
	ElseIf (strDeleteOrAdd = "Delete") Then
		s_objScriptTab.Screen.Send "Servord; deo $ " & s_strO_Tele & " "
		SendFeatures("Delete")
	End If
	s_objUtils.SendAndWaitFor "$", ">", s_intTimeout
End Sub 'AddOrDeleteFeatures

Sub SendFeatures(strDeleteOrAdd)
	Dim strFeature
	Dim strSendString
	Dim blnProcess
	
	For Each strFeature In s_objQDN.Features.keys
		blnProcess = False 
		If (strDeleteOrAdd = "Add") Then
			blnProcess = ((Not s_objQDN.Flags.Item(strFeature)) And s_objQDN.Features.Item(strFeature))
		ElseIf (strDeleteOrAdd = "Delete") Then 
			blnProcess = (s_objQDN.Flags.Item(strFeature) And (Not s_objQDN.Features.Item(strFeature)))
		End If
		If (blnProcess) Then
			If (strDeleteOrAdd = "Add") Then
				strSendString = s_dicstrCommands(strFeature)
			ElseIf (strDeleteOrAdd = "Delete") Then 
				strSendString = strFeature
			End If
			s_objUtils.SendAndWaitFor strSendString, ">", s_intTimeout
		End If
	Next
End Sub 'SendFeatures

'********'*********'*********'*********'*********'*********'*********'*********'
'display_results
'TODO: This was commented out from Sub Main.  Why?  Can it be removed?
'********'*********'*********'*********'*********'*********'*********'*********'
Function display_results(objQDN)
	Dim strFileName
	Dim arrResults
	Dim objShell

	arrResults = objQDN.Report
	strFileName = DIR_NGVN_DATA & "QDN_Results.txt"

'	arrResults = s_objArrays.AppendVariant(arrResults, " QDN Results for DN " + s_strO_Tele)
'	arrResults = s_objArrays.AppendVariant(arrResults, "SNPA:" + s_objStrings.pack(s_objQDN.Data.Item("SNPA")," ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "   DN:" + s_objStrings.pack(s_objQDN.Data.Item("DN")," ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "  " + s_objStrings.pack(s_objQDN.Data.Item("PORTED-IN")," ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "  TYPE:" + s_objStrings.pack(s_objQDN.Data.Item("TYPE")," ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "   SIG:" + s_objStrings.pack(s_objQDN.Data.Item("SIG")," ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "LNATTIDX:" + s_objStrings.pack(s_objQDN.Data.Item("LNATTIDX")," ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "   LEN:" + s_objStrings.pack(s_objQDN.Data.Item("LINE EQUIPMENT NUMBER")," ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "   LINE CLASS CODE:" + s_objStrings.pack(s_objQDN.Data.Item("LINE CLASS CODE")," ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "IBN TYPE:" + s_objStrings.pack(s_objQDN.Data.Item("IBN TYPE")," ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "   CUSTGRP:" + s_objStrings.pack(s_objQDN.Data.Item("CUSTGRP")," ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "   SUBGRP:" + s_objStrings.pack(s_objQDN.Data.Item("SUBGRP")," ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "   NCOS:" + s_objStrings.pack(s_objQDN.Data.Item("NCOS")," ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "   LTG:" + s_objStrings.pack(s_objQDN.Data.Item("LINE TREATMENT GROUP")," ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "CARDCODE:" + s_objStrings.pack(s_objQDN.Data.Item("CARDCODE")," ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "   GND:" + s_objStrings.pack(s_objQDN.Data.Item("GND")," ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "   PADGRP:" + s_objStrings.pack(s_objQDN.Data.Item("PADGRP")," ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "   BNV:" + s_objStrings.pack(s_objQDN.Data.Item("BNV")," ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "   MNO:" + s_objStrings.pack(s_objQDN.Data.Item("BNV")," ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "PM NODE NUMBER:" + s_objStrings.pack(s_objQDN.Data.Item("PM NODE NUMBER")," ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "   PM TERMINAL NUMBER:" + s_objStrings.pack(s_objQDN.Data.Item("PM TERMINAL NUMBER")," ",1))
'	arrResults = s_objArrays.AppendVariant(arrResults, "   CFW INDEX:" + s_objStrings.pack(s_dicstrCommands("CFW")," ",1))
'	if (s_objQDN.Flags.Item("CWT")) Then
'		arrResults = s_objArrays.AppendVariant(arrResults, "CWT is built on the line")
'	End If 

	s_objArrays.WriteFile arrResults, strFileName
	s_objUtils.OpenFileImmediately = True 
	s_objUtils.OpenFile strFileName
End Function 'display_results

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