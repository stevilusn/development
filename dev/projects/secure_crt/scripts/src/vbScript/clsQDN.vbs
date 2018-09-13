'********'*********'*********'*********'*********'*********'*********'*********'
'FileName:	clsQDN.vbs
'Author:	Kent Garvis
'Purpose:	This class acts adds functionality for QDN data
'Dependency:
'	This Class is dependant on the s_objUtils being Set by the calling script.
'	This Class is dependant on the s_objStrings being Set by the calling script.
'	This Class is not intened to be used directly by SecureCRT, but to
'		be included/imported into and used by other SecureCRT scripts at runtime.
'Properties:
'  	...
'Methods:
'	...
'********'*********'*********'*********'*********'*********'*********'*********'
Option Explicit

Class clsQDN
	'********'*********'*********'*********'*********'*********'*********'*********'
	'DisplayMessages
	'********'*********'*********'*********'*********'*********'*********'*********'
	Private c_blnDisplayMessages
	Public Property Let DisplayMessages(blnDisplayMessages)
		c_blnDisplayMessages = blnDisplayMessages
	End Property
	Public Property Get DisplayMessages
		DisplayMessages =c_blnDisplayMessages
	End Property 
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'Data
	'********'*********'*********'*********'*********'*********'*********'*********'
	Private c_dicData
	Public Property Set Data(strData)
		Set c_dicData = dicData
	End Property
	Public Property Get Data
		set Data = c_dicData
	End Property 'Data
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'Flags
	'********'*********'*********'*********'*********'*********'*********'*********'
	Private c_dicFlags
	Public Property Set Flags(strFlags)
		Set c_dicFlags = dicFlags
	End Property
	Public Property Get Flags
		set Flags = c_dicFlags
	End Property 'Flags
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'Features
	'********'*********'*********'*********'*********'*********'*********'*********'
	Private c_dicFeatures
	Public Property Set Features(strFeatures)
		Set c_dicFeatures = dicFeatures
	End Property
	Public Property Get Features
		set Features = c_dicFeatures
	End Property 'Features
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'TeleNumber
	'********'*********'*********'*********'*********'*********'*********'*********'
	Private c_strTeleNumber
	Public Property Let TeleNumber(strTeleNumber)
		c_strTeleNumber = strTeleNumber
	End Property
	Public Property Get TeleNumber
		TeleNumber = c_strTeleNumber
	End Property 
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'Initialize and Terminate the class.
	'********'*********'*********'*********'*********'*********'*********'*********'
	Private Sub Class_Initialize   ' Setup Initialize event.
		c_blnDisplayMessages = True 
		Set c_dicFlags = CreateObject("scripting.dictionary")
		Set c_dicData = CreateObject("scripting.dictionary")
		Set c_dicFeatures = CreateObject("scripting.dictionary")
		
		Features.Item("3WC") = False 
		Features.Item("ACB") = False 
		Features.Item("ACRJ") = False 
		Features.Item("AR") = False 
		Features.Item("CCW") = False 
		Features.Item("CFBL") = False 
		Features.Item("CFDA") = False 
		Features.Item("CFW") = False 
		Features.Item("CFW-FIXED") = False 
		Features.Item("CNAMD") = False 
		Features.Item("CND") = False 
		Features.Item("CNDB") = False 
		Features.Item("COT") = False 
		Features.Item("CWT") = False 
		Features.Item("DGT") = False 
		Features.Item("INTPIC") = False 
		Features.Item("LPIC") = False 
		Features.Item("LSPAO") = False 
		Features.Item("MTZ") = False 
		Features.Item("MWT") = False 
		Features.Item("PIC") = False 
		Features.Item("PORT") = False
		Features.Item("SC1") = False 
		Features.Item("SCA") = False 
		Features.Item("SCF") = False 
		Features.Item("SCRJ") = False 
		Features.Item("SCWID") = False 
		Features.Item("SPB") = False
	End Sub 
	Private Sub Class_Terminate()    ' Setup Terminate event.
		Set c_dicFlags = Nothing
		Set c_dicData = Nothing
		Set c_dicFeatures = Nothing
	End Sub 
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'ScreenData
	'********'*********'*********'*********'*********'*********'*********'*********'
	Private c_arrstrScreenData
	Public Property Let ScreenData(arrstrScreenData)
		c_arrstrScreenData = arrstrScreenData
	End Property
	Public Property Get ScreenData
		ScreenData = c_arrstrScreenData
	End Property 
	
	Public Property Get ScreenDataAsString
		ScreenDataAsString = Join(ScreenData, vbCrLf)
	End Property 
	
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'GetData
	'********'*********'*********'*********'*********'*********'*********'*********'
	Public Function GetData(strTeleNumber)
		Dim strScreenRow
		Dim intPos 	
		
		TeleNumber = strTeleNumber
	
		s_objUtils.SendAndWaitFor "leave all", ">", s_intTimeout
		ScreenData = s_objUtils.SendAndReadStrings("QDN " & TeleNumber, ">", s_intTimeout)
	
		For Each strScreenRow In ScreenData
			If (InStr(strScreenRow, "UNASSIGNED") > 0) Then 
				If (blnDisplayMessages) Then  
					crt.Dialog.MessageBox "The DN is unassigned in the switch.", "ENDING EXECUTION", vbOKOnly & vbInformation
				End If 
				GetData = False 
				Exit Function 
			End if
			If (InStr(strScreenRow, "INVALID FOR THIS OFFICE") > 0) Then    
				If (blnDisplayMessages) Then  
					crt.Dialog.MessageBox "The DN is not valid for this switch.", "ENDING EXECUTION", vbOKOnly & vbInformation
				End If 
				GetData = False 
				Exit Function 
			End If
			
			'Parse Data
			If (InStr(strScreenRow, "BNV:") > 0) Then
				intPos = InStr(strScreenRow, " :")
				c_dicData.Item("BNV") = Mid(strScreenRow, intPos + 50, 2)
			End If 
			If (InStr(strScreenRow, "CARDCODE:") > 0) Then
				intPos = InStr(strScreenRow, " : ")
				c_dicData.Item("CARDCODE") = Mid(strScreenRow, intPos + 12, 6)
			End If 
			If (InStr(strScreenRow, "CUSTGRP:") > 0) Then
				intPos = InStr(strScreenRow, " : ")
				c_dicData.Item("CUSTGRP") = Mid(strScreenRow, intPos + 14, 10)
			End If 
			If (InStr(strScreenRow, "CFW INDEX:") > 0) Then
				intPos = InStr(strScreenRow, " :")
				c_dicData.Item("CFW INDEX") = Mid(strScreenRow, intPos + 11, 6)
			End If 
			If (InStr(strScreenRow, "DN:") > 0) Then    
				intPos = InStr(strScreenRow, ":")
				c_dicData.Item("DN") = Mid(strScreenRow, intPos + 6, 7)
			End If 
			If (InStr(strScreenRow, "GND:") > 0) Then
				intPos = InStr(strScreenRow, " :")
				c_dicData.Item("GND") = Mid(strScreenRow, intPos + 27, 2)
			End If 
			If (InStr(strScreenRow, "IBN TYPE:") > 0) Then
				intPos = InStr(strScreenRow, ":")
				if (intPos = 9) Then 
					c_dicData.Item("IBN TYPE") = Mid(strScreenRow, intPos + 1, 15)
				End If 
			End If 
			If (InStr(strScreenRow, "LINE CLASS CODE:") > 0) Then                           
				intPos = InStr(strScreenRow, " : ")
				c_dicData.Item("LINE CLASS CODE") = Mid(strScreenRow, intPos + 23, 15)
			End If 
			If (InStr(strScreenRow, "LINE EQUIPMENT NUMBER:") > 0) Then
				intPos = InStr(strScreenRow, " : ")
				c_dicData.Item("LINE EQUIPMENT NUMBER") = Mid(strScreenRow, intPos + 28, 25)
			End If 
			If (InStr(strScreenRow, "LINE TREATMENT GROUP:") > 0) Then
				intPos = InStr(strScreenRow, " : ")
				c_dicData.Item("LINE TREATMENT GROUP") = Mid(strScreenRow, intPos + 27, 6)
			End If 
			If (InStr(strScreenRow, "LNATTIDX:") > 0) Then
				intPos = InStr(strScreenRow, " : ")
				c_dicData.Item("LNATTIDX") = Replace(Mid(strScreenRow, intPos + 34, 15), " ", "")
			End If 
			If (InStr(strScreenRow, "MNO:") > 0) Then
				intPos = InStr(strScreenRow, " :")
				c_dicData.Item("MNO") = Mid(strScreenRow, intPos + 58, 2)
			End If 
			If (InStr(strScreenRow, "NCOS:") > 0) Then
				intPos = InStr(strScreenRow, " : ")
				c_dicData.Item("NCOS") = Mid(strScreenRow, intPos + 47, 3)
			End If 
			If (InStr(strScreenRow, "PADGRP:") > 0) Then
				intPos = InStr(strScreenRow, " :")
				c_dicData.Item("PADGRP") = Mid(strScreenRow, intPos + 38, 5)
			End If 
			If (InStr(strScreenRow, "PM NODE NUMBER     :") > 0) Then
				intPos = InStr(strScreenRow, " :")
				c_dicData.Item("PM NODE NUMBER") = Mid(strScreenRow, intPos + 6, 6)
			End If 
			If (InStr(strScreenRow, "PM TERMINAL NUMBER :") > 0) Then
				intPos = InStr(strScreenRow, " :")
				c_dicData.Item("PM TERMINAL NUMBER") = Mid(strScreenRow, intPos + 6, 6)
			End If 
			If (InStr(strScreenRow, "PORTED-IN") > 0) Then
				intPos = InStr(strScreenRow, "-")
				c_dicData.Item("PORTED-IN") = Mid(strScreenRow, intPos - 7, 11)
			End If 
			If (InStr(strScreenRow, "SIG:") > 0) Then
				intPos = InStr(strScreenRow, " : ")
				c_dicData.Item("SIG") = Mid(strScreenRow, intPos + 18, 3)
			End If 
			If (InStr(strScreenRow, "SNPA:") > 0) Then
				intPos = InStr(strScreenRow, ":")
				c_dicData.Item("SNPA") = Mid(strScreenRow, intPos + 2, 3)
			End If 
			If (InStr(strScreenRow, "SUBGRP:") > 0) Then     
				intPos = InStr(strScreenRow, " : ")
				c_dicData.Item("SUBGRP") = Mid(strScreenRow, intPos + 36, 10)
			End If 
			If (InStr(strScreenRow, "TYPE:") > 0) Then
				intPos = InStr(strScreenRow, ":")
				If (intPos = 5) Then
				  c_dicData.Item("TYPE") = Mid(strScreenRow, intPos + 2, 20)
				End If 
			End If 	
			
			'Parse Features. 			
			If (InStr(strScreenRow, "3WC") > 0 And InStr(strScreenRow, "U3WC") = 0) Then
				c_dicFlags.Item("3WC") = True
				c_dicFeatures.Item("3WC") = True 
			End If
			If (InStr(strScreenRow, "ACB") > 0) Then
				c_dicFlags.Item("ACB") = True
				c_dicFeatures.Item("ACB") = True
			End if
			If (InStr(strScreenRow, "ACRJ") > 0) Then
				c_dicFlags.Item("ACRJ") = True
				c_dicFeatures.Item("ACRJ") = True 
			End if
			If (InStr(strScreenRow, "AR ") > 0) Then
				c_dicFlags.Item("AR") = True
				c_dicFeatures.Item("AR") = True
			End if
			If (InStr(strScreenRow, "CCW") > 0) Then
				c_dicFlags.Item("CCW") = True
				c_dicFeatures.Item("CCW") = True
			End if
			If (InStr(strScreenRow, "CFBL") > 0) Then
				c_dicFlags.Item("CFBL") = True
				c_dicFeatures.Item("CFBL") = True 
			End if
			If (InStr(strScreenRow, "CFDA") > 0) Then
				c_dicFlags.Item("CFDA") = True
				c_dicFeatures.Item("CFDA") = True
			End if
			If (InStr(strScreenRow, "CFW C") > 0) Then
				c_dicFlags.Item("CFW") = True
				c_dicFeatures.Item("CFW") = True
			End if
			If (InStr(strScreenRow, "CFW F") > 0) Then
				c_dicFlags.Item("CFW-FIXED") = True
				c_dicFeatures.Item("CFW-FIXED") = True
			End If
			If (InStr(strScreenRow, "CNAMD") > 0) Then
				c_dicFlags.Item("CNAMD") = True
				c_dicFeatures.Item("CNAMD") = True
			End if
			If (InStr(strScreenRow, "CND NOAMA") > 0) Then
				c_dicFlags.Item("CND") = True
				c_dicFeatures.Item("CND") = True
			End if
			If (InStr(strScreenRow, "CNDB") > 0) Then
				c_dicFlags.Item("CNDB") = True
				c_dicFeatures.Item("CNDB") = True 
			End if
			If (InStr(strScreenRow, "COT") > 0) Then
				c_dicFlags.Item("COT") = True
				c_dicFeatures.Item("COT") = True
			End if
			If (InStr(strScreenRow, "CWT") > 0) Then
				c_dicFlags.Item("CWT") = True
				c_dicFeatures.Item("CWT") = True
			End if
			If (InStr(strScreenRow, "DGT") > 0) Then
				c_dicFlags.Item("DGT") = True
				c_dicFeatures.Item("DGT") = True
			End if
			If (InStr(strScreenRow, "INTPIC") > 0) Then
				c_dicFlags.Item("INTPIC") = True
				c_dicFeatures.Item("INTPIC") = True 
			End if
			If (InStr(strScreenRow, "LPIC") > 0) Then
				c_dicFlags.Item("LPIC") = True
				c_dicFeatures.Item("LPIC") = True 
			End if
			If (InStr(strScreenRow, "LSPAO") > 0) Then
				c_dicFlags.Item("LSPAO") = True
				c_dicFeatures.Item("LSPAO") = True 
			End if
			If (InStr(strScreenRow, "MTZ") > 0) Then
				c_dicFlags.Item("MTZ") = True
				c_dicFeatures.Item("MTZ") = True
			End if
			If (InStr(strScreenRow, "MWT") > 0) Then
				c_dicFlags.Item("MWT") = True
				c_dicFeatures.Item("MWT") = True
			End if
			If (InStr(strScreenRow, " PIC") > 0) Then
				c_dicFlags.Item("PIC") = True
				c_dicFeatures.Item("PIC") = True
			End if
			If (InStr(strScreenRow, "PORT ") > 0) Then
				c_dicFlags.Item("PORT") = True
				c_dicFeatures.Item("PORT") = True 
			End if
			If (InStr(strScreenRow, "SC1") > 0) Then
				c_dicFlags.Item("SC1") = True
				c_dicFeatures.Item("SC1") = True
			End if
			If (InStr(strScreenRow, "SCA") > 0) Then
				c_dicFlags.Item("SCA") = True
				c_dicFeatures.Item("SCA") = True
			End if
			If (InStr(strScreenRow, "SCF") > 0) Then
				c_dicFlags.Item("SCF") = True
				c_dicFeatures.Item("SCF") = True
			End if
			If (InStr(strScreenRow, "SCRJ") > 0) Then
				c_dicFlags.Item("SCRJ") = True
				c_dicFeatures.Item("SCRJ") = True 
			End if
			If (InStr(strScreenRow, "SCWID") > 0) Then
				c_dicFlags.Item("SCWID") = True
				c_dicFeatures.Item("SCWID") = True 
			End if
			If (InStr(strScreenRow, "SPB") > 0) Then
				c_dicFlags.Item("SPB") = True
				c_dicFeatures.Item("SPB") = True
			End if	
		
		Next 
		GetData = True 
	End Function 'GetData
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'Report
	'********'*********'*********'*********'*********'*********'*********'*********'
	Public Property Get Report()
		Dim arrResults
		Dim objReArr
		
		Set objReArr = New clsArrays
		arrResults = Array()
			
		arrResults = objReArr.AppendVariant(arrResults, " QDN Results for DN " + TeleNumber)
		arrResults = objReArr.AppendVariant(arrResults, "SNPA:" + s_objStrings.Pack(Data.Item("SNPA"), " ",1))
		arrResults = objReArr.AppendVariant(arrResults, "   DN:" + s_objStrings.Pack(Data.Item("DN"), " ",1))
		arrResults = objReArr.AppendVariant(arrResults, "  " + s_objStrings.Pack(Data.Item("PORTED-IN"), " ",1))
		arrResults = objReArr.AppendVariant(arrResults, "  TYPE:" + s_objStrings.Pack(Data.Item("TYPE"), " ",1))
		arrResults = objReArr.AppendVariant(arrResults, "   SIG:" + s_objStrings.Pack(Data.Item("SIG"), " ",1))
	
		arrResults = objReArr.AppendVariant(arrResults, "LNATTIDX:" + s_objStrings.Pack(Data.Item("LNATTIDX"), " ",1))
		arrResults = objReArr.AppendVariant(arrResults, "   LEN:" + s_objStrings.Pack(Data.Item("LINE EQUIPMENT NUMBER"), " ",1))
		arrResults = objReArr.AppendVariant(arrResults, "   LINE CLASS CODE:" + s_objStrings.Pack(Data.Item("LINE CLASS CODE"), " ",1))
	
		arrResults = objReArr.AppendVariant(arrResults, "IBN TYPE:" + s_objStrings.Pack(Data.Item("IBN TYPE"), " ",1))
		arrResults = objReArr.AppendVariant(arrResults, "   CUSTGRP:" + s_objStrings.Pack(Data.Item("CUSTGRP"), " ",1))
		arrResults = objReArr.AppendVariant(arrResults, "   SUBGRP:" + s_objStrings.Pack(Data.Item("SUBGRP"), " ",1))
		arrResults = objReArr.AppendVariant(arrResults, "   NCOS:" + s_objStrings.Pack(Data.Item("NCOS"), " ",1))
		arrResults = objReArr.AppendVariant(arrResults, "   LTG:" + s_objStrings.Pack(Data.Item("LINE TREATMENT GROUP"), " ",1))
	
		arrResults = objReArr.AppendVariant(arrResults, "CARDCODE:" + s_objStrings.Pack(Data.Item("CARDCODE"), " ",1))
		arrResults = objReArr.AppendVariant(arrResults, "   GND:" + s_objStrings.Pack(Data.Item("GND"), " ",1))
		arrResults = objReArr.AppendVariant(arrResults, "   PADGRP:" + s_objStrings.Pack(Data.Item("PADGRP"), " ",1))
		arrResults = objReArr.AppendVariant(arrResults, "   BNV:" + s_objStrings.Pack(Data.Item("BNV"), " ",1))
		arrResults = objReArr.AppendVariant(arrResults, "   MNO:" + s_objStrings.Pack(Data.Item("MNO"), " ",1))
	
		arrResults = objReArr.AppendVariant(arrResults, "PM NODE NUMBER:" + s_objStrings.Pack(Data.Item("PM NODE NUMBER"), " ",1))
		arrResults = objReArr.AppendVariant(arrResults, "   PM TERMINAL NUMBER:" + s_objStrings.Pack(Data.Item("PM TERMINAL NUMBER"), " ",1))
		arrResults = objReArr.AppendVariant(arrResults, "   CFW INDEX:" + s_objStrings.Pack(Data.Item("CFW INDEX"), " ",1))
		if (Flags.Item("CWT")) Then
			arrResults = objReArr.AppendVariant(arrResults, "CWT is built on the line")
		End If 
		
		Set objReArr = Nothing
		Report = arrResults
	End Property 'Report
	
	'********'*********'*********'*********'*********'*********'*********'*********'
	'ReportAsString
	'********'*********'*********'*********'*********'*********'*********'*********'
	Public Property Get ReportAsString
		ReportAsString = Join(Report, vbCrLf)
	End Property 'ReportAsString
End class
