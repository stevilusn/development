#$language = "VBScript"
#$interface = "1.0"
option explicit
crt.Screen.Synchronous = True	

Sub Main
	dim strVlan
	dim result
	dim x 'temp place holder
	dim y 'temp place holder
	dim SessionName
	dim strCascade
	dim strCsrHostname

	strCascade = crt.Dialog.Prompt("Please enter the ALU Cascade: ", "Enter the cascade","")
	crt.Screen.WaitForCursor(2)
	crt.Screen.Send "show run int gi010 | include output aav" & chr(124) & " " & chr(13)
	crt.Screen.WaitForString "A:HNVAMDAY-IPA-01# "
	crt.Screen.ReadString()
	x = crt.screen.CurrentRow - 2
	strVlan = crt.screen.selection
	strVlan = crt.screen.Get(x,72, x, 75)
	msgbox strVlan

	'crt.Screen.WaitForString "A:"
	'crt.Screen.Send "show service sdp " & strVlan & " detail" & chr(13)
	'crt.Screen.WaitForString "A:"
	'crt.Screen.ReadString()
	'y = crt.screen.CurrentRow - 53
	'msgbox y
	'strCsrHostname = crt.screen.selection
	'strCsrHostname = crt.screen.Get(y,32, y, 46)
	'msgbox strCsrHostname
	'crt.Screen.WaitForString "*A:ALBYNYKJ-IPA-01# "
	
	'crt.Screen.WaitForString "Press any key to continue (Q to quit)" & chr(0)
	'crt.Screen.Send " show router interface " & chr(34) & strCsrHostname & chr(34) & " detail " & chr(13)
'	crt.Screen.WaitForString "A:"
	'crt.Screen.WaitForString "*A:ALBYNYKJ-IPA-01# "
	'crt.Screen.Send "show router arp " & chr(34) & "to-PNBSNY03-BHMH-1" & chr(34) & chr(13)
	'crt.Screen.WaitForString "*A:ALBYNYKJ-IPA-01# "
	'crt.Screen.Send "show router bfd session " & chr(124) & " match PNBSNY03-BHMH-1 context all" & chr(13)
	
	'If crt.Session.Connected Then crt.Session.Disconnect
	'SessionName =  crt.GetScriptTab.caption
	'msgbox SessionName

End Sub
