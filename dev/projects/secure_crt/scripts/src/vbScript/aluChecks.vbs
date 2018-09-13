'#$language = "VBScript"
'#$interface = "1.0"
' aluChecks.vbs
' VBScript - This script will check for loopbacks on on port 1/2/7, using a pre-defined list of CSR IP's then save results in the specified excel spreadsheet.
' change the file name in line 14...
' Author S. Nelson
' Version 1.2 - May 2018
' ----------------------------------------------------' 
Option Explicit
crt.Screen.Synchronous = True

' This automatically generated script may need to be
' edited in order to work correctly.

Dim strUsername
Dim strPassword
Dim strCascade
Dim strHostName
Dim strVlan
Dim strCSRip

Const SITE_TITLE = "Sprint CBSA Power Scripts ver. 1.0.1"


Sub Main
	On Error Resume Next
	strHostname = InputBox("Enter Hostname: ", SITE_TITLE, "Please enter hostname")
	strUsername = InputBox("Enter Username: ", SITE_TITLE, "Please enter username")  
	strPassword = InputBox("Enter Password: ", SITE_TITLE, "Please enter Password")
	strCascade = InputBox("Enter Cascade: ", SITE_TITLE, "Please enter Cascade")
	strVlan = InputBox("Enter Vlan: ", SITE_TITLE, "Please enter Vlan")  
	strCSRip = InputBox("Enter CSR IP Address: ", SITE_TITLE, "Please enter CSR IP Address")

	
	'msgbox("Please enter credentials")
	'crt.Screen.WaitForString "Username: "
	crt.Screen.Send "rt948076" & chr(13)
	crt.Screen.WaitForString "Password: "
	crt.Screen.Send "V33AtlG#" & chr(13)
	crt.Screen.WaitForString "A:CLMBSCCI-IPA-01# "
	crt.Screen.Send "show router 2 interface " & chr(124) & " match CR03AW277 " & chr(13)
	crt.Screen.WaitForString "A:CLMBSCCI-IPA-01# "
	crt.Screen.Send "show servier" & chr(8) & chr(8) & "ce sdop" & chr(8) & chr(8) & "p 2633 detial" & chr(8) & chr(8) & chr(8) & "ail." & chr(8) & chr(13)
	crt.Screen.WaitForString "Press any key to continue (Q to quit)" & chr(0)
	crt.Screen.Send " show o" & chr(8) & "router interface " & chr(34) & "to-HAMRSCAA-BHRS-1" & chr(34) & " detail" & chr(13)
	crt.Screen.WaitForString "Press any key to continue (Q to quit)" & chr(0)
	crt.Screen.Send " show " & chr(27) & "[A" & chr(27) & "[Brouter ap" & chr(8) & "rp " & chr(34) & chr(34) & "to-HAMRSCAA-BHRS-1" & chr(34) & chr(27) & "[D" & chr(27) & "[D" & chr(27) & "[D" & chr(27) & "[D" & chr(27) & "[D" & chr(27) & "[D" & chr(27) & "[D" & chr(27) & "[D" & chr(27) & "[D" & chr(27) & "[D" & chr(27) & "[D" & chr(27) & "[D" & chr(27) & "[D" & chr(27) & "[D" & chr(27) & "[D" & chr(27) & "[D" & chr(27) & "[D" & chr(27) & "[D" & chr(27) & "[D" & chr(8) & chr(13)
	crt.Screen.WaitForString "A:CLMBSCCI-IPA-01# "
	crt.Screen.Send "show router bfd session " & chr(124) & " match to-HAMRSCAA-BHRS-1 contac" & chr(8) & chr(8) & "e" & chr(9) & "all" & chr(13)
	crt.Screen.WaitForString "A:CLMBSCCI-IPA-01# "
	crt.Screen.Send "ping 214.16.225.59" & chr(13)
	crt.Screen.WaitForString "A:CLMBSCCI-IPA-01# "
	crt.Screen.Send "show" & chr(8) & chr(8) & "ow sev" & chr(8) & "rvice sdp-usion" & chr(8) & chr(8) & "ng 5" & chr(8) & "2633" & chr(13)
	crt.Screen.WaitForString "A:CLMBSCCI-IPA-01# "
	crt.Screen.Send "show port 8/2/21" & chr(13)
	crt.Screen.WaitForString "Press any key to continue (Q to quit)" & chr(0)
	crt.Screen.Send " " & chr(27) & "[A" & chr(13)
	crt.Screen.WaitForString "Press any key to continue (Q to quit)" & chr(0)
	crt.Screen.Send " show qos queue-group egress " & chr(124) & " match HAMRSCAA-BHRS-1" & chr(13)
	crt.Screen.WaitForString "A:CLMBSCCI-IPA-01# "
	crt.Screen.Send chr(27) & "[A" & chr(8) & chr(8) & chr(8) & chr(8) & chr(8) & chr(8) & chr(8) & chr(8) & chr(8) & chr(8) & chr(8) & chr(8) & chr(8) & chr(8) & chr(8) & chr(8) & chr(34) & "HAMRSCAA-BHRS-1_100M-R11-7Q" & chr(34) & chr(13)
	crt.Screen.WaitForString "A:CLMBSCCI-IPA-01# "
	crt.Screen.Send "ssh admin@214.16.225.59" & chr(13)
	'crt.Screen.WaitForString "Are you sure you want to continue connecting (yes/no)? "
	'crt.Screen.Send "yes" & chr(13)
	crt.Screen.WaitForString "admin@214.16.225.59's password: "
	crt.Screen.Send "admin" & chr(13)
	crt.Screen.WaitForString "A:HAMRSCAA-BHRS-1# "
	crt.Screen.Send "show port 1/2/8" & chr(13)
	crt.Screen.WaitForString "Press any key to continue (Q to quit)" & chr(0)
	crt.Screen.Send "     " & chr(27) & "[A" & chr(8) & chr(8) & "7" & chr(13)
	crt.Screen.WaitForString "Press any key to continue (Q to quit)" & chr(0)
	crt.Screen.Send "     "
	On Error Goto 0

End Sub
