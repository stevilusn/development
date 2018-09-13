#$language = "VBScript"
#$interface = "1.0"

crt.Screen.Synchronous = True

' This automatically generated script may need to be
' edited in order to work correctly.

Sub Main
	crt.Screen.Send "config port 1/2/7" & chr(13)
	crt.Screen.WaitForString "A:KNPLNC47-BHRS-1>config>port# "
	crt.Screen.Send "info" & chr(13)
End Sub
