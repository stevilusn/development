#$language = "VBScript"
#$interface = "1.0"

crt.Screen.Synchronous = True

' This automatically generated script may need to be
' edited in order to work correctly.
Dim vStrHostname

Sub Main
	crt.Screen.Send "show qos queue-group egress | match " & vStrHostname & chr(13)
End Sub
