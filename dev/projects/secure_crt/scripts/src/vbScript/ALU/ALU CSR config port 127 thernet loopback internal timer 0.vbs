#$language = "VBScript"
#$interface = "1.0"

crt.Screen.Synchronous = True

' This automatically generated script may need to be
' edited in order to work correctly.

Sub Main
	crt.Screen.Send "configure port 1/2/7 ethernet loopback internal timer 0" & chr(13)
End Sub
