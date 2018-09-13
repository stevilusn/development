#$language = "VBScript"
#$interface = "1.0"

crt.Screen.Synchronous = True

' This automatically generated script may need to be
' edited in order to work correctly.

Sub Main
	crt.Screen.Send "show config switch traffic-diagnostic-tools mac-swap-loopback  " & chr(8) & chr(8) & chr(13)
End Sub
