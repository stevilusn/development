#$language = "VBScript"
#$interface = "1.0"

crt.Screen.Synchronous = True

' This automatically generated script may need to be
' edited in order to work correctly.
Dim vStrPORT

Sub Main
	crt.Screen.Send "admin display-config match | " & vStrPORT & "context all" & chr(13)
End Sub
