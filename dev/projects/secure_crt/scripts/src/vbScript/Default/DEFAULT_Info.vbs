#$language = "VBScript"
#$interface = "1.0"
Option Explicit

Sub Main
	Dim objTab, objConfig
	Dim szHostname, szUsername, szSessionName, vPortForwards, nElements
	Dim strDescription, strPort, strFirewall
	Dim strMessage, strRet

	crt.Screen.Synchronous = True

	Set objTab = crt.GetScriptTab
	Set objConfig = objTab.Session.Config
	szUsername = objConfig.GetOption("Username")
	szHostname = objConfig.GetOption("Hostname")
	szSessionName = objTab.Session.Path
	strMessage = "Information for the current tabs session" & vbCrLf & vbCrLf
	strMessage = strMessage & "Session Name: " & szSessionName & vbCrLf
	strMessage = strMessage & "Host Name: " & szHostname & vbCrLf
	strMessage = strMessage & "Remote Address: " & objTab.Session.RemoteAddress & vbCrLf
	strMessage = strMessage & "Remote Port: " & objTab.Session.RemotePort & vbCrLf
	strMessage = strMessage & "Local Address: " & objTab.Session.LocalAddress & vbCrLf
	strMessage = strMessage & "Username: " & szUsername
	vPortForwards = objConfig.GetOption("Port Forward Table V2")
	nElements = UBound(vPortForwards)

	strMessage = strMessage & vbCRLF & vbCrLf
	strDescription = Join(objConfig.GetOption("Description"), vbcrlf)
	strMessage = strMessage & "Description: " & strDescription
	
	strMessage = strMessage & vbCRLF & vbCRLF
	If nElements = -1 Then
		strMessage = strMessage &  "No port forward configuration defined"
	Else
		strMessage = strMessage &  nElements + 1 & _
		" port forward entries exist in this session (" & _
		objTab.Session.Path & ")"
	End If

	strRet = crt.Dialog.MessageBox(strMessage, "SecureCRT: Current Tabs Session Information", vbInformation + vbOKOnly)
End Sub