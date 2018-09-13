#$language = "VBScript"
#$interface = "1.0"
' SaveSelectedTextToCSVFile.vbs
'
' Description:
' If non-whitespace text is selected within the terminal screen, the user
' will be prompted for a location + filename in which to store the selected
' text as a CSV file. Multiple space characters within the data will be
' converted to a single comma character.
'
' Demonstrates:
' - How to use the Screen.Selection property new to SecureCRT 6.1 to access
' text selected within the terminal window.
' - How to use the Scripting.FileSystemObject to write data to a
' text file.
' - How to use RegExp's Replace() method to convert sequential space
' characters into a single comma character.
' - One way of determining if the script is running on Windows XP or not.
' - One way of displaying a file browse/open dialog in Windows XP
'
'
' g_nMode is used only if the user specifies a file that already exists, in
' which case the user will be prompted to overwrite the existing file, append
' to the existing file, or cancel the operation.
Dim g_nMode
Const ForWriting = 2
Const ForAppending = 8
' Be "tab safe" by getting a reference to the tab for which this script
' has been launched:
Set objTab = crt.GetScriptTab
Set g_shell = CreateObject("WScript.Shell")
Set g_fso = CreateObject("Scripting.FileSystemObject")
SaveSelectedTextToCSVFile
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Sub SaveSelectedTextToCSVFile()
	' Capture the selected text into a variable. The 'Selection' property
	' is available in SecureCRT 6.1 and later. This line of code will cause
	' an error if launched in a version of SecureCRT earlier than 6.1.
	strSelection = objTab.Screen.Selection
	' Check to see if the selection really has any text to save... we don't
	' usually want to write out nothing to a file.
	If Trim(strSelection) = "" Then
		crt.Dialog.MessageBox "Nothing to save!"
		Exit Sub
	End If
	strFilename = g_shell.SpecialFolders("MyDocuments") & _
	"\" & crt.Session.RemoteAddress & "-SelectedText.csv"
	Do
		strFilename = BrowseForFile(strFilename)
		If strFilename = "" Then Exit Sub
		' Do some sanity checking if the file specified by the user already
		' exists...
		If g_fso.FileExists(strFilename) Then
			nResult = crt.Dialog.MessageBox( _
			"Do you want to replace the contents of """ & _
			strFilename & _
			""" with the selected text?" & vbCrLf & vbCrLf & _
			vbTab & "Yes = overwrite" & vbCrLf & vbCrLf & _
			vbTab & "No = append" & vbCrLf & vbCrLf & _
			vbTab & "Cancel = end script" & vbCrLf, _
			"Replace Existing File?", _
			vbYesNoCancel)
			Select Case (nResult)
				Case vbYes
				g_nMode = ForWriting
				Exit Do
				Case vbNo
				g_nMode = ForAppending
				Exit Do
				Case Else
				Exit Sub
			End Select
		Else
			g_nMode = ForWriting
			Exit Do
		End If
	Loop
	' Automatically append a .csv if the filename supplied doesn't include
	' any extension.
	If g_fso.GetExtensionName(strFilename) = "" Then
		strFilename = strFilename & ".csv"
	End If
	' Replace instances of one or more space characters with a comma. Use
	' the VBScript built-in RegExp object's Replace method to perform this
	' task
	Set re = New RegExp
	' The pattern below means "one or more sequential space characters"
	re.Pattern = "[ ]+"
	re.Global = True
	re.MultiLine = True
	strCSVData = re.Replace(strSelection, ",")
	Do
		On Error Resume Next
		Set objFile = g_fso.OpenTextFile(strFilename, g_nMode, True)
		nError = Err.Number
		strErr = Err.Description
		On Error Goto 0
		If nError = 0 Then Exit Do
		' Display a message indicating there were problems opening
		' the file.
		nResponse = crt.Dialog.MessageBox( _
		"Failed to open """ & strFilename & """ (" & nError & "): " & _
		vbCrLf & vbTab & strErr & vbCrLf & vbCrLf & _
		"Check to see if the file is already open in another " & _
		"application and make sure you have permissions to " & _
		vbCrLf & "edit the file and create new files within the " & _
		"destination folder.", _
		"Save Operation Failed", _
		vbRetryCancel)
		If nResponse <> vbRetry Then Exit Sub
	Loop
	objFile.Write strCSVData & vbCrLf
	objFile.Close
	g_strMode = "Wrote"
	If g_nMode = ForAppending Then g_strMode = "Appended"
	crt.Dialog.MessageBox _
	g_strMode & " " & Len(strSelection) & " bytes to file:" & vbCrLf & _
	vbCrLf & strFilename
	' Now open the CSV file in the default .csv file application handler...
	g_shell.Run Chr(34) & strFilename & Chr(34)
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function BrowseForFile(strDefault)
	' Determine if we're running on Windows XP or not...
	Dim strOSName
	Set objWMIService = GetObject("winmgmts:" & _
	"{impersonationLevel=impersonate}!\\.\root\cimv2")
	Set colSettings = _
	objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
	For Each objOS In colSettings
		' Windows XP "Name" might look like this:
		' "Microsoft Windows XP Professional|C:\WINDOWS|\Device\Harddisk0\"...
		' Vista might appear as follows:
		' "Microsoft® Windows Vista™ Business |C:\Windows|\Device\Harddisk0\"...
		strOsName = Split(objOS.Name, "|")(0)
		Exit For
	Next
	If InStr(strOsName, "XP") > 0 Then
		' Based on information obtained from
		' http://blogs.msdn.com/gstemp/archive/2004/02/17/74868.aspx
		' NOTE: Will only work with Windows XP or newer since other OS's
		' don't have a UserAccounts.CommonDialog ActiveX
		' object registered.
		Set objDialog = CreateObject("UserAccounts.CommonDialog")
		objDialog.FileName = strDefault
		objDialog.Filter = "CSV Files|*.csv|Text Files|*.txt|All Files|*.*"
		objDialog.FilterIndex = 1
		objDialog.InitialDir = g_shell.SpecialFolders("MyDocuments")
		nResult = objDialog.ShowOpen
		If nResult <> 0 Then
			BrowseForFile = objDialog.FileName
		End If
	Else
		' On Windows other than XP, we'll just pop up an InputBox
		BrowseForFile = crt.Dialog.Prompt(_
		"Save selected text to file:", _
		"SecureCRT - Save Selected Text To File", _
		strDefault)
	End If
End Function
