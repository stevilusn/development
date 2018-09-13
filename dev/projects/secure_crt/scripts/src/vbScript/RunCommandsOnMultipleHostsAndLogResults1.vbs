# $language = "VBScript"
# $interface = "1.0"

'''
    'Found at https://forums.vandyke.com/showthread.php?p=37294#post37294 
    'RunCommandsOnMultipleHostsAndLogResults.vbs
    '
    '   Last Modified: 21 May, 2018
    '     - If parent folder specified in the log file template doesn't
    '       exist, try to create it before attempting to write results
    '       or errors.
    '     - If unable to write to a results file or an error log file,
    '       the script will continue and attempt to perform the work it
    '       has been instructed to do -- regardless of ability to log
    '       errors/results. Errors will still be displayed when the
    '       script completes.
    '     - Allow host to be specified as either existing sessions in the
    '       session manager (in which case the existing session is used to
    '       establish the connection), or a hostname/IP which which an ad
    '       hoc connection is made. To force SecureCRT to make an ad hoc
    '       connection even if an existing session is found to match the
    '       host entry specified, set g_bUseExistingSessions = False below.
    '       If a host matches an existing session, and the session has
    '       a saved username and password (or Automate Logon option is
    '       enabled, the script will not send any credentials in the Connect()
    '       function defined below. Instead, it is expected that the session
    '       will authenticate itself using credentials stored within the session.
    '
    '     - Standardize logging to one function to avoid code duplication and
    '       facilitate consistency with error reporting message format.
    '
    '     - Ensure that errors are logged to the same file with "All(Errors)"
    '       in the name, even if the template log file name doesn't have the
    '       "IPADDRESS" substitution.
    '
    '     - Only respond to Password: prompts if they appear as the left-most
    '       item on the screen.
    '
    '   Last Modified: 17 May, 2018
    '     - Make it so that the global commands file is not needed if the hosts
    '       file has a host-specific command file specified for all hosts.

    '   Last Modified: 24 Apr, 2018
    '     - Added global boolean values to control logging to:
    '         --> One common file for all hosts (or individual files for each host)
    '         --> One common file for all cmds (or individual files for each cmd)
    '         --> Any combination of the above 2
    '     - Prevent inadvertent sending of usernames and passwords if the
    '       corresponding prompt text isn't located as the left-most word on the
    '       line.
    '     - Add support for specifying a unique command file specific to each host.
    '       To take advantage of this feature, your hosts.txt file should have this
    '       format:
    '          ---------------------------------------------------------------------
    '          hostname1;commandfileA.txt
    '          hostname2;commandfileA.txt
    '          hostname3;commandfileA.txt
    '          ipaddres1;commandFileB.txt
    '          ipaddres2;commandFileB.txt
    '          hostname4;commandFileB.txt
    '          ---------------------------------------------------------------------
    '       If a host-specific file is not specified for a host, the commands from
    '       the file specified in the global g_strCommandsFile variable are used.
    '       Host-specific command files can be specified as a fully-qualified (AKA
    '       "Absolute") file path or as a file name, in which case it is required
    '       that the file exist in the same location as the hosts file.
    '     - Add support for loading host/command files saved in either ANSI or
    '       Unicode formats. UTF-8 formats are not supported.
    '
    '   Last Modified: 13 Mar, 2015
    '     - Renamed script to better reflect work performed.
    '       Old name was:
    '          ReadDataFromHostFile-SendCommandsFromCommandsFile-LogResultsToIndividualFiles.vbs
    '       New name is:
    '          RunCommandsOnMultipleHostsAndLogResults.vbs
    '     - By default, logging is now done as separate log files per host, rather
    '       than per command. Beginning with this version of the script, each host
    '       targeted will have its own log file created, and all the commands run and
    '       their results will be logged to that host-specific log file.
    '     - Log file names will all share the same time stamp which correlates to the
    '       time at which this script was intially launched, instead of the time at
    '       which each command is run. This will help in collating log files in
    '       terms of instantiations of the script.
    '     - Errors like connection/authentication are now logged to a single
    '       "AllErrors" log file (same base name as the other log file(s)).
    '
    '   Last Modified: 6 Jun, 2013
    '     - Initial version published.
    '
    '
    ' DESCRIPTION:
    '   Demonstrates how to connect to hosts and send commands, logging results of
    '   each command to separate, uniquely-named files based on the host IP.
    '   Logging behavior can largely be controlled by changing the variable named
    '   "g_strLogFileTemplate" (examples shown below) to add/remove components. By
    '   default, this script will log commands and their results to individual
    '   files - one file per each host. However, there are two other log variations
    '   (or options) that can be achieved by merely changing the template pattern.
    '      Default Log Filename: Log all commands launched on a host to a separate
    '                            file, one file per host.
    '        g_strLogFileTemplate = g_strMyDocs & _
    '            "\##IPADDRESS--YYYY-MM-DD--hh'mm'ss.txt"
    '
    '      Log Option #1: Log everything (all hosts and commands) to a single file.
    '        g_strLogFileTemplate = g_strMyDocs & _
    '            "\##YYYY-MM-DD--hh'mm'ss.txt"
    '
    '      Log Option #2: Log hosts & commands in separate files; one file per cmd.
    '        g_strLogFileTemplate = g_strMyDocs & _
    '            "\##IPADDRESS--COMMAND--YYYY-MM-DD--hh'mm'ss.txt"
    '
    '
    '   Note: Host information is read from a file named "##hosts.txt" (located in
    '   "Documents" folder)
    '
    '   Note: Commands to be run on each remote host are read in from a file named
    '   "##commands.txt" (also located in "Documents" folder)
    '
    '   This script does not interfere with a session's logging settings; instead,
    '   the crt.Screen.ReadString() method is used to capture the output of a
    '   command, and then the built-in Windows FileSystemObject is used to write
    '   the captured data to a file.
'''

Set g_fso = CreateObject("Scripting.FileSystemObject")
Set g_shell = CreateObject("WScript.Shell")

g_strMyDocs = g_shell.SpecialFolders("MyDocuments")

' ##hosts.txt and ##commands.txt files are located in the current user's
' ##MyDocuments folder.  Hard-code to different paths as your needs dictate.
g_strHostsFile    = g_strMyDocs & "\3-Performance\Source Files For Scripts\2.5_Audit_hosts.txt"
g_strCommandsFile = g_strMyDocs & "\3-Performance\Source Files For Scripts\2.5_Audit_Commands.txt"
'C:\Users\crl8036\OneDrive - Sprint\3-Performance\Source Files for Scripts
' Template used for formulating the name of the results file in which the
' command output will be saved.  You can choose to arrange the various
' components (i.e. "IPADDRESS", "COMMAND", "YYYY", "MM", "DD", "hh", etc.) of
' the log file name into whatever order you want it to be.
'g_strLogFileTemplate = _
'g_strMyDocs & "\Log\##IPADDRESS-COMMAND-YYYY-MM-DD--hh'mm'ss.txt"

'      Log Option #1: Log everything (all hosts and commands) to a single file.
g_strLogFileTemplate = g_strMyDocs & "\3-Performance\results\2.5 Audit\2.5_Audit_YYYY-MM-DD--hh'mm'ss.txt"

g_bLogToIndividualFilesPerHost = False
g_bLogToIndividualFilesPerCommand = False

' Add Time stats to the log file name based on the Template
' defined by the script author.
g_strLogFileTemplate = Replace(g_strLogFileTemplate, "YYYY-", Year(Date) & "-")
g_strLogFileTemplate = Replace(g_strLogFileTemplate, "-MM-", "-" & NN(Month(Date)) & "-")
g_strLogFileTemplate = Replace(g_strLogFileTemplate, "-DD-", "-" & NN(Day(Date)) & "-")
g_strLogFileTemplate = Replace(g_strLogFileTemplate, "-hh'", "-" & NN(Hour(Time)) & "'")
g_strLogFileTemplate = Replace(g_strLogFileTemplate, "'mm'", "'" & NN(Minute(Time)) & "'")
g_strLogFileTemplate = Replace(g_strLogFileTemplate, "'ss", "'" & NN(Second(Time)))

' Comment character allows for comments to exist in either the host.txt or
' commands.txt files. Lines beginning with this character will be ignored.
g_strComment = "#"

' If connecting through a proxy is required, comment out the second statement
' below, and modify the first statement below to match the name of the firewall
' through which you're connecting (as defined in global options within
' SecureCRT)
'g_strFirewall = " /FIREWALL=myFireWallName "
'g_strFirewall = ""

' Username for authenticating to the remote system
g_strUsername = "admin"
' Password for authenticating to the remote system
g_strPassword = "admin"

' Global variable for housing details of any errors that might be encountered
g_strError = ""

' Constants used for reading and writing files.
Const ForReading   = 1
Const ForWriting   = 2
Const ForAppending = 8

Dim g_objNewTab, g_strHost
Dim g_bUseExistingSessions
g_bUseExistingSessions = True

'#####################################################################################################################
' Create an Excel workbook/worksheet
    '
    'Dim app, wb, ws
    'Set app = CreateObject("Excel.Application")
    'Set wb = app.Workbooks.Add
    'Set ws = wb.Worksheets(1)

    'Dim row, screenrow, readline, items
    'row = 1
    
    'Dim objWb, strErrors
    'Dim strWorkBookPath
    'strWorkBookPath = "C:\Users\rt948076\OneDrive - Sprint\3-Performance\results\2.5 Audit\chart.xlsx"

    '
'#####################################################################################################################



' Call the main subroutine named "MainSub"
MainSub

'-------------------------------------------------------------------------------
Sub MainSub()



' Create arrays in which lines from the hosts and commands files will be
' stored.
	Dim vHosts(), vCommands()
' Create variables for storing information about the lines read from the
' file
	Dim nHostCount, nCommandCount, nCommentLines
' Call the ReadDataFromFile() function defined in this script.  It will
' read in the hosts file and populate an array with non-comment lines that
' will represent the hosts to which connections will be made later on.
	If Not ReadDataFromFile(g_strHostsFile, _
        g_strComment, _
        vHosts, _
        nHostCount, _
        nCommentLines) Then
        DisplayMessage g_strError
        Exit Sub
    End If

Dim strErrors, strSuccesses

' Now call the ReadDataFromFile() function for the commands file.
If Not g_fso.FileExists(g_strCommandsFile) Then
	strError = "WARNING: Default commands file does not contain any " & _
	"commands: " & g_strCommandsFile & "." & vbcrlf & _
	"Any lines in the hosts file which don't include a " & _
	"commands file specification will fail."
	strErrors = CaptureError(strErrors, strError)
End If


' Before attempting any connections, ensure that the "Auth Prompts In
' Window" option in the Default session is already enabled.  If not, prompt
' the user to have the script enable it automatically before continuing.
' Before continuing on with the script, ensure that the session option
' for handling authentication within the terminal window is enabled
Set objConfig = crt.OpenSessionConfiguration("Default")
bAuthInTerminal = objConfig.GetOption("Auth Prompts In Window")
If Not bAuthInTerminal Then
	strMessage = _
	"The 'Default' session (used for all ad hoc " & _
	"connections) does not have the ""Display logon prompts in " & _
	"terminal window"" option enabled, which is required for this " & _
	"script to operate successfully." & vbcrlf & vbcrlf
	If PromptYesNo(_
	strMessage & _
	"Would you like to have this script automatically enable this " & _
	"option in the 'Default' session so that next time you run " & _
	"this script, the option will already be enabled?") <> vbYes Then
	Exit Sub
End If

' User answered prompt with Yes, so let's set the option and save
objConfig.SetOption "Auth Prompts In Window", True
objConfig.Save
End If


' Iterate through each element of our vHosts array...
For nIndex = 0 To nHostCount - 1
' Arrays are indexed starting at 0 (zero)

' Store the current host in a variable so we don't have to remember
' what "vHosts(nIndex)" means.
g_strHost = vHosts(nIndex)

' Exit the loop if the host name is empty (this means we've
' reached the end of our array
If g_strHost = "" Then Exit For
    strCommandsFilename = ""
    bContinue = True
    If Instr(g_strHost, ";") > 0 Then
        vHostElems = Split(g_strHost, ";")
        g_strHost = vHostElems(0)
        strCommandsFilename = vHostElems(1)
        
        If strCommandsFilename = "" Then
            strCommandsFilename = g_strCommandsFile
        End If
        
        If Not g_fso.FileExists(strCommandsFilename) Then
            ' Check if the file path is relative to where the hosts.txt
            ' file exists...
            strCommandsFilename = g_fso.GetParentFolderName(g_strHostsFile) & "\" & strCommandsFilename
            If Not g_fso.FileExists(strCommandsFilename) Then
                strError = "Host-specific command file not found for host '" & g_strHost & "': " & strCommandsFilename
                strErrors = CaptureError(strErrors, strError)
                bContinue = False
            End If
        End If
        
        If bContinue Then
            g_strError = ""
            If Not ReadDataFromFile(strCommandsFilename, _
            g_strComment, _
            vCommands, _
            nCommandCount, _
            nCommentLines) Then
            
            strError = "Error attempting to read host-specific file for host '" & _
            g_strHost & "': " & g_strError
            strErrors = CaptureError(strErrors, strError)
            bContinue = False
        End If
    End If
Else
    ' If we've had any host-specific command files change, we'll need
    ' to reload the common commands file to avoid running the last-known
    ' host-specific commands with this new host that doesn't have a
    ' commands file specified; use the default commands file.
    If Not ReadDataFromFile(g_strCommandsFile, _
        g_strComment, _
        vCommands, _
        nCommandCount, _
        nCommentLines) Then
        strError = "While working on host '" & g_strHost & "... " & g_strError
        strErrors = CaptureError(strErrors, strError)
        bContinue = False
    End If
End If

If bContinue Then
    ' Build up a string containing connection information and options.
    ' /ACCEPTHOSTKEYS should only be used if you suspect that there might be
    ' hosts in the hosts.txt file to which you haven't connected before, and
    ' therefore SecureCRT hasn't saved out the SSH2 server's host key.  If
    ' you are confused about what a host key is, please read through the
    ' white paper:
    '    http://www.vandyke.com/solutions/host_keys/index.html
    '
    '   A best practice would be to connect manually to each device,
    '   verifying each server's hostkey individually before running this
    '   script.
    '
    ' If you want to authenticate with publickey authentication instead of
    ' password, in the assignment of strConnectString below, replace:
    '   " /AUTH password,keyboard-interactive /PASSWORD " & g_strPassword & _
    ' with:
    '   " /AUTH publickey /I ""full_path_to_private_key_file"" " & _

    If g_bUseExistingSessions and SessionExists(g_strHost) Then
        strConnectString = "/S " & g_strHost
    Else
        strConnectString = _
        g_strFirewall & _
        " /SSH2 " & _
        " /L " & g_strUsername & _
        " /AUTH password,keyboard-interactive /PASSWORD " & g_strPassword & _
        " " & g_strHost

        strConnectString = _
        g_strFirewall & _
        " /SSH2 " & _
        " " & g_strHost
    End If

        ' Call the Connect() function defined below in this script.  It handles
        ' the connection process, returning success/fail.
    If Not Connect(strConnectString) Then
        strError = "Failed to connect to " & g_strHost & _
        ": " & g_strError
        strErrors = CaptureError(strErrors, strError)
    Else
        ' If we get to this point in the script, we're connected (including
        ' authentication) to a remote host successfully.
        g_objNewTab.Screen.Synchronous = True
        g_objNewTab.Screen.IgnoreEscape = True

            ' Once the screen contents have stopped changing (polling every
            ' 350 milliseconds), we'll assume it's safe to start interacting
            ' with the remote system.

                If Not WaitForScreenContentsToStopChanging(350) Then
                    strError = "Error: " & _
                    "Failed to detect remote ready status for host: " & _
                    g_strHost & ".  " & g_strError
                    strErrors = CaptureError(strErrors, strError)
                Else
                    ' Get the shell prompt so that we can know what to look for when
                    ' determining if the command is completed. Won't work if the
                    ' prompt is dynamic (e.g. changes according to current working
                    ' folder, etc)
                    nRow = g_objNewTab.Screen.CurrentRow
                    strPrompt = g_objNewTab.screen.Get(nRow, 0, nRow, g_objNewTab.Screen.CurrentColumn - 1)
                    strPrompt = Trim(strPrompt)
                    
                    '###################################
                    
                    row = nRow

                    '###################################


                    ' Send each command one-by-one to the remote system:
                    For Each strCommand In vCommands
                            If strCommand = "" Then Exit For
                            
                            ' Send the command text to the remote
                            g_objNewTab.Screen.Send strCommand & vbcr
                            
                            ' Wait for the command to be echo'd back to us.
                            g_objNewTab.Screen.WaitForString strCommand
                            
                            ' Since we don't know if we're connecting to a cisco switch or a
                            ' linux box or whatever, let's look for either a Carriage Return
                            ' (CR) or a Line Feed (LF) character in any order.
                            vWaitFors = Array(vbcr, vblf)
                            bFoundEOLMarker = False
                            
                            Do
                                ' Call WaitForStrings, passing in the array of possible
                                ' matches.
                                    g_objNewTab.Screen.WaitForStrings vWaitFors, 1
                                    
                                ' Determine what to do based on what was found)
                                    Select Case g_objNewTab.Screen.MatchIndex
                                        Case 0 ' Timed out
                                            Exit Do
                                            
                                        Case 1,2 ' found either CR or LF
                                        ' Check to see if we've already seen the other
                                        ' EOL Marker
                                        If bFoundEOLMarker Then Exit Do
                                            
                                        ' If this is the first time we've been through
                                        ' here, indicate as much, and then loop back up
                                        ' to the  top and try to find the other EOL
                                        ' marker.
                                            bFoundEOLMarker = True
                                    End Select
                            Loop
                            
                            ' Now that we know the command has been sent to the remote
                            ' system, we'll begin the process of capturing the output of
                            ' the command.
                            
                            Dim strResult
                            ' Use the ReadString() method to get the text displayed
                            ' while the command was runnning.  Note that the ReadString
                            ' usage shown below is not documented properly in SecureCRT
                            ' help files included in SecureCRT versions prior to 6.0
                            ' Official.  Note also that the ReadString() method captures
                            ' escape sequences sent from the remote machine as well as
                            ' displayed text.  As mentioned earlier in comments above,
                            ' if you want to suppress escape sequences from being
                            ' captured, set the Screen.IgnoreEscape property = True.
                            strResult = g_objNewTab.Screen.ReadString(strPrompt)
                            
                            Dim objFile, strLogFile
                        
                            ' Set the log file name based on the remote host's IP
                            ' address and the command we're currently running.  We also
                            ' add a  date/timestamp to help make each filename unique
                            ' over time.

                            If g_bLogToIndividualFilesPerHost Then
                                strLogFile = Replace( _
                                g_strLogFileTemplate, _
                                "IPADDRESS", _
                                g_objNewTab.Session.RemoteAddress)
                            Else
                                strLogFile = Replace( _
                                g_strLogFileTemplate, _
                                "IPADDRESS", _
                                "ALLHOSTS")
                            End If
                            
                            If Instr(strLogFile, "COMMAND") Then
                                If g_bLogToIndividualFilesPerCommand Then
                                    ' Replace any illegal characters that might have been
                                    ' introduced by the command we're running (e.g. if the
                                    ' command had a path or a pipe in it)
                                    strCleanCmd = Replace(strCommand, "/", "[SLASH]")
                                    strCleanCmd = Replace(strCleanCmd, "\", "[BKSLASH]")
                                    strCleanCmd = Replace(strCleanCmd, ":", "[COLON]")
                                    strCleanCmd = Replace(strCleanCmd, "*", "[STAR]")
                                    strCleanCmd = Replace(strCleanCmd, "?", "[QUESTION]")
                                    strCleanCmd = Replace(strCleanCmd, """", "[QUOTE]")
                                    strCleanCmd = Replace(strCleanCmd, "<", "[LT]")
                                    strCleanCmd = Replace(strCleanCmd, ">", "[GT]")
                                    strCleanCmd = Replace(strCleanCmd, "|", "[PIPE]")
                                    strLogFile = Replace(strLogFile, "COMMAND", strCleanCmd)
                                Else
                                    strLogFile = Replace(strLogFile, "COMMAND", "ALLCMDS")
                                End If
                            End If
                            
                            ' If the log folder doesn't exist, try to create it
                            strParentFolder = g_fso.GetParentFolderName(strLogFile)
                            strExcelParentFolder = g_fso.GetParentFolderName(strLogFile)

                            If Not g_fso.FolderExists(strParentFolder) Then
                                If Not CreateFolderTree(strParentFolder) Then
                                    strErrors = CaptureError(strErrors, _
                                    "Failed to create folder '" & _
                                    strParentFolder & "' for logging " & _
                                    "results of command '" & strCommand & _
                                    "'.")
                                End If
                            End If

                            On Error Resume Next
                                Set objFile = g_fso.OpenTextFile(strLogFile, ForAppending, True)
                                If Err.Number <> 0 Then
                                    strError = "Failed to open '" & _
                                    strLogFile & "' for writing results of " & _
                                    "'" & strCommand & "' command: " & Err.Description
                                    strErrors = CaptureError(strErrors, strError)
                                Else
                                
                                    ' If you do not want the command logged along with the results,
                                    ' command out the following script statement:
                                                
                                    'THE FOLLOWING IS THE ORIGINAL INFO:
                                    'objFile.WriteLine String(80, "=") & vbcrlf & _
                                    '   "Results of command """ & strCommand & _
                                    '  """ sent to host """ & g_strHost & """:" & vbcrlf & _
                                    ' String(80, "-")
                                    ' Write out the results of the command
                                    'objFile.WriteLine strResult
                                                
                                    'THE FOLLOWING WAS CHANGED BY ME:
                                    objFile.WriteLine String(10, "-") & vbcrlf & _
                                    "Command: """& strCommand & "       CSR IP:   """ & g_strHost & "       Results: " '& g_strResult & _
                                    objFile.WriteLine strResult
                                    
                                    ' Write out the results of the command
                                    
                                    If Err.Number <> 0 Then
                                        strError = "Failed to write results to " & _
                                        "log file (" & strLogFile & "): " & _
                                        Err.Description
                                        strErrors = CaptureError(strErrors, strError)
                                    End If
                                    ' Close the log file
                                    objFile.Close
                                    
                                End If
                            On Error Goto 0
                    Next

                    'SaveExcelSheet() '
                    
                    ' Now disconnect from the current machine before connecting to
                    ' the next machine
                    Do
                        g_objNewTab.Session.Disconnect
                        crt.Sleep 100
                    Loop While g_objNewTab.Session.Connected
                    
                    strSuccesses = strSuccesses & vbcrlf & g_strHost
                End If ' WaitForScreenContentsToStopChanging()
            End If ' Connect()
    End If ' bContinue (commands file found)

    Next
    
    strMsg = "Commands were sent to the following hosts: " & vbcrlf & strSuccesses

    If strErrors <> "" Then
        strMsg = strMsg & vbcrlf & vbcrlf & _
        "Errors encountered include:" & _
        vbcrlf & strErrors
    End If
    
    ' DisplayMessage strMsg
    ' If strErrors <> "" Then
    '     If crt.Dialog.MessageBox("Would you like these errors copied to the clipboard?", "Copy to Clipboard?", vbyesNo) = vbYes Then
    '         crt.Clipboard.Text = strErrors
    '     End If
    ' End If

    ' ' Prompt the user if changes to the workbook should be saved:
    '     ' DisplayMessage(strSuccesses)

    ' If strSuccesses = True Then
    '     wb.Close False
    ' Else
    '     wb.SaveAs("C:\Users\rt948076\OneDrive - Sprint\3-Performance\results\2.5 Audit\chart.xlsx")
    '     wb.Close True
    ' End If
    
    ' 'wb.Close
    ' app.Quit

    ' Set ws = nothing
    ' Set wb = nothing
    ' Set app = nothing

End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function ReadDataFromFile(strFile, _
    strComment, _
    ByRef vLines, _
    ByRef nLineCount, _
    ByRef nCommentLines)
    ' Returns True if the file was found and could be opened for reading.
    '        strFile: IN  parameter specifying full path to data file.
    '     strComment: IN  parameter specifying string that preceded
    '                    by 0 or more space characters will indicate
    '                    that the line should be ignored.
    '        vLines: OUT parameter (destructive) containing array
    '                    of lines read in from file.
    '    nLineCount: OUT parameter (destructive) indicating number
    '                    of lines read in from file.
    ' nCommentLines: OUT parameter (destructive) indicating number
    '                    of comment/blank lines found
    '
    '
    ' Check to see if the file exists... if not, bail early.
    If Not g_fso.FileExists(strFile) Then
    g_strError = "ReadDataFromFromFile: File not found: " & strFile
    Exit Function
    End If

    On Error Resume Next
    ' Open a TextStream Object to the file...
    Set objTextStream = g_fso.OpenTextFile(strFile, 1, False, 0)
    nError = Err.Number
    strErr = Err.Description
    On Error Goto 0

    If nError <> 0 Then
    g_strError = "ReadDataFromFile: Error opening file for reading (" & _
    nError & "): " & strErr & vbcrlf & vbtab & strFile
    Exit Function
    End If

    Set objFFile = g_fso.GetFile(strFile)
    If objFFile.Size < 1 Then
    g_strError = "ReadDataFromFile: File is empty: " & strFile
    ReadDataFromFile = False
    Exit Function
    End If

    ' Attempt to determine file encoding
    b1 = objTextStream.Read(1)
    objTextStream.Close

    ' Re-open the file using the proper encoding
    nVal = ASC(b1)
    Select Case nVal
    Case 255 ' Unicode
    Set objTextStream = g_fso.OpenTextFile(strFile, 1, False, -1)

    Case 239
    g_strError = "ReadDataFromFile: 'UTF-8' File type is not supported." & _
    "File must be saved in Unicode or ANSI format." & vbcrlf & _
    vbtab & strFile
    Exit Function

    Case 254 ' Unicode BE
    g_strError = "ReadDataFromFile: 'Unicode BE' File type is not supported." & _
    "File must be saved in Unicode or ANSI format." & vbcrlf & _
    vbtab & strFile
    Exit Function

    Case Else ' Regular ASCII/ANSI
    Set objTextStream = g_fso.OpenTextFile(strFile, 1, False, 0)
    End Select

    ' Start of with a reasonable size for the array:
    ReDim vLines(5)


    ' Used for detecting comment lines, a regular expression object
    Set re = New RegExp
    re.Pattern = "(^[ \t]*(?:" & strComment & ")+.*$)|(^[ \t]+$)|(^$)"
    re.Multiline = False
    re.IgnoreCase = False

    ' Now read in each line of the file and add an element to
    ' the array for each line that isn't just spaces...
    nLineCount = 0
    nCommentLines = 0

    Do While Not objTextStream.AtEndOfStream
    strLine = ""

    ' Find out if we need to make our array bigger yet
    ' to accommodate all the lines in the file.  For large
    ' files, this can be very memory-intensive.
    If UBound(vLines) >= nLineCount Then
    ReDim Preserve vLines(nLineCount + 5)
    End If

    strLine = Trim(objTextStream.ReadLine)

    ' Look for comment lines that match the pattern
    '   [whitespace][strComment]
    If re.Test(strLine) Then
    ' Line matches our comment pattern... ignore it
    nCommentLines = nCommentLines + 1
    Else
    vLines(nLineCount) = strLine
    nLineCount = nLineCount + 1
    End If
    Loop

    objTextStream.Close

    If nLineCount < 1 Then
    g_strError = "ReadDataFromFile: No valid lines found in file: " & strFile
    ReadDataFromFile = False
    Exit Function
    End If

    ReadDataFromFile = True
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function Connect(strConnectInfo)
    ' Connect in a new tab to the host specified
    bWaitForAuthToCompleteBeforeReturning = False
    bLetCallerDetectAndHandleConnectionErrors = True
    Set g_objNewTab = crt.Session.ConnectInTab( _
    strConnectInfo, _
    bWaitForAuthToCompleteBeforeReturning, _
    bLetCallerDetectAndHandleConnectionErrors)

    If g_objNewTab.Session.Connected <> True Then
    If crt.GetLastErrorMessage = "" Then
    g_strError = "Unknown error"
    Else
    g_strError = crt.GetLastErrorMessage
    End If

    ' You're not allowed to close the script tab (the tab in which the
    ' script was launched oringinally), so only try if the new tab really
    ' was a new tab -- not just reusing a disconnected tab.
    If g_objNewTab.Index <> crt.GetScriptTab().Index Then g_objNewTab.Close

    Exit Function
    End If

    ' Make sure the new tab is "Synchronous" so we can properly wait/send/etc.
    g_objNewTab.Screen.Synchronous = True

    ' Handle authentication in the new tab using the new tab's object reference
    ' instead of 'crt'
    nAuthTimeout = 10  ' seconds

    ' Modify the "$", the "]#", and/or the "->" in the array below to reflect
    ' the variety of legitimate shell prompts you would expect to see when
    ' authentication is successful to one of your remote machines.
    vPossibleShellPrompts = Array(_
    "ogin:", _
    "name:", _
    "sword:", _
    "Login incorrect", _
    "authentication failed.", _
    "$", _
    "#", _
    ">")

    bSendUsername = True
    bSendPassword = True
    If SessionExists(g_strHost) Then
    Set objConfig = crt.OpenSessionConfiguration(g_strHost)
    If objConfig.GetOption("Use Login Script") = True Then
    bSendUsername = False
    bSendPassword = False
    End If

    If objConfig.GetOption("Session Password Saved") = True Then
    bSendPassword = False
    End If

    If objConfig.GetOption("Username") <> "" Then
    bSendUsername = False
    End If
    End If

    Do
    On Error Resume Next
    g_objNewTab.Screen.WaitForStrings vPossibleShellPrompts, nAuthTimeout
    nError = Err.Number
    strErr = Err.Description
    On Error Goto 0
    ' Most likely if there was a problem here, it would have been caused by
    ' an unexpected disconnect occurring while WaitForStrings() was running
    ' as called above. If error, set the global description variable and
    ' then exit the function.
    If nError <> 0 Then
    g_strError = Err.Description

    ' Ensure that the session is disconnected before we exit this
    ' function. If there are subsequent hosts to loop through, we don't
    ' want a connected tab interfering with the next host's connection
    ' attempts
    Do
    g_objNewTab.Session.Disconnect
    crt.Sleep 100
    Loop While g_objNewTab.Session.Connected

    Exit Function
    End If

    ' This Select..Case statement represents somewhat of a "state machine"
    ' in which the value of Screen.MatchIndex represents the index of the
    ' array of strings we told WaitForStrings() to look for.  Based on this
    ' index, we know what action needs to be performed next.
    Select Case g_objNewTab.Screen.MatchIndex
    Case 0
    g_strError = "Authentication timed out!" & vbcrlf & _
    "(Or you forgot to add a case for a successful shell " & _
    "prompt in the vPossibleShellPrompts array)"
    ' Disconnect from the host so that we can reuse the disconnected
    ' tab for the next connection in the loop
    Do
    g_objNewTab.Session.Disconnect
    crt.Sleep 100
    Loop While g_objNewTab.Session.Connected

    Exit Function

    Case 1' "ogin:"
    If bSendUsername Then
    ' Send the username... but ONLY if the cursor position
    ' makes sense (If "Last login:" shows up on the screen,
    ' ignore it, so that we don't send the username as a
    ' command, potentially)
    If crt.Screen.CurrentColumn <= Len("Login: ") Then
        g_objNewTab.Screen.Send g_strUsername & vbcr
    End If
    End If


    Case 2' "name:"
    If bSendUsername Then
    ' Send the username... but ONLY if the cursor position
    ' makes sense (If "Username:" shows up anywhere else
    ' on the screen but the left-most word on the line,
    ' ignore this "prompt")
    If crt.Screen.CurrentColumn <= Len("Username: ") Then
        g_objNewTab.Screen.Send g_strUsername & vbcr
    End If
    End If

    Case 3 ' "sword:"
    If bSendPassword Then
    ' Send the password... but ONLY if the cursor position
    ' makes sense (If "Password:" shows up on the screen,
    ' anywhere but as the left-most word on the line, ignore
    ' it, so that we don't send the password command)
    If crt.Screen.CurrentColumn <= Len("Password: ") Then
        g_objNewTab.Screen.Send g_strPassword & vbcr
    End If
    End If

    Case 4,5 ' "Login incorrect", "authentication failed."
    g_strError = _
    "Password authentication to '" & g_strHost & "' as user '" & _
    g_strUsername & "' failed." & vbcrlf & vbcrlf & _
    "Please specify the correct password for user " & _
    "'" & g_strUsername & "'"

    ' Disconnect from the host so that we can reuse the disconnected
    ' tab for the next connection in the loop
    Do
    g_objNewTab.Session.Disconnect
    crt.Sleep 100
    Loop While g_objNewTab.Session.Connected

    Exit Function

    Case 6,7,8 ' "$", "#", or ">" <-- Shell prompt means auth success
    g_objNewTab.Session.SetStatusText _
    "Connected to " & g_strHost & " as " & g_strUsername
    Exit Do

    Case Else
    g_strError = _
    "Ooops! Looks like you forgot to add code " & _
    "to handle this index: " & g_objNewTab.Screen.MatchIndex & _
    vbcrlf & _
    vbcrlf & _
    "Modify your script code's ""Select Case"" block " & _
    "to have 'Case' statements for all of the strings you " & _
    "are passing to the ""WaitForStrings"" method."

    Do
    g_objNewTab.Session.Disconnect
    crt.Sleep 100
    Loop While g_objNewTab.Session.Connected

    Exit Function
    End Select
    Loop

    ' If the code gets here, then we must have been successful connecting and
    ' authenticating to the remote machine; Assign True as the return value
    ' for the Connect() function.
    Connect = True
End Function
' -----------------------------------------------------------------------------
Function CaptureError(strErrors, strError)
    If strErrors = "" Then
    strErrors = strError
    Else
    If Right(strErrors, 2) <> vbcrlf Then
    strErrors = strErrors & vbcrlf & strError
    Else
    strErrors = strErrors & strError
    End If
    End If

    strLogFile = g_strLogFileTemplate
    If Instr(strLogFile, "IPADDRESS") > 0 Then
    strLogFile = Replace(strLogFile, "IPADDRESS", "All(Errors)")
    Else
    strLogFile = g_fso.GetParentFolderName(strLogFile) & "\" & _
    "All(Errors)" + g_fso.GetFileName(strLogFile)
    End If
    strLogFile = Replace(strLogFile, "COMMAND", "")


    ' Try to create the log file's parent folder if it doesn't
    ' already exist
    strParentFolder = g_fso.GetParentFolderName(strLogFile)
    If Not g_fso.FolderExists(strParentFolder) Then
    CreateFolderTree(strParentFolder)
    End If

    On Error Resume Next
    Set objFile = g_fso.OpenTextFile(strLogFile, ForAppending, True)
    If Err.Number <> 0 Then
    If Instr(strError, "Failed to open") = 0 And _
    Instr(strError, "Failed to write") = 0 And _
    Instr(strError, "Failed to create") = 0 Then
    strAdditionalErrors = "Error opening log file (" & strLogFile & _
    ") for writing: " & Err.Description
    End If
    Else
    Err.Clear
    objFile.WriteLine String(80, "=") & vbcrlf & strError & vbcrlf
    If Err.Number <> 0 Then
    strAdditionalErrors = "Error writing to log file (" & strLogFile & _
    "): " & Err.Description
    End If
    End If
    On Error Goto 0

    If strAdditionalErrors <> "" Then
    strErrors = strErrors & vbcrlf & vbtab & strAdditionalErrors
    End If

    ' Return the full list of errors (now updated to contain the most
    ' recent one:
    CaptureError = strErrors
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function CreateFolderTree(szFolderPath)
    On Error Resume Next

    if g_fso.FolderExists(szFolderPath) then
    'WScript.Echo "   CreateFolderTree:: Folder exists: " & szFolderPath
    CreateFolderTree = True
    exit function
    end if

    Do
    Err.Clear
    g_fso.CreateFolder szFolderPath
    If Err.Number <> 0 then
    if Err.Number <> 70 Then
    Err.Clear
    CreateFolderTree(g_fso.GetParentFolderName(szFolderPath))
    Else
    CreateFolderTree = False
    Exit Function
    End If
    Else
    Exit do
    end if
    Loop

    CreateFolderTree = True

    On Error Goto 0
End Function

' -----------------------------------------------------------------------------
Function WaitForScreenContentsToStopChanging(nMsDataReceiveWindow)
    ' This function relies on new data received being different from the
    ' data that was already received.  It won't work if, as one example, you
    ' have a screenful of 'A's and more 'A's arrive (because one screen
    ' "capture" will look exactly like the previous screen "capture").

    ' Store Synch flag for later restoration
    bOrig = g_objNewTab.Screen.Synchronous
    ' Turn Synch off since speed is of the essence; we'll turn it back on (if
    ' it was already on) at the end of this function
    g_objNewTab.Screen.Synchronous = False

    ' Be "safe" about trying to access Screen.Get().  If for any reason we
    ' get disconnected, we don't want the script to error out on the problem
    ' so we'll just return false and handle writing something to our log
    ' file for this host.
    On Error Resume Next
    strLastScreen = g_objNewTab.Screen.Get(1,1,g_objNewTab.Screen.Rows,g_objNewTab.Screen.Columns)
    nError = Err.Number
    strErr = Err.Description
    On Error Goto 0
    If nError <> 0 Then
    g_strError = strErr
    Exit Function
    End If
    Do
    crt.Sleep nMsDataReceiveWindow

    ' Be "safe" about trying to access Screen.Get().  If for any reason we
    ' get disconnected, we don't want the script to error out on the problem
    ' so we'll just return false and handle writing something to our log
    ' file for this host.
    On Error Resume Next
    strNewScreen = g_objNewTab.Screen.Get(1,1,g_objNewTab.Screen.Rows, g_objNewTab.Screen.Columns)
    nError = Err.Number
    strErr = Err.Description
    On Error Goto 0
    If nError <> 0 Then
    g_strError = strErr
    Exit Function
    End If

    If strNewScreen = strLastScreen Then Exit Do

    strLastScreen = strNewScreen
    Loop

    WaitForScreenContentsToStopChanging = True
    ' Restore the Synch setting
    g_objNewTab.Screen.Synchronous = bOrig
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function NN(nNumber)
    ' Normalizes a single digit number into a double-digit number with a leading 0
    If Len(nNumber) < 2 Then nNumber = "0" & nNumber
    NN = nNumber
End Function

'-------------------------------------------------------------------------------
Function PromptYesNo(strText)
    PromptYesNo = crt.Dialog.MessageBox(strText, "SecureCRT", vbYesNo)
End Function

'-------------------------------------------------------------------------------
Function DisplayMessage(strText)
    crt.Dialog.MessageBox strText
End Function

'-------------------------------------------------------------------------------
Function SessionExists(strSessionPath)
    ' Returns True if a session specified as value for strSessionPath already
    ' exists within the SecureCRT configuration.
    ' Returns False otherwise.
    On Error Resume Next
    Set objTosserConfig = crt.OpenSessionConfiguration(strSessionPath)
    nError = Err.Number
    strErr = Err.Description
    On Error Goto 0
    ' We only used this to detect an error indicating non-existance of session.
    ' Let's get rid of the reference now since we won't be using it:
    Set objTosserConfig = Nothing
    ' If there wasn't any error opening the session, it's a 100% indication
    ' that the session named in strSessionPath already exists
    If nError = 0 Then
    SessionExists = True
    Else
    SessionExists = False
    End If
End Function

'#############################################################################################################

    ' '-------------------------------------------------------------------------------
    ' Function SaveExcelSheet()
    '     ' Attempt to open the workbook by calling OpenWorkbook function defined
    '     ' within this script:
    '         ' If Not OpenWorkbook(strWorkBookPath, wb, strErrors) Then
    '         '     crt.Dialog.MessageBox strErrors
    '         '     Exit Function
    '         ' End If
            
    '         ' ' Get a reference to an existing sheet in the workbook
    '         ' Dim objWs
    '         ' strExistingSheetName = "Sheet1"
    '         ' If SheetExists(wb, strExistingSheetName) Then
    '         '     Set objWs = wb.Sheets(strExistingSheetName)
    '         ' End If
            
    '     'On Error Resume Next
    '             'nError = Err.Number
    '             'If nError <> 0 Then
    '             '    ws.Cells(row,1).Value = Mid(g_strLogFileTemplate, 79,20)
    '             '    ws.Cells(row,2).Value = g_strHost 
    '             '    ws.Cells(row,3).Value = "Failed"
    '             '    DisplayMessage("Failed")
    '             'Else
    '             'ws.Cells(row,1).Value = Mid(g_strLogFileTemplate, 79,20)
    '             'ws.Cells(row,2).Value = g_strHost
    '             'ws.Cells(row,3).Value = strResult
    '             '    DisplayMessage( strResult & " TEST " )
    '             'End If
    '     'On Error Goto 0

    '     crt.screen.synchronous = true
        
    '       ' Create an Excel workbook/worksheet
    '       '
    '       Dim app, wb, ws
    '       Set app = CreateObject("Excel.Application")
    '       Set wb = app.Workbooks.Add
    '       Set ws = wb.Worksheets(1)
        
    '       ' Send the initial command to run and wait for the first linefeed
    '       '
    '       'crt.Screen.Send("cat /etc/passwd" & Chr(10) )
    '       'crt.Screen.WaitForString Chr(10)
        
    '       ' Create an array of strings to wait for.
    '       '
    '       '   Dim waitStrs
    '       '   waitStrs = Array( Chr(10), "linux$" )
        
    '       Dim row, screenrow, readline, items
    '       row = 1
        
    '       'Do
    '         'While True
        
    '           ' Wait for the linefeed at the end of each line, or the shell prompt
    '           ' that indicates we're done.
    '           '	
    '           result = crt.Screen.WaitForStrings( vWaitFors )
        
    '           ' We saw the prompt, we're done.
    '           '
    '           If result = 2 Then
    '             'Exit Do
    '           End If
        
    '           ' Fetch current row and read the first 40 characters from the screen
    '           ' on that row. Note, since we read a linefeed character subtract 1 
    '           ' from the return value of CurrentRow to read the actual line.
    '           '
    '             screenrow = crt.screen.CurrentRow -1
    '             readline = crt.Screen.Get(screenrow, 1, screenrow, 75 )

    '             ' Split the line ( ":" delimited) and put some fields into Excel
    '             '
    '             'DisplayMessage(readline)
    '             'items = Split( readline, ":", -1 )
    '             'DisplayMessage(items)
                
    '             if readline <> "" Then
    '                 ws.Cells(row,1).Value = Mid(g_strLogFileTemplate, 79,20)
    '                 ws.Cells(row,2).Value = g_strHost
    '                 ws.Cells(row,3).Value = "Failed"                  
    '             else
    '                 ws.Cells(row,1).Value = Mid(g_strLogFileTemplate, 79,20)
    '                 ws.Cells(row,2).Value = g_strHost
    '                 ws.Cells(row,3).Value = readline               
    '             end if



    '             'ws.Cells(row, 1).Value = items(0)
    '             'ws.Cells(row, 2).Value = items(1)
    '             row = row + 1
    '         'Wend
    '       'Loop
            
    '       wb.SaveAs("C:\Users\rt948076\OneDrive - Sprint\3-Performance\results\2.5 Audit\chart.xlsx")
    '       wb.Close
    '       app.Quit
        
    '       Set ws = nothing
    '       Set wb = nothing
    '       Set app = nothing
        
    '       crt.screen.synchronous = false
    ' End Function

    ' '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    ' Function OpenWorkbook(strWkBkPath, ByRef wb, ByRef strErrors)
    '     ' Returns True if workbook was found to exist and was successfully opened.
    '     ' Returns False otherwise.
    '     ' 
    '     ' Requires 'objExcel' global variable to already have been initialized via
    '     ' CreateObject("Excel.Application")
    '     ' 
    '     ' strWkBkPath: [In] String representing the full path to the workbook .xls file
    '     '              to open.
    '     '
    '     '       wb: [Out] Reference to Excel Workbook object.  This will be Nothing
    '     '              if the workbook .xls file specified in the strWkWBkPath couldn't
    '     '              be loaded.
    '     '              
    '     '   strErrors: [Out] Used only if there were errors encountered when attempting
    '     '              to open the Excel workbook .xls file specified in strWkBkPath.

    '   On Error Resume Next
    '   Set wb = objExcel.Workbooks.Open(strWkBkPath)
    '   nError = Err.Number
    '   strErr = Err.Description
    '   On Error Goto 0

    '   If nError <> 0 Then
    '       strErrors = "Error opening workbook file: " & strWkBkPath & _
    '           vbcrlf & vbcrlf & strErr
    '       Set wb = Nothing
    '       Exit Function
    '   End If
    
    '   ' Set the reference to the [Out] param wb:
    '   Set wb = wb
    
    '   ' Set the return value of this function:
    '   OpenWorkbook = True
    ' End Function
    ' '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    ' Function SheetExists(objWorkbook, strSheetname)
    '   ' In this "trick", we ask the workbook to provide us
    '   ' with a reference to a sheet by name.  If the sheet
    '   ' doesn't exist, we simply handle the error and use
    '   ' the error as a way of determining if the named sheet
    '   ' already exists or not.
    
    '   ' Start off assuming the sheet exists:
    '   bExists = True

    '   ' Now, ask the workbook to provide a reference to the
    '   ' named sheet (passed into this function as 'strSheetname'):
    '   On Error Resume Next
    '   set ws = objWorkbook.Sheets(strSheetName)
    '   nError = Err.Number
    '   strErr = Err.Description
    '   On Error Goto 0
    
    '   ' If there was an error, it means the sheet doesn't exist:
    '   If nError <> 0 Then bExists = False
    
    '   ' Set the return value of our function:
    '   SheetExists = bExists
    ' End Function

    ' Function XXX() 
    '     If strErrors = "" Then
    '         strErrors = strError
    '         Else
    '         If Right(strErrors, 2) <> vbcrlf Then
    '             strErrors = strErrors & vbcrlf & strError
    '         Else
    '             strErrors = strErrors & strError
    '         End If
    '     End If

    '     ' Attempt to open the workbook by calling OpenWorkbook function defined
    '     ' within this script:
    '     If Not OpenWorkbook(strWorkBookPath, objWb, strErrors) Then
    '         crt.Dialog.MessageBox strErrors
    '         Exit Function
    '     End If

    '     ' Get a reference to an existing sheet in the workbook
    '     Dim objWs
    '     strExistingSheetName = "Sheet1"
    '     If SheetExists(objWb, strExistingSheetName) Then
    '         Set objWs = objWb.Sheets(strExistingSheetName)
    '     End If
    ' End Function