'#$language = "VBScript"
'#$interface = "1.0" 
'Name: Post SIP TRKGRP.simple.vbs
'This is a simplified sample script to demonstrate SecureCRT scriting without the use of supporting classes.

Option Explicit

Dim s_objScriptTab : Set s_objScriptTab = crt.GetScriptTab

crt.Screen.Synchronous = True

Sub Main()
    Dim strClli
    
    s_objScriptTab.Screen.Send "leave all;" & Chr(13)
    If (Not s_objScriptTab.Screen.WaitForString(">", 5)) Then
    	TimeoutMessage
        Exit Sub
    End If 
    strClli = crt.Dialog.Prompt("Enter Trunk Group:", "Post Trunk Group Command", "")
    If (Len(strClli) > 0) Then 
        s_objScriptTab.Screen.Send "send sink;leave all;mapci;mtc;trks;dptrks;SEND PREVIOUS;post g "
        s_objScriptTab.Screen.Send strClli
        s_objScriptTab.Screen.Send vbCr 
        If (Not s_objScriptTab.Screen.WaitForString(">", 5)) Then
   			TimeoutMessage
            Exit Sub
        End If 
    End If 
End Sub

Sub TimeoutMessage()
	crt.Dialog.Messagebox "Did not find the '>' prompt on the screen within 5 seconds.  Aborting script now.", "TIMEOUT!", vbOKOnly & vbCritical
End sub