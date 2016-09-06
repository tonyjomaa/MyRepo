On Error Resume Next 

fixedAllLogging = "q:\ATS\AutoTasks\All\AiSquared.Logging.Logger.dll.config"

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set wshShell = CreateObject("WScript.Shell")
'''''''''''''''''''''''''''Get color code ''''''''''''''''''''''''''
Set objcolorTextFile = objFSO.GetFile ("d:\Utils\color.txt")
Set objcolorTextFileStream = objcolorTextFile.OpenAsTextStream(1,0)  ' for Reading
color = objcolorTextFileStream.Readline '("3399FF")
WScript.Sleep (50)
objcolorTextFileStream.close
Set objcolorTextFile = Nothing 
Set objcolorTextFileStream = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

ZTVersion = WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\" &_
     "Installer\UserData\S-1-5-18\Products\039E7F954EF3FC948A929C8C5110AA8C\InstallProperties\DisplayVersion")

ZTEVVersion = WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Ai Squared\ZoomText Xtra\10.1\Settings\Reader\Software Version")
      
Set WshSysEnv = wshShell.Environment("PROCESS")
userprofile = WshSysEnv("USERPROFILE")
'msgbox ztversion
If InStr(ZTVersion,"10.10.0.")>0 Then 
    ZTloggerfile = userprofile & "\AppData\Roaming\Ai Squared\ZoomText 10.1\" & ZTVersion & "\LoggerConfig\AiSquared.Logging.Logger.dll.Config"
	ZTEVloggerfile = userprofile & "\AppData\Roaming\Ai Squared\ZoomText 10.1\" & ZTEVVersion & "\LoggerConfig\AiSquared.Logging.Logger.dll.Config"
Else
	ZTloggerfile = userprofile & "\AppData\Roaming\Ai Squared\ZoomText 11\" & ZTVersion & "\LoggerConfig\AiSquared.Logging.Logger.dll.Config"
End If 

Const ForReading = 1
Const ForWriting = 2


'If objFSO.FileExists(fixedAllLogging) And ZTVersion <> "" Then
'	ZTloggerfile = Replace (ZTloggerfile,"\LoggerConfig\AiSquared.Logging.Logger.dll.Config","")
'msgbox ztloggerfile
'	objFSO.CreateFolder ZTloggerfile
'	ztloggerfile = ztloggerfile & "\LoggerConfig"
'objFSO.CreateFolder ZTloggerfile
'	objFSO.CopyFile fixedAllLogging,ZTloggerfile & "\",True
'	objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>Setting Verbosity to ALL</font></th><th><font>PASSED</font></th></tr>")
			
	
'Else
If objFSO.FileExists(ZTloggerfile) Then
	Set objFile = objFSO.OpenTextFile(ZTloggerfile, ForReading)
	
	strText = objFile.ReadAll
	WScript.Sleep(2000)
	
	objFile.Close
	strNewText = Replace(strText, "add key=""Level"" value=""None""", "add key=""Level"" value=""All""")
	strNewText = Replace(strNewText, "add key=""Level"" value=""Info""", "add key=""Level"" value=""All""")
	
	strNewText = Replace(strNewText, "add key=""Level"" value=""Warn""", "add key=""Level"" value=""All""")
	strNewText = Replace(strNewText, "add key=""Level"" value=""Info""", "add key=""Level"" value=""All""")
	
	strNewText = Replace(strNewText, "add key=""AllowAllDebugOutput"" value=""false""", "add key=""AllowAllDebugOutput"" value=""true""")
	
	Set objFile = objFSO.OpenTextFile(ZTloggerfile, ForWriting)
	objFile.WriteLine strNewText
	WScript.Sleep(5000)
	objFile.Close
	'command = userprofile & "
	WScript.Sleep(100)
	Set objlogTextFile = objFSO.GetFile ("d:\Utils\Results.txt")  ' Append to email message
	Set objlogTextFileStream = objlogTextFile.OpenAsTextStream(8,-1)  'for appending
	
	
	
	Set objlogTextFile2 = objFSO.GetFile (ZTloggerfile)
	Set objlogTextFileStream2 = objlogTextFile2.OpenAsTextStream(1,0)
	
	AllOfIt = objlogTextFileStream2.ReadAll
	WScript.Sleep(2000)
	'returnvalue = wshshell.Run ("%comspec% /c c:\windows\system32\find.exe ""All"" " & """"& ZTloggerfile & """",8 ,True)
	'If returnvalue <> 0 Then  '' condition Logging All was found - pass
	If InStr(AllOfIt,"add key=""Level"" value=""All""") > 0 Then 
	
	
	
	'If AppFound = "Yes" Then 
			objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>Setting Verbosity to ALL in: "&ZTloggerfile&"</font></th><th><font>TRUE</font></th></tr>")
			
	Else 
			objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>Setting Verbosity to ALL in: "&ZTloggerfile&"</font></th><th><font color=GoldenRod>FALSE</font></th></tr>")
		
	End If 
	WScript.Sleep(1000)	
End If 
'End if
If color = "3399FF" Then  
    	color = "99CCFF"
	Else 
		color = "3399FF"
End If 

If InStr(AllOfIt,"add key=""AllowAllDebugOutput"" value=""true""") > 0 Then 



'
		objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>Setting 'AllowAllDebugOutput' to true was successfull in: "&ZTloggerfile&"</font></th><th><font>TRUE</font></th></tr>")
		
Else 
		objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>Setting 'AllowAllDebugOutput' to true was not successfull in: "&ZTloggerfile&"</font></th><th><font color=GoldenRod>FALSE</font></th></tr>")
	
End If 


If objFSO.FileExists(ZTEVloggerfile) Then
	Set objFile = objFSO.OpenTextFile(ZTEVloggerfile, ForReading)
	
	strText = objFile.ReadAll
	WScript.Sleep(2000)
	
	objFile.Close
	strNewText = Replace(strText, "add key=""Level"" value=""None""", "add key=""Level"" value=""All""")
	strNewText = Replace(strNewText, "add key=""Level"" value=""Info""", "add key=""Level"" value=""All""")
	
	strNewText = Replace(strNewText, "add key=""Level"" value=""Warn""", "add key=""Level"" value=""All""")
	strNewText = Replace(strNewText, "add key=""Level"" value=""Info""", "add key=""Level"" value=""All""")
	
	strNewText = Replace(strNewText, "add key=""AllowAllDebugOutput"" value=""false""", "add key=""AllowAllDebugOutput"" value=""true""")
	
	Set objFile = objFSO.OpenTextFile(ZTEVloggerfile, ForWriting)
	objFile.WriteLine strNewText
	WScript.Sleep(5000)
	objFile.Close
	'command = userprofile & "
	WScript.Sleep(100)
	Set objlogTextFile = objFSO.GetFile ("d:\Utils\Results.txt")  ' Append to email message
	Set objlogTextFileStream = objlogTextFile.OpenAsTextStream(8,-1)  'for appending
	
	
	
	Set objlogTextFile2 = objFSO.GetFile (ZTEVloggerfile)
	Set objlogTextFileStream2 = objlogTextFile2.OpenAsTextStream(1,0)
	
	AllOfIt = objlogTextFileStream2.ReadAll
	WScript.Sleep(2000)
	'returnvalue = wshshell.Run ("%comspec% /c c:\windows\system32\find.exe ""All"" " & """"& ZTloggerfile & """",8 ,True)
	'If returnvalue <> 0 Then  '' condition Logging All was found - pass
	If InStr(AllOfIt,"add key=""Level"" value=""All""") > 0 Then 
	
	
	
	'If AppFound = "Yes" Then 
			objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>Setting Verbosity to ALL in: "&ZTEVloggerfile&"</font></th><th><font>TRUE</font></th></tr>")
			
	Else 
			objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>Setting Verbosity to ALL in: "&ZTEVloggerfile&"</font></th><th><font color=GoldenRod>FALSE</font></th></tr>")
		
	End If 
	WScript.Sleep(1000)	
End If 
'End if
If color = "3399FF" Then  
    	color = "99CCFF"
	Else 
		color = "3399FF"
End If 

If InStr(AllOfIt,"add key=""AllowAllDebugOutput"" value=""true""") > 0 Then 



'
		objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>Setting 'AllowAllDebugOutput' to true was successfull in: "&ZTEVloggerfile&"</font></th><th><font>TRUE</font></th></tr>")
		
Else 
		objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>Setting 'AllowAllDebugOutput' to true was not successfull in: "&ZTEVloggerfile&"</font></th><th><font color=GoldenRod>FALSE</font></th></tr>")
	
End If 

'''''''''''''''''''''''''''Flip color code ''''''''''''''''''''''''''
	Set objcolorTextFile = objFSO.GetFile ("d:\Utils\color.txt")
	Set objcolorTextFileStream = objcolorTextFile.OpenAsTextStream(2,0)  ' for writing
	'color = objcolorTextFileStream.Readline '("3399FF")

	If color = "3399FF" Then  
    	color = "99CCFF"
	Else 
		color = "3399FF"
	End If 
	objcolorTextFileStream.Writeline (color)
	WScript.Sleep (50)
	objcolorTextFileStream.close
	Set objcolorTextFile = Nothing 
	Set objcolorTextFileStream = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Set objlogTextFileStream = Nothing 
Set objlogTextFile = Nothing 
Set objFSO = Nothing 

ejectProcess()

Set wshShell = Nothing
WScript.Quit


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function ejectProcess()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	strComputerName()	
 strStatusFile = "q:\ATS\ATM\atm_servers\" & strComputerName & "\Init\taskstatus.txt"
 Dim objFSO, objTextFile
   Set objFSO = CreateObject("Scripting.FileSystemObject")
	' OpenTextFile Method needs a Const value
	' ForAppending = 8 ForReading = 1, ForWriting = 2
	Const ForWriting = 2
	Set objTextFile = objFSO.OpenTextFile _
					(strStatusFile, ForWriting, True, -1)
	objTextFile.WriteLine("IDLE")
    objTextFile.Close

   Set objFSO        = Nothing
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function strComputerName()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Set objNetwork  = CreateObject("WScript.Network")
	strComputerName = objNetwork.ComputerName
	
	set objNetwork = Nothing

End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
