On Error Resume Next 

Set WShell = CreateObject ("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

'''''''''''''''''''''''''''Get color code ''''''''''''''''''''''''''
Set objcolorTextFile = objFSO.GetFile ("d:\Utils\color.txt")
Set objcolorTextFileStream = objcolorTextFile.OpenAsTextStream(1,0)  ' for Reading
color = objcolorTextFileStream.Readline '("3399FF")
WScript.Sleep (50)
objcolorTextFileStream.close
Set objcolorTextFile = Nothing 
Set objcolorTextFileStream = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'
Set objlogTextFile = objFSO.GetFile ("d:\Utils\Results.txt")
Set objlogTextFileStream = objlogTextFile.OpenAsTextStream(8,-1)  'for appending

Set objpathTextFile = objFSO.GetFile ("d:\Utils\path.txt")
Set objpathTextFileStream = objpathTextFile.OpenAsTextStream(1,0) ' for reading

logfilepath=objpathTextFileStream.ReadLine
WScript.Sleep(500)
objpathTextFileStream.Close

ZTVersion = objfso.GetFileVersion("C:\Program Files (x86)\ZoomText 11.0\AiSquared.Magnification.ZoomText.exe")

If ZTVersion = "" Then objfso.GetFileVersion("C:\Program Files (x86)\ZoomText 10.1\AiSquared.Magnification.ZoomText.exe")
     
ZTEVVersion = objfso.GetFileVersion("C:\Program Files (x86)\ZoomText 11.0\Zt.exe")
If ZTEVVersion = "" Then objfso.GetFileVersion("C:\Program Files (x86)\ZoomText 10.1\Zt.exe")
    
'msgbox ztversion
Set WshSysEnv = WShell.Environment("PROCESS")
userprofile = WshSysEnv("USERPROFILE")
'msgbox userprofile & "\AppData\Roaming\Ai Squared\ZoomText 11.0"
If objfso.FolderExists(userprofile & "\AppData\Roaming\Ai Squared\ZoomText 11.0") Then 
	ZTlogfilepath = userprofile & "\AppData\Roaming\Ai Squared\ZoomText 11.0" 

Else

	ZTlogfilepath = userprofile & "\AppData\Roaming\Ai Squared\ZoomText 10.1" 
End If

'ZTEVlogfilepath = userprofile & "\AppData\Roaming\Ai Squared\ZoomText 10.1\" & ZTEVVersion 
'wscript.quit
'msgbox ZTlogfilepath&vbcrlf& logfilepath & "\" 
If objFSO.FolderExists(ZTlogfilepath) Then objFSO.CopyFolder ZTlogfilepath, logfilepath & "\", True
'If objFSO.FolderExists(ZTEVlogfilepath) Then objFSO.CopyFolder ZTEVlogfilepath, logfilepath & "\", True

SomeQAFolder=logfilepath & "\"
cmd=""
WScript.Sleep(5000)
If objfso.FileExists("C:\Program Files (x86)\ZoomText 10.1\Zxplog.txt") Then 
  cmd="cmd /c xcopy /Y ""C:\Program Files (x86)\ZoomText 10.1\Zxplog*"" " & SomeQAFolder
ElseIf objfso.FileExists("C:\Program Files (x86)\ZoomText 11.0\Zxplog.txt") Then
 cmd="cmd /c xcopy /Y ""C:\Program Files (x86)\ZoomText 11.0\Zxplog*"" " & SomeQAFolder
End If
If cmd <> "" Then WShell.run cmd,4,true

If objFSo.FileExists (logfilepath & "\Zxplog.txt") Then 
		objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>Copying log files to " & logfilepath & "</font></th><th><font>PASSED</font></th></tr>")
		
Else 
		objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>Copying log files to " & logfilepath & "</font></th><th><font color=GoldenRod>FALSE</font></th></tr>")
		
End If 
WScript.Sleep(5000)	

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

Set WshSysEnv = Nothing
Set objInstallLocationTextFile = Nothing
Set objInstallLocationTextStream = Nothing
Set objlogTextFile = Nothing 
Set objlogTextFileStream = Nothing 
Set objpathTextFileStream = Nothing
Set objpathTextFile = Nothing
Set WShell = Nothing
Set objFSO = Nothing
ejectProcess()

'WScript.Quit()

Function ejectProcess()
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

Function strComputerName()

	Set objNetwork  = CreateObject("WScript.Network")
	strComputerName = objNetwork.ComputerName
	
	set objNetwork = Nothing
End Function