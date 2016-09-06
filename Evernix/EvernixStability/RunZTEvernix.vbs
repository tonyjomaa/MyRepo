On Error Resume Next 

Set WShell = WScript.CreateObject ("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

logfilepath = "C:\Windows\Ai2Install.log"
ZTArchivePath = "\\ai2s0017\ts01\GROUP\BUILDSTORAGE\Images\ZoomText\Mainline"

'''''''''''''''''''''''''''Get color code ''''''''''''''''''''''''''
Set objcolorTextFile = objFSO.GetFile ("d:\Utils\color.txt")
Set objcolorTextFileStream = objcolorTextFile.OpenAsTextStream(1,0)  ' for Reading
color = objcolorTextFileStream.Readline '("3399FF")
WScript.Sleep (50)
objcolorTextFileStream.close
Set objcolorTextFile = Nothing 
Set objcolorTextFileStream = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

 Set objInstalllogTextFile = objFSO.GetFile (logfilepath)
 Set objInstalllogTextFileStream = objInstalllogTextFile.OpenAsTextStream(2,-1) 'For Writing to delete contents
objInstalllogTextFileStream.close
Set objInstalllogTextFileStream = Nothing
Set objInstalllogTextFile = Nothing

'Set objlogTextFile = objFSO.CreateTextFile("d:\Utils\Results.txt")
Set objlogTextFile = objFSO.GetFile ("d:\Utils\Results.txt")
Set objlogTextFileStream = objlogTextFile.OpenAsTextStream(8,-1) 'For appending

InstallLocationFilePath = "q:\ATS\ATM\atm_resources\installLocationEverest.txt"
Set objInstallLocationTextFile = objFSO.GetFile(InstallLocationFilePath)
Set objInstallLocationTextStream = objInstallLocationTextFile.OpenAsTextStream(1,-1)  'open for reading
buildpath = objInstallLocationTextStream.ReadLine
WScript.Sleep(100)
str = buildpath
number = InStr (str,"10.0.0.")-1
number = Len (str) - number
str = Right (str, number)
str = Left (str,11)
str = Replace (str,"\","")
str = Replace (str,"Pa","")
ZTVer = Replace (str,"(","")

'WScript.Sleep(100)
'ZTPath = wshell.RegRead ("HKEY_LOCAL_MACHINE\SOFTWARE\Ai Squared\ZoomText Xtra\10.1\Settings\Reader\Program Directory")
'Command = ZTPath & "\Zt.exe"

If objfso.FolderExists("C:\Program Files (x86)\ZoomText 10.1") then
   str="C:\Program Files (x86)\ZoomText 10.1\Zt.exe"
   Command = "cmd /c ""C:\Program Files (x86)\ZoomText 10.1\Zt.exe"""
Else
  str="C:\Program Files (x86)\ZoomText 11.0\Zt.exe"
  Command = "cmd /c ""C:\Program Files (x86)\ZoomText 11.0\Zt.exe"""
End If

'msgbox command
strFileVersion = ""
strFileVersion = objFSO.GetFileVersion(str)
If strFileVersion <> "" Then 
	objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>The ZoomText version from file is</font></th><th><font>"&strFileVersion&"</font></th></tr>")
	If color = "FFFFFF" Then  
    	color = "99CCFF"
	Else 
		color = "FFFFFF"
	End If 
End If

If (WScript.Arguments.Count=0) Then

	
	WShell.run Command,1,false
	WScript.Sleep(5000)
	
	For I = 1 To 100
	    set ZT =createObject("ZoomText.Application")
	    if isobject(ZT) then exit for
		'If wshell.AppActivate("ZoomText 10.1",False) Then Exit For
		WScript.Sleep(1000)
	Next
	
	If I <= 100 Then
		ZT_TimeToReady = I + 5
		objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>The time in seconds it took ZoomText 10 UI to appear on the Desktop</font></th><th><font>"&ZT_TimeToReady&"</font></th></tr>")
			
	Else
		objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>ZoomText UI took to more than 1 minute 41 seconds to appear on the Desktop</font></th><th><font color=Red>FAILED</font></th></tr>")
	End If
	
Else

    
	WShell.Exec Command
	WScript.Sleep(500)
	
	For I = 1 To 100
	
		If wshell.AppActivate("ZoomText 10.1",False) Or wshell.AppActivate("ZoomText 11.0",False) Or wshell.AppActivate("ZoomText 11",False) Then Exit For
		WScript.Sleep(1000)
	Next
	
	If I <= 100 Then
		ZT_TimeToReady = I-1
		objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>The time in seconds it took ZoomText UI to appear on the Desktop</font></th><th><font>"&ZT_TimeToReady&"</font></th></tr>")
			
	Else
		objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>ZoomText UI took to more than 1 minute 41 seconds to appear on the Desktop</font></th><th><font color=Red>FAILED</font></th></tr>")
	End If 		 		

End If

'''''''''''''''''''''''''''Flip color code ''''''''''''''''''''''''''
	Set objcolorTextFile = objFSO.GetFile ("d:\Utils\color.txt")
	Set objcolorTextFileStream = objcolorTextFile.OpenAsTextStream(2,0)  ' for writing
	'color = objcolorTextFileStream.Readline '("3399FF")

	If color = "FFFFFF" Then  
    	color = "99CCFF"
	Else 
		color = "FFFFFF"
	End If 
	objcolorTextFileStream.Writeline (color)
	WScript.Sleep (50)
	objcolorTextFileStream.close
	Set objcolorTextFile = Nothing 
	Set objcolorTextFileStream = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
ejectProcess()
Set objInstallLocationTextFile = Nothing
Set objInstallLocationTextStream = Nothing
Set objlogTextFile = Nothing 
Set objlogTextFileStream = Nothing 
Set WShell = Nothing
Set objFSO = Nothing
WScript.Quit

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

   Set objFSO = Nothing
End Function

Function strComputerName()

	Set objNetwork  = CreateObject("WScript.Network")
	strComputerName = objNetwork.ComputerName
	
	set objNetwork = Nothing
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''