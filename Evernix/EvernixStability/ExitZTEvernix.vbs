On Error Resume Next 

Set WShell = WScript.CreateObject ("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

logfilepath = "C:\Windows\Ai2Install.log"
ZTArchivePath = "\\ts01\GROUP\BUILDSTORAGE\Images\Evernix\Mainline"

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

InstallLocationFilePath = "q:\ATS\ATM\atm_resources\installLocationEvernix.txt"
Set objInstallLocationTextFile = objFSO.GetFile(InstallLocationFilePath)
Set objInstallLocationTextStream = objInstallLocationTextFile.OpenAsTextStream(1,-1)  'open for reading
buildpath = objInstallLocationTextStream.ReadLine
WScript.Sleep(100)


Set ZTN = CreateObject("ZoomText.Application")
If IsObject(ZTN) Then
	Set ZTC =ZTN.ZTCommand
	ZTC.Command 206,1
	WScript.Sleep(800)
ElseIf wshell.AppActivate("ZoomText 10.1") Then
	WScript.Sleep(800)
	wshell.SendKeys("%{F4}")
ElseIf wshell.AppActivate("ZoomText for Windows 8") Then
	WScript.Sleep(800)
	wshell.SendKeys("%{F4}")
ElseIf 	wshell.AppActivate("ZoomText 10.1") then
 	WScript.Sleep(1800)
	wshell.SendKeys("%{F4}")
ElseIf	wshell.AppActivate("ZoomText 11.0") then
 	WScript.Sleep(1800)
	wshell.SendKeys("%{F4}")
Else
	wshell.AppActivate("ZoomText 11")
 	WScript.Sleep(1800)
	wshell.SendKeys("%{F4}")
End If

For I = 1 To 60
AppFound = True

CheckProcess "Zt.exe",AppFound
If Not AppFound Then Exit For
WScript.Sleep(500)
Next


If Not AppFound Then
	
	objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>ZoomText exited in this many second(s)</font></th><th><font>"&I/2&"</font></th></tr>")
		
Else
	objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>ZoomText did not exit after "&I/2&" second(s)</font></th><th><font>TRUE</font></th></tr>")
End If 		
'If color = "3399FF" Then  
'	    	color = "99CCFF"
'		Else 
'			color = "3399FF"
'	End If
'WScript.Sleep(2000)


	'If Err.Number <> 0 Then 
	'	objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>ZoomText 10 installation</font></th><th><font color=Red>FAILED</font></th></tr>")
		'objlogTextFileStream.writeline("<Action>ZoomText 10 installation Failed</Action><br>" )
		'Exit Do 
	'End If 

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

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function CheckProcess(process,AppFound)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'MsgBox process
    AppFound = False
    strComputerName
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputerName & "\root\cimv2")
    Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process")
    For Each objProcess in colProcessList     
       If objProcess.Name = Process Then     
              AppFound = True        
        End If   
    Next    
    CheckProcess = AppFound
    'MsgBox CheckProcess
    Set colProcessList = Nothing
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
