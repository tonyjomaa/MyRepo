On Error Resume Next 

Set WShell = WScript.CreateObject ("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
CloseProcess  ' exit ZT processes

ZTArchivePath = "\\ai2s0017\ts01\GROUP\BUILDSTORAGE\Images\Evernix\Mainline"

'''''''''''''''''''''''''''Get color code ''''''''''''''''''''''''''
Set objcolorTextFile = objFSO.GetFile ("d:\Utils\color.txt")
Set objcolorTextFileStream = objcolorTextFile.OpenAsTextStream(1,0)  ' for Reading
color = objcolorTextFileStream.Readline '("3399FF")
WScript.Sleep (50)
objcolorTextFileStream.close
Set objcolorTextFile = Nothing 
Set objcolorTextFileStream = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Set objpathTextFile = objFSO.GetFile ("d:\Utils\path.txt")
Set objpathTextFileStream = objpathTextFile.OpenAsTextStream(1,0) ' for reading
SomeQAFolder = objpathTextFileStream.ReadLine
WScript.Sleep (500)
objpathTextFileStream.Close
Set objpathTextFileStream = Nothing 
Set objpathTextFile = Nothing 

if WShell.AppActivate("ZoomText Activation") then
WShell.SendKeys("%{F4}")
wscript.sleep(1000)
end if

if WShell.AppActivate("ZoomText 10.1 Trial") then
 WShell.SendKeys("%{ENTER}")
wscript.sleep(1000)
end if

if WShell.AppActivate("ZoomText 11.0 Trial") then
 WShell.SendKeys("%{ENTER}")
wscript.sleep(1000)
end if

if WShell.AppActivate("ZoomText 11 Trial") then
 WShell.SendKeys("%{ENTER}")
wscript.sleep(1000)
end if

if WShell.AppActivate("ZoomText Error Reporting") then
 WShell.SendKeys("%{ENTER}")
wscript.sleep(1000)
end if


If objfso.FileExists("C:\Program Files (x86)\ZoomText 10.1\Zxplog.txt") Then 
  cmd="cmd /c xcopy /Y ""C:\Program Files (x86)\ZoomText 10.1\Zxplog*"" " & SomeQAFolder
	'msgbox cmd
	WShell.run cmd
  	'objfso.CopyFile "C:\Program Files (x86)\ZoomText 10.1\Zxplog.txt",SomeQAFolder
ElseIf objfso.FileExists("C:\Program Files\ZoomText 10.1\Zxplog.txt") Then
	'objfso.CopyFile "C:\Program Files\ZoomText 10.1\Zxplog.txt",SomeQAFolder,True
End If
	
 
'Set objlogTextFile = objFSO.CreateTextFile("d:\Utils\Results.txt")
Set objlogTextFile = objFSO.GetFile ("d:\Utils\Results.txt")
Set objlogTextFileStream = objlogTextFile.OpenAsTextStream(8,-1) 'For appending

InstallLocationFilePath = "q:\ATS\ATM\atm_resources\installLocationEvernix.txt"
Set objInstallLocationTextFile = objFSO.GetFile(InstallLocationFilePath)
Set objInstallLocationTextStream = objInstallLocationTextFile.OpenAsTextStream(1,-1)  'open for reading
buildpath = objInstallLocationTextStream.ReadLine
WScript.Sleep(100)
str = buildpath
number = InStr (str,"10.10.0.")-1
number = Len (str) - number
str = Right (str, number)
str = Left (str,11)
str = Replace (str,"\","")
str = Replace (str,"Pa","")
ZTVer = Replace (str,"(","")

'ZTPath = wshell.RegRead ("HKEY_LOCAL_MACHINE\SOFTWARE\Ai Squared\ZoomText Xtra\10.0\Settings\Reader\Program Directory")
If objfso.FolderExists("C:\Program Files (x86)\InstallShield Installation Information\{F7F20305-1476-4421-B909-BB5B90D1F222}") Then
  ' ZT 10.1
  Command = "cmd /k ""C:\Program Files (x86)\InstallShield Installation Information\{F7F20305-1476-4421-B909-BB5B90D1F222}\setup.exe"" -runfromtemp -l0x0009 -ir -niruninst"
Else
  'ZT 11.0
  Command = "cmd /k ""C:\Program Files (x86)\InstallShield Installation Information\{E54BD31E-1F3E-493C-BA71-8203BE18B2DE}\setup.exe"" -runfromtemp -l0x0009 -ir -niruninst"
End If

WShell.run Command,1,false
WScript.Sleep(10000)

objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>ZoomText uninstall script initiated successfully</font></th><th><font>TRUE</font></th></tr>")
		


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
	
'ejectProcess()
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
Function CloseProcess
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    strComputerName
    ProcessArry = Array("ActivationWizard","AiSquared.Loader.Elevated","AiSquared.Loader.Generic","AiSquared.Magnification.ZoomText","ProtectedUI","ProtectedUI64","ZER","Zt")
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputerName & "\root\cimv2")
    Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process")
    For Each OneProcess In ProcessArry
      For Each objProcess in colProcessList     
        If objProcess.Name = oneProcess Then     
              objProcess.Terminate()
              WScript.Sleep(100)        
        End If   
      Next
    Next    
    
    Set colProcessList = Nothing
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
