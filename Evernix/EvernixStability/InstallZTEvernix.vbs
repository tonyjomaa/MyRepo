On Error Resume Next 

Set WShell = WScript.CreateObject ("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

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

 
'Set objlogTextFile = objFSO.CreateTextFile("d:\Utils\Results.txt")
Set objlogTextFile = objFSO.GetFile ("d:\Utils\Results.txt")
Set objlogTextFileStream = objlogTextFile.OpenAsTextStream(8,-1) 'For appending

InstallLocationFilePath = "q:\ATS\ATM\atm_resources\installLocationEvernix.txt"
Set objInstallLocationTextFile = objFSO.GetFile(InstallLocationFilePath)
Set objInstallLocationTextStream = objInstallLocationTextFile.OpenAsTextStream(1,-1)  'open for reading
buildpath = objInstallLocationTextStream.ReadLine
WScript.Sleep(100)
str = buildpath
strarry = Split(buildpath,"\")

ZTVer = strarry(UBound(strarry))

'Command = ZTArchivePath & "\" & ZTVer & "\Installers\Trial\Ftp(Trials)\ZT_basic-en.exe -s -anir -noas"
'Command = buildpath & "\Installers\Trial\Ftp(Trials)\ZT_basic-en.exe -s -ani -noas -novls"
If WScript.Arguments.Count = 0 Then

	Command = "D:\temp\ZT_basic-en.exe /s /ani /soc /noas /nols"
Else

    Command = "D:\temp\10.0.6\ZT_basic-en.exe /s /ani /soc /noas /nols"
    
End If

ObjFSO.CopyFile buildpath & "\Installers\Trial\Ftp(Trials)\ZT_basic-en.exe", "D:\temp\",True
WScript.Sleep(30000)

WShell.Run Command,4,True
If Err.Number <> 0 Then
	objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>Starting ZoomText installation</font></th><th><font color=Red>FALSE</font></th></tr>")
Else
	objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>Starting ZoomText installation</font></th><th><font>TRUE</font></th></tr>")
End If 	
If color = "FFFFFF" Then  
    	color = "99CCFF"
	Else 
		color = "FFFFFF"
	End If 	
I=0
ProcessCheck "ZT_basic-en.exe",I
I=0
ProcessCheck "Setup.exe",I
If I > 300 Then 
	objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>ZoomText installation process finished successfully</font></th><th><font color=Red>FALSE</font></th></tr>")
Else
	objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>ZoomText installation process finished successfully</font></th><th><font>TRUE</font></th></tr>")
End If 

objfso.CopyFile "c:\windows\Ai2Install*.log","d:\utils\",True
WScript.Sleep(1000)



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

If objfso.FolderExists("C:\Program Files (x86)\ZoomText 10.1") then
  objfso.CopyFile "d:\utils\Zten-US.zxc","C:\Program Files (x86)\ZoomText 10.1\Config\Defaults\Zten-US.zxc",True
  objfso.CopyFile "d:\utils\TestSupp.ini","C:\Program Files (x86)\ZoomText 10.1\TestSupp.ini",True
  Command = "C:\Program Files (x86)\ZoomText 10.1\Zt.exe"
Else
  objfso.CopyFile "d:\utils\Zten-US.zxc","C:\Program Files (x86)\ZoomText 11.0\Config\Defaults\Zten-US.zxc",True
  objfso.CopyFile "d:\utils\TestSupp.ini","C:\Program Files (x86)\ZoomText 11.0\TestSupp.ini",True
  Command = "C:\Program Files (x86)\ZoomText 11.0\Zt.exe"
End If

WShell.Run "cmd /c shutdown /r /f /t 0",4, False
	
Set objInstallLocationTextFile = Nothing
Set objInstallLocationTextStream = Nothing
Set objlogTextFile = Nothing 
Set objlogTextFileStream = Nothing 

Set WShell = Nothing
Set objFSO = Nothing
WScript.Quit

Function ProcessCheck(ProcessName,I)

	strComputerName()
	
	Set objWMIService = GetObject("winmgmts:" _
	    & "{impersonationLevel=impersonate}!\\" & strComputerName & "\root\cimv2")
	
	Set colProcessList = objWMIService.ExecQuery _
	    ("Select * from Win32_Process")     
	' wait 10 minutes to install ZT
	Do Until I > 300 Or Not found
		found = False 
		For Each objProcess in colProcessList
		   'MsgBox objProcess.Name
		   If objProcess.Name = ProcessName Then 
		   		  
		    	found = True
		    	Exit For 
		   End If  
		   		  	
		Next
		WScript.Sleep(3000)
		I = I + 1 
		If I > 300 Then Exit Do
		If Not found Then Exit Do
	
	Loop

End Function

Function strComputerName()

	Set objNetwork  = CreateObject("WScript.Network")
	strComputerName = objNetwork.ComputerName
	
	set objNetwork = Nothing
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''