On Error Resume Next 

Set WShell = WScript.CreateObject ("WScript.Shell")
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

 
'Set objlogTextFile = objFSO.CreateTextFile("d:\Utils\Results.txt")
Set objlogTextFile = objFSO.GetFile ("d:\Utils\Results.txt")
Set objlogTextFileStream = objlogTextFile.OpenAsTextStream(8,-1) 'For appending

InstallLocationFilePath = "q:\ATS\ATM\atm_resources\installLocationEvernix.txt"
Set objInstallLocationTextFile = objFSO.GetFile(InstallLocationFilePath)
Set objInstallLocationTextStream = objInstallLocationTextFile.OpenAsTextStream(1,-1)  'open for reading

strlevel = ""
'WScript.Sleep(1000)
'Set ZT = CreateObject("Zoomtext.application")
WScript.Sleep(5000)
Set ZT = CreateObject("Zoomtext.application")
I = 0
IsZTReady = False
IsZTReadyfunc IsZTReady,I
If IsObject(ZT) Then
	If IsZTReady then
	 	Set mag = zt.Magnification
	 	Set test = mag.PrimaryWindow
	 	low = test.Enabled
	 	
	 	mag.PrimaryWindow.Power.Level="12"
	 	objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>ZoomText scripting object was loaded successfully in "& I * 5 &" seconds of wait time</font></th><th><font>TRUE</font></th></tr>")
		Set strlevel = mag.PrimaryWindow.Power
		strlevel=strlevel.Level
		WScript.Sleep(50)
		FlipColor
		If low Then
			objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>ZoomText is indeed enabled</font></th><th><font>TRUE</font></th></tr>")
			FlipColor
		Else
			objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>ZoomText is indeed enabled</font></th><th><font color=Red>FALSE</font></th></tr>")
			FlipColor
		End If
		If strlevel = "12" Then
			objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>ZoomText power magnification was set successfully</font></th><th><font>TRUE</font></th></tr>")
			FlipColor
		Else
			objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>ZoomText power magnification was not set successfully</font></th><th><font color=Red>FALSE</font></th></tr>")
			FlipColor
		End If
	Else
	 	objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>ZoomText scripting object was not loaded successfully after this many seconds of wait time</font></th><th><font color=Red>"& I * 5 &"</font></th></tr>")
		FlipColor
	End If
End If 

Set ie = CreateObject("InternetExplorer.Application")
ie.Visible = True
ie.Navigate2 "http://www.aisquared.com"
ie.Navigate2 "http://www.cnn.com",2048
ie.Navigate2 "http://www.yahoo.com",2048
ie.Navigate2 "https://news.google.com/nwshp?hl=en&tab=wn",2048
ie.Navigate2 "http://www.nbc.com",2048
ie.Navigate2 "http://www.bbc.com",2048
ie.Navigate2 "http://www.greenrenaissance.co.za/",2048
WScript.Sleep (4000) 'allow four seconds

set sh= createobject("WScript.Shell")
  If sh.AppActivate("Internet Explorer") Then 
    objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>Internet Explorer was loaded successfully</font></th><th><font>TRUE</font></th></tr>")
	FlipColor
	wshell.SendKeys("%{F4}")
	WScript.Sleep (500)
  End If
  ie.Quit
  AppFound = False

CheckProcess "iexplore.exe",AppFound
If AppFound Then 
  stopprocessFunction "iexplore.exe"
End If

Set ie = Nothing

	'''''''''''''''''''''''''''Write color code ''''''''''''''''''''''''''
		Set objcolorTextFile = objFSO.GetFile ("d:\Utils\color.txt")
		Set objcolorTextFileStream = objcolorTextFile.OpenAsTextStream(2,0)  ' for writing
		'color = objcolorTextFileStream.Readline '("3399FF")
	
		objcolorTextFileStream.Writeline (color)
		WScript.Sleep (50)
		objcolorTextFileStream.close
		Set objcolorTextFile = Nothing 
		Set objcolorTextFileStream = Nothing
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'wait 10 seconds before ejecting - give ZT 10 seconds of runtime
I = 1
For I=1 To 10
 WScript.Sleep(1000)
Next	
ejectProcess()
'WScript.Quit

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
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Function IsZTReadyFunc(IsZTReady,I)

	For I = 1 To 24 ' wait 2 minutes to install ZT
		Set ZT = CreateObject("Zoomtext.application")
		If IsObject(ZT) Then
		 	IsZTReady = True 
		 	Exit For
		Else
		 	IsZTReady = False
		End If
		WScript.Sleep(5000)
		'Set ZT = Nothing
	Next
	   
	
End Function

Function strComputerName()

	Set objNetwork  = CreateObject("WScript.Network")
	strComputerName = objNetwork.ComputerName
	
	set objNetwork = Nothing
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function FlipColor
	
	If color = "FFFFFF" Then  
    	color = "99CCFF"
	Else 
		color = "FFFFFF"
	End If 
	
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

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function stopprocessFunction(process)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputerName & "\root\cimv2")

Set colProcessList = objWMIService.ExecQuery _
    ("Select * from Win32_Process")     

For Each objProcess in colProcessList
   'MsgBox objProcess.Name
   If objProcess.Name = process Then 
   		objProcess.Terminate()
   		Exit For 
   End If 
Next
	Set colProcessList = Nothing
	Set objWMIService = Nothing 

End Function 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
