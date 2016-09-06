On Error Resume Next 
'--------------------------------------------
'
' FindProcesses
'
' strNameList - used in WMI query to find process LIKE given name
'
' returns - comma separated list of processes matching query string
'
'--------------------------------------------

Function FindProcesses (strNameLike)

  Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\localhost\root\cimv2")

  Set colProcess = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name LIKE '" & strNameLike & "'")

  strProcessList = ""

  For Each objProcess in colProcess

    If Len(objProcess.Name) Then

      If Len(strProcessList) then

        strProcessList = strProcessList & "," & objProcess.Name

      Else

        strProcessList = objProcess.Name

      End If

    End If

  Next

  FindProcesses = strProcessList

  Set colProcess = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name LIKE 'Zt.exe'")

  strProcessList = ""

  For Each objProcess in colProcess

    If Len(objProcess.Name) Then

      If Len(strProcessList) then

        strProcessList = strProcessList & "," & objProcess.Name

      Else

        strProcessList = objProcess.Name

      End If

    End If

  Next

  FindProcesses = FindProcesses & strProcessList

End Function


'--------------------------------------------
'
' FilterExpectedProcessNames
'
' strNameList - comma separated list of process names
'
' returns - strNameList with expected process names removed
'
'--------------------------------------------

Function FilterExpectedProcessNames (strRemainingList)  'exemption function

  ' TODO: add list of expected applications
  'arrayExpectedProcesses = Array("csrss.exe","cctray.exe","cmd.exe", "chrome.exe", "conhost.exe", "cscript.exe") ''"AiSquared.Logging.SLogServer.exe",
  arrayExpectedProcesses = Array("AiSquared.Magnification.Service.exe","ZtUAC.exe", "ZtUAC64.exe", "conhost.exe", "cscript.exe","AiSquared.ZoomText.News.exe","AiSquared.Loader.Generic.exe","AiSquared.Loader.Elevated.exe")

  FilterExpectedProcessNames = strRemainingList

  For Each name In arrayExpectedProcesses

    FilterExpectedProcessNames = Replace(FilterExpectedProcessNames, name, " ")
    
  Next

  FilterExpectedProcessNames = Replace(FilterExpectedProcessNames, " ,", "")
  FilterExpectedProcessNames = Replace(FilterExpectedProcessNames, ", ", "")

  If FilterExpectedProcessNames = " " Then
    FilterExpectedProcessNames = ""
  End If

End Function
'--------------------------------------------------

Function FlipColor (color)
	
	If color = "3399FF" Then  
    	color = "99CCFF"
	Else 
		color = "3399FF"
	End If 
	
End Function 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

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

Function strComputerName()

	Set objNetwork  = CreateObject("WScript.Network")
	strComputerName = objNetwork.ComputerName
	
	set objNetwork = Nothing
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function CallAppCheckFunction
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	Set WShell2 = CreateObject ("WScript.Shell")

	flag = WShell2.Run ("q:\ATS\AutoTasks\ZTExists.exe", 4, True)
'MsgBox flag
	DoneTime = Time 
	'FlipColor color
	If flag = 1 Then 
	 
		objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>ZoomText application window exists on the Desktop at</font></th><th><font>"&Time&"</font></th></tr>")
		FlipColor color
		
	Else 
	
		objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>ZoomText application window does not exist on the Desktop at</font></th><th><font>"&Time&"</font></th></tr>")
		FlipColor color
		
	End If 
	Set WShell2 = Nothing 
	
End Function 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function CallAeroCheckFunction()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	Set WShell2 = CreateObject ("WScript.Shell")

	flag = WShell2.Run ("q:\ATS\AutoTasks\AeroCheck.exe", 4, True)

	DoneTime = Time 
	
	If flag = 0 Then 
	 
		objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>Checking if Aero Glass is enabled at " & Time & "</font></th><th><font>ENABLED</font></th></tr>")
		FlipColor color
		
	Else 
	
		objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>Checking if Aero Glass is enabled at " & Time & "</font></th><th><font color=GoldenRod>DISABLED</font></th></tr>")
		FlipColor color
		'CallDWMRestartFucntion "dwm.exe"
		'FindProcesses "dwm.exe"
		
	End If 
	Set WShell2 = Nothing 
	
End Function 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function GetMarkFunction(MARK)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	istime = (DateDiff("d",date,"8/31/2011"))
  
 	If istime <=0 Then 
		MARK = "color=Red>FAILED"
	Else 
		MARK = "color=GoldenRod>WARNING"
	End If
	
	
End Function 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''' MAIN ''''''''''''''''''''''''''''''''''
Set wshell = WScript.CreateObject ("Wscript.Shell")
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

Set objlogTextFile = objFSO.GetFile ("d:\Utils\Results.txt")
Set objlogTextFileStream = objlogTextFile.OpenAsTextStream(8,-1) ' for appending


secondsToWait = 5
numberOfWaits = 25  ' total wait time = secondsToWait x numberOfWaits

bProcessesFound = false

For i = 1 To numberOfWaits

  strRemainingList = FindProcesses("aisquared%")
  strZtRemList = FindProcesses("Zt%")

  strRemainingList = FilterExpectedProcessNames(strRemainingList)
  strZtRemList = FilterExpectedProcessNames(strRemainingList)
  
  If Len(strRemainingList) Or Len(strZtRemList) Then

    bProcessesFound = true

    arrayProcessNames = Split(strRemainingList, ",")
	arrayProcessNames2 = Split(strZtRemList,",")
	
	If UBound(arrayProcessNames)>0 then
      For Each process In arrayProcessNames
	    objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>ZT process still in memory " & process & "</font></th><th><font>" & Time & "</font></th></tr>")	
        FlipColor color	
      Next
	End If
	
	If UBound(arrayProcessNames2)>0 then
      For Each process In arrayProcessNames2
	    objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>ZT process still in memory " & process & "</font></th><th><font>" & Time & "</font></th></tr>")	
        FlipColor color	
      Next
    End If
    
    If i < numberOfWaits Then

      ' I find the timings helpful
      'WScript.Echo "processes found, waiting " & secondsToWait & " seconds to check again (" & i & " of " & numberOfWaits & ")..."
	  objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>Processes found, waiting " & secondsToWait & " seconds to check again (" & i & " of " & numberOfWaits & ")...</font></th><th><font>" & Time & "</font></th></tr>")	
       FlipColor color
      CallAeroCheckFunction
     
      CallAppCheckFunction
      WScript.Sleep(secondsToWait * 1000)
	  'FlipColor color
	  
    Else

      ' I find the timings helpful
      'WScript.Echo "processes found, but max wait hit (" & (i * secondsToWait) & ")..."
	  objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>Processes found, but maximum wait was reached (wait time is set to " & (i * secondsToWait) & " seconds)... " & process & "</font></th><th><font>" & Time & "</font></th></tr>")	
      FlipColor color
      
    End If

  Else

    'WScript.Echo "nothing found, exiting..."
    bProcessesFound = false

    Exit For

  End If

Next

' TODO: log overall status
If Not bProcessesFound Then

  'WScript.Echo "no match running :)"
  objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>Check if ZT processes closed at " & Time & "</font></th><th><font>PASSED</font></th></tr>")	
  FlipColor color
  
Else

  'WScript.Echo "complete filure :("
  MARK =""
  GetMarkFunction MARK
  objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>Check if ZT processes closed at " & Time & "</font></th><th><font "&MARK&"</font></th></tr>")	
  FlipColor color
  
End If

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

ejectProcess()
Set objlogTextFileStream = Nothing
Set objlogTextFile = Nothing
Set objFSO = Nothing

Set wshell = Nothing

WScript.Quit
