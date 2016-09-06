On Error Resume Next 
'''''''
Set objCimv2 = GetObject("winmgmts:root\cimv2")
Set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set objFSO = CreateObject("Scripting.FileSystemObject")
strComputerName()

'''''''''''''''''''''''''''Get color code ''''''''''''''''''''''''''
Set objcolorTextFile = objFSO.GetFile ("d:\Utils\color.txt")
Set objcolorTextFileStream = objcolorTextFile.OpenAsTextStream(1,0)  ' for Reading
color = objcolorTextFileStream.Readline '("3399FF")
WScript.Sleep (50)
objcolorTextFileStream.close
Set objcolorTextFile = Nothing 
Set objcolorTextFileStream = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Set objMemory = objRefresher.AddEnum _
    (objCimv2, _ 
    "Win32_PerfFormattedData_PerfOS_Memory").ObjectSet

' Initial refresh needed to get baseline values
objRefresher.Refresh
intTotalHealth = 0
    
    For each intAvailableBytes in objMemory
        
            MemLeft = intAvailableBytes.AvailableMBytes
        
    Next

'*******  Read system memory   *****************************
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputerName & "\root\cimv2")
Set colSettings = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
For Each objComputer in colSettings
getTotalMemory = objComputer.TotalPhysicalMemory
Next
getTotalMemory = getTotalMemory /1000000
getTotalMemory = Left (getTotalMemory,4)

getTotalMemory = CInt(getTotalMemory)
percent = (getTotalMemory - MemLeft)/getTotalMemory * 100

percent = Round(percent)
'MsgBox "Physical Memory percentage used "&percent&"%"
Set objlogTextFile = objFSO.GetFile ("d:\Utils\Results.txt")
Set objlogTextFileStream = objlogTextFile.OpenAsTextStream(8,-1) 'For appending
delay = 50

If percent <= 45 Then 
	objlogTextFileStream.WriteLine("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>Physical Memory (RAM) usage at " & time & " is " & percent & "%</font></th><th><font color=green>PASSED</font></th></tr>")
	FlipColor color
End If 
If percent > 45 And percent <= 50 Then 
	objlogTextFileStream.WriteLine("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>Physical Memory (RAM) usage at " & time & " is " & percent & "%</font></th><th><font color=goldenrod>WARNING</font></th></tr>")
	FlipColor color
End If
If percent > 50 Then 
	objlogTextFileStream.WriteLine("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>Physical Memory (RAM) usage at " & time & " is " & percent & "%</font></th><th><font color=red>FAILED</font></th></tr>")
	FlipColor color
End If

WScript.Sleep (delay)

'strComputerName()
'Set objWMIService2 = GetObject("winmgmts:\\" & strComputerName & "\root\CIMV2") 
'Set colItems2 = objWMIService2.ExecQuery("SELECT * FROM Win32_PerfFormattedData_PerfOS_Processor WHERE Name = '_Total'") 
'For Each objItem in colItems2 
'CPU = objItem.PercentProcessorTime 
'next 

'objlogTextFileStream.WriteLine("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>Processor (CPU) usage at " & time & "</font></th><th><font>" & CPU & "%</font></th></tr>")
'FlipColor color
'WScript.Sleep (delay)

objlogTextFileStream.Close
Set objlogTextFileStream = Nothing
Set objlogTextFile = Nothing

'''''''''''''''''''''''''''Flip color code and write to file''''''''''''''''''''''''''
	Set objcolorTextFile = objFSO.GetFile ("d:\Utils\color.txt")
	Set objcolorTextFileStream = objcolorTextFile.OpenAsTextStream(2,0)  ' for writing
	
	objcolorTextFileStream.Writeline (color)
	WScript.Sleep (50)
	objcolorTextFileStream.Close
	Set objcolorTextFile = Nothing 
	Set objcolorTextFileStream = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

CPU = Nothing
Set objWMIService2 = Nothing
Set colItems2 = Nothing
percent = Nothing
getTotalMemory = Nothing
Set colSettings = Nothing
Set objWMIService = Nothing
MemLeft = Nothing
Set objMemory = Nothing
Set objRefresher = Nothing
Set objCimv2 = Nothing
strComputerName = Nothing
color = Nothing

ejectProcess()
WScript.Quit

Function strComputerName()

	Set objNetwork  = CreateObject("WScript.Network")
	strComputerName = objNetwork.ComputerName
	
	set objNetwork = Nothing

End Function

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

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function FlipColor (color)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	If color = "FFFFFF" Then  
    	color = "99CCFF"
	Else 
		color = "FFFFFF"
	End If 
	
End Function 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
