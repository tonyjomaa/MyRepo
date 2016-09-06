On Error Resume Next 
Set objFSO = CreateObject  ("scripting.filesystemobject")

'''''''''''''''''''''''''''Get color code ''''''''''''''''''''''''''
Set objcolorTextFile = objFSO.GetFile ("d:\Utils\color.txt")
Set objcolorTextFileStream = objcolorTextFile.OpenAsTextStream(1,0)  ' for Reading
color = objcolorTextFileStream.Readline '("3399FF")
WScript.Sleep (50)
objcolorTextFileStream.close
Set objcolorTextFile = Nothing 
Set objcolorTextFileStream = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Get folder name for later use
Set pathfile=objfso.GetFile("d:\utils\path.txt")
Set pathfiletext=pathfile.OpenAsTextStream(1,0)
buildpath = pathfiletext.ReadLine
WScript.Sleep(500)
pathfiletext.Close
Set pathfile=Nothing
Set pathfiletext=Nothing

If WScript.Arguments.Count=0 Then

  resultfile = "\\ai2s-lab02\D\ShareThis\EVXResults\ResultsEvxStabilityWin864.txt"
  
Else

  resultfile=WScript.Arguments.Item(0)
  
End If


Set objlogTextFile = objFSO.GetFile ("d:\Utils\Results.txt")
Set objlogTextFileStream = objlogTextFile.OpenAsTextStream(1,-1)
statusPF = 0
statusW = 0
strline = objlogTextFileStream.ReadAll
statusPF = InStr (strline,"FAILED")	' see if it failed. case insensitive
statusW =  InStr (strline,"WARNING")

If statusPF = 0 Then
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f2 = fso.GetFile("d:\Utils\grade.txt")
	Set f2 = f2.OpenAsTextStream(2, -1)
		f2.WriteLine("PASSED")
		Result = "PASSED"
		WScript.Sleep(1000)
Else
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f2 = fso.GetFile("d:\Utils\grade.txt")
	Set f2 = f2.OpenAsTextStream(2, -1)
		f2.WriteLine("FAILED")
		Result = "FAILED"
		WScript.Sleep(1000)
End If 

If statusW <> 0 And statusPF = 0 Then
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f2 = fso.GetFile("d:\Utils\grade.txt")
	Set f2 = f2.OpenAsTextStream(2, -1)
		f2.WriteLine("WARNING")
		Result = "WARNING"
		WScript.Sleep(1000)
End If 

	f2.Close
	Set f2 = Nothing	
	Set fso = Nothing

buildpath = buildpath &"\Results.html"

If objfso.FileExists(resultfile) Then		
	Set objstream = objFSO.OpenTextFile (resultfile,2,True,0)
	objstream.WriteLine (buildpath & " " & Result)
	WScript.Sleep (1000)
	objstream.Close



	Set objstream = objFSO.OpenTextFile (resultfile,1,False,0)
	line = objstream.ReadLine
	
	Set objlogTextFile = objFSO.GetFile ("d:\Utils\Results.txt")
	Set objlogTextFileStream = objlogTextFile.OpenAsTextStream(8,-1)  'for appending
	If InStr (line,"\\ai2s0017\") > 0 Or InStr (line,"\\AI2S0017\") > 0 Then 
		objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>Passing results back to SQL was successful</font></th><th><font>TRUE</font></th></tr>")
			'objlogTextFileStream.writeline("<Action>Copying log files to " & logfilepath & " has PASSED</Action><br>" )
	Else 
			objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>Passing results back to SQL was successful</font></th><th><font color=Red>FALSE</font></th></tr>")
			'objlogTextFileStream.writeline("<Action>Copying log files to " & logfilepath & " has FAILED</Action><br>" )
	End If 
	
	objstream.Close
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

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function LogToFileFunction(CommandPerformance,MachineTaskFolder,Stamp,action)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

QueueLogFile = "\\ai2s-lab02\D\ShareThis\QueueLogFile.txt"
Set objFSO2 = CreateObject ("Scripting.FileSystemObject")
Set ObjQueueLogFile = objFSO.GetFile(QueueLogFile)
Set ObjQueueLogFileStream = ObjQueueLogFile.OpenAsTextStream(8,-1) ' open file for appending

	ObjQueueLogFileStream.WriteLine (Stamp & " | " & action & " | " & CommandPerformance & " | " & MachineTaskFolder)
	ObjQueueLogFileStream.WriteLine ("")

WScript.Sleep(500)
ObjQueueLogFileStream.Close
Set ObjQueueLogFileStream = Nothing 
Set ObjQueueLogFile = Nothing 
Set objFSO2 = Nothing 

End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
