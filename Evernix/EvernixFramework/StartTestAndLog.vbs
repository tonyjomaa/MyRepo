''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function LogToFileFunction(CommandPerformance,MachineTaskFolder,Stamp,action)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

QueueLogFile = "D:\ShareThis\QueueLogFile.txt"
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

Function startTestFunc(TestID,CS,RF,MachineTaskFolder,Build,copyCS)
  Set WShell = CreateObject ("WScript.Shell")
  Set fso=CreateObject("Scripting.FileSystemObject")
  
  If Not copyCS Then
    WScript.Sleep(4899)
  Else
    WScript.Sleep(2000)
  End If
  'Write to a management file
  ForceWriteFunc TestID,Build,CS
  
  If copyCS Then 
    
    If fso.FileExists(CS) Then 
      fso.CopyFile CS, MachineTaskFolder
      If Err.Number <> 0 Then
        Err.Clear
        LogToFileFunction "Error - while coping: " & CS, " to machine task folder: " & MachineTaskFolder, Now, "Is the network, or that file, available?"
      End If  
    Else
      LogToFileFunction "Error - could not find: " & CS, "", Now, "Is the network, or that file, available?"
    End If
  
  End If
  
  TaskFilePath = ""
  GetTaskFilePath TaskFilePath, CS, MachineTaskFolder
  
  If InStr(CS,"\ImageReader\") Then
    strcmd = "D:\ShareThis\EvernixScripts\EVXSmokeTestResultsToSQL.vbs "&RF&" "&TestID&" "&Build&" IR " & TaskFilePath
  Else
    strcmd = "D:\ShareThis\EvernixScripts\EVXSmokeTestResultsToSQL.vbs "&RF&" "&TestID&" "&Build & " " & TaskFilePath
  End If
  
  WShell.Run strcmd,4,False
  LogToFileFunction "Test started: "&TestID&"; Build: "&Build,MachineTaskFolder,Now,"Command: "&CS&"; Result in: "&RF
End Function

Function ForceWriteFunc(TestID,Build,CS)
  On Error Resume Next
  path="D:\ShareThis\ManageInsert.txt"
  Do
  Set fso= CreateObject("Scripting.FileSystemObject")

  If Not fso.FileExists(path) Then 
    Set file=fso.OpenTextFile(path,8,True)
 
  Else
    'Set file=fso.OpenTextFile(path,8,False)
    Set file=fso.GetFile(path).OpenAsTextStream(8)
    If Err.Number <> 0 Then 
      WScript.Sleep (3200)
      Set fso=Nothing
      Set file=Nothing
      Err.Clear
    Else
      If InStr(CS,"\ImageReader\") Then
        file.Writeline(TestID&" "&Build&" IR")
      Else
        file.Writeline(TestID&" "&Build)
      End If  
      file.close
      Set file=Nothing
      Exit Do
    End If
  End If
 Loop

End Function 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function ProcessKill(Process)
  strComputerName()

  Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputerName & "\root\cimv2")
  Set colProcessList = objWMIService.ExecQuery _
    ("Select * from Win32_Process")     
  For Each objProcess in colProcessList
   processName = objProcess.CommandLine
   	If InStr(processName,Process)>0 Then 
   	   objProcess.Terminate
   	   Exit For
    End If 
  Next

  Set processName = Nothing 
  Set colProcessList = Nothing
  Set objWMIService = Nothing
End Function

Function strComputerName()

	Set objNetwork  = CreateObject("WScript.Network")
	strComputerName = objNetwork.ComputerName
	
	set objNetwork = Nothing
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function GetTaskFilePath(TaskFilePath, CS, MachineTaskFolder)

	strTaskArry = Split(CS,"\")
	strFilename = strTaskArry(UBound(strTaskArry))
	
	If strFilename = "" Then Exit Function
	
	TaskFilePath = MachineTaskFolder & strFilename

End Function