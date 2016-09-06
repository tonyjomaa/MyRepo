' ********************************************************************************
'
' Copyright (c) 2011-2014, Algorithmic Implementations, Inc. (dba Ai Squared)
' All Rights Reserved.
'
' Developer:     Tony Jomaa
' 
' Description:   EVXSmokeTestResultsToSQL.vbs writes results to a file
'
' Parameters:    
'
' Keywords:      
'                
' Notes:         
'
' ********************************************************************************

On Error Resume Next 
	Set objFSO = CreateObject("Scripting.FileSystemObject")

	Resultfile = WScript.Arguments.Item(0)
	TestID = WScript.Arguments.Item(1)
	TestSPID = TestID
	testtable=""
	Build = WScript.Arguments.Item(2)
	TaskFilePath = WScript.Arguments.Item(3)
	If (WScript.Arguments.Count = 5) Then
	  testtable=WScript.Arguments.Item(3)
	  TaskFilePath = WScript.Arguments.Item(4)
	End If
	If (Resultfile="empty") Or (TestID=99999) Then 
	  WScript.Quit(0)
	End If

action = "Evernix Test look for result script "&WScript.ScriptFullName& " is launched."
Stamp = Now
LogToFileFunction Resultfile,TestID,Stamp,action
If Not objFSO.FileExists(Resultfile) Then
  LogToFileFunction Resultfile,TestID,Stamp,"Error- Result file doesn't exist"
  WScript.Quit(1)
End If

Set ObjResultTextFile = objFSO.GetFile(Resultfile)
' clear out file contents
Set objResultTextStream = ObjResultTextFile.OpenAsTextStream(2,0)
WScript.Sleep(100)
objResultTextStream.Close
Set objResultTextStream = Nothing
''' get set to read from the file
Set ObjResultTextFile = Nothing
count = 1
j=0

	Do 
	    
	    Set objResultTextStream = ObjFSO.OpenTextFile(Resultfile,1,False,0) 
		Do Until objResultTextStream.AtEndOfStream

			Resultline = objResultTextStream.ReadLine
			
			
			If InStr(Resultline,"\\")>0 Then 
				
				action = "Evernix Test > Results reported"
   				Stamp = Now
   				action = action & " "&TestID
   				LogToFileFunction Resultline,"",Stamp,action
				'WScript.Sleep (60000)
				'WScript.Sleep (30000)
				'WScript.Sleep (30000)
				'WScript.Sleep (30000)
				
				If ObjFSO.FileExists(TaskFilePath) Then
					objResultTextStream.Close
					Set objResultTextStream = Nothing
					Stamp = Now
					j=j+1
					LogToFileFunction "Iteration number " & j,TestID,Stamp,"Test CS file was found."
					WScript.Sleep (30000)
					Exit Do
				
				Else    
				
					Resultlinearray = Split (Resultline," ")
					Buildpath = Resultlinearray(0)
					Result = Resultlinearray(1)
				
					InsertToSQLTableFunction buildpath,Result,TestID,build,testtable',BuildResultInfo
					
					Set objResultTextStream = ObjFSO.OpenTextFile(Resultfile,2,False,0) ' clear the Result file
					objResultTextStream.Close
					
					Resultline = ""
					'WScript.Sleep (30000)
					WScript.Quit(0)
				
				End If
				 				
			End If 
			
			Resultline = ""
		Loop 
 WScript.Sleep (30000)
Set objResultTextStream = Nothing

If TestSPID = 2 Or TestSPID = 59 Or TestSPID = 78 Or TestSPID = 79 Or TestSPID = 80 Or TestSPID = 392 Then 
  If count >= 540 Then  ' 540 x 30 seconds = 4.5 hours
	timeoutfunc Build,TestSPID,True
	WScript.Quit(0)
  End If
ElseIf TestSPID = 199 Or TestSPID = 208 Or TestSPID = 193 Then
  If count >= 240 Then  ' 240 x 30 seconds = 2 hour
	timeoutfunc Build,TestSPID,True
	WScript.Quit(0)
  End If
Else
  If count >= 200 Then  ' 200 x 30 seconds = 100 minutes = 1 hour 40 minutes
	timeoutfunc Build,TestSPID,True
	WScript.Quit(0)
  End If
End If

count = count + 1

Loop 
		
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function InsertToSQLTableFunction (buildpath,Result,TestID,build,testtable)',BuildResultInfo)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  loc =""
  ResultDetails =""
  ArchiveResultsToRemoteLoc buildpath,ResultDetails,TestID,loc
  Set sh = CreateObject("WScript.Shell")
  Set FSO = CreateObject ("Scripting.FileSystemObject")  
  
  If Not FSO.FileExists(buildpath) Then
    ResultDetails2 = buildpath&"-FILE-NOT-FOUND"
  Else
    ResultDetails2=ResultDetails
  End If
  	  
  ForceWriteFunc TestID,Build,Result,ResultDetails2,testtable
	  
  If flag = 0 Then
    action = "Evernix Test > Results into SP - Succeeded"
   	Stamp = Now
   	action = action & " "&TestID &" | Result="&Result&" | "&ResultDetails
   	LogToFileFunction Resultline,"",Stamp,action
  Else
    action = "Evernix Test > Results into SP - Failed"
   	Stamp = Now
   	action = action & " "&TestID &" | Result="&Result&" | "&ResultDetails
   	LogToFileFunction Resultline,"",Stamp,action
  End if
	  
End Function 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

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

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function timeoutFunc(Build,TestSPID,emailme)
	Set SendMail = CreateObject("CDO.Message")
	emailaddress = "ajomaa@aisquared.com"
	SendTimedOutToSQLFunc Build,TestSPID
	action = Build & " test timed out = "&TestSPID
   	Stamp = Now
   	LogToFileFunction Resultfile,"",Stamp,action
	
	If emailme	Then
	
		SendMail.From = "AutoTest-ResultManager@aisquared.com"
		SendMail.To = emailaddress
		SendMail.Subject = "10.1 Build "& Build &" - Test ID "& TestSPID&" Timed Out"	
		SendMail.TextBody = "Test ID "& TestSPID&" did not report results after one hour of wait time. Build version = " & Build & vbcrlf & "This has occurred on " & Now & "; " & vbCrLf & Resultfile
		SendMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		SendMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "AI2SMAIL01.aisquared.com"
		SendMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = "25" 
		SendMail.Configuration.Fields.Update
		SendMail.Send
		WScript.Sleep(500)	
	
	End If
	

End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function SendTimedOutToSQLFunc(Build,TestSPID)

	TestID = TestSPID
	TestID = CStr(TestID)
	
TestType=""
GetTestTypeFunc TestID,TestType	
If InStr(TestType,"Regression")>0 Then

	tablePathname = "D:\SharePoint\TestResults\11Regression.accdb"
	tablename = "`TestAutomationResults-11Regression`"
Else
    tablePathname = "D:\SharePoint\TestResults\11Other.accdb"
    tablename = "`TestAutomationResults-11Other`"
End If		
	ResultDetails = "Timed Out"
    Result = "FAILED"
    Note = "Test timed out"
    dateandtime = Date & " " & Time
    ResultColor = "http://ai2s_spps/AI2%20Pictures/red.jpg"
  
	Field1 = "`ID`"
	Field2 = "`Test ID`"
	Field3 = "`Test Type`"
	Field4 = "`Test Name`"
	Field5 = "`Build`"
	Field6 = "`ResultColor`"
	Field7 = "`Result`"
	Field8 = "`Result Details`"
	Field9 = "`Test Started`"
	Field10 = "`Test Completed`"
	Field11 = "`Qa Owner`"
	Field12 = "`PE Lead`" 
	Field13 = "`Framework Data Base Key`"
	Field14 = "`note`"
	connection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & tablePathname & ";Persist Security Info=False;" 

	Set AccessObj = CreateObject ("adodb.connection")
 	AccessObj.Open Connection
 	
 	ResultDetails2 = Replace (ResultDetails,"\\","\")
	'SQL = "UPDATE " & tablename & " SET "&Field6&" = '" & ResultColor & "', "&Field7&" = '"&ResultDetails&"', "&Field9&" = '"&dateandtime&"' WHERE " & tablename & "."&Field5&" = '"&Build&"' AND " & tablename & "."&Field4&" = 'Deployment Smoke Test Win7 32 bit'"
  SQL = "UPDATE " & tablename & " SET "&Field6&" = '" & ResultColor & "', "&_
  Field7&" = '"&Result&"', "&Field8&" = '"&ResultDetails2&"', "&Field10&_
  " = '"&dateandtime&"', "&Field14&" = '" & Note & "' WHERE " & tablename & "."&Field5&" = '"&Build&"' AND " & tablename &_
   "."&Field2&" = '"&TestID&"' AND " & tablename & "."&Field7&" = 'In Process' "

Set RS = AccessObj.Execute (SQL)
 	Set RS = Nothing 
 	AccessObj.Close 
End Function


Function ArchiveResultsToRemoteLoc(buildpath,ResultDetails,TestID,loc)

  Loc = "D:\ShareResults\ZT\"
  OS =""
  If InStr(buildpath,"\Win732\")>0 Then OS = "Win732"
  If InStr(buildpath,"\Win764\")>0 Then OS = "Win764"
  If InStr(buildpath,"\Vista32\")>0 Then OS = "Vista32"
  If InStr(buildpath,"\Vista64\")>0 Then OS = "Vista64"
  If InStr(buildpath,"\XP\")>0 Then OS = "XP"
  If InStr(buildpath,"\Win832\")>0 Then OS = "Win832"
  If InStr(buildpath,"\Win864\")>0 Then OS = "Win864"
  If InStr(buildpath,"\Win8164\")>0 Then OS = "Win8164"
  If InStr(buildpath,"\Win8132\")>0 Then OS = "Win8132"
  If InStr(buildpath,"\Win1032\")>0 Then OS = "Win1032"
  If InStr(buildpath,"\Win1064\")>0 Then OS = "Win1064"
  If OS <> "" Then 
    loc = loc & OS & "\"
  Else
    'WScript.Quit(1)
  End If
    
  patharry = Split(buildpath,"\")
  number = UBound(patharry)
  number = number -1
  pathAssemble = ""
  For I = 1 To number
   pathAssemble = pathAssemble & "\" & patharry(I)
  Next
  localpath=Replace(pathAssemble,"\\ai2s0017\ts01\GROUP\BUILDSTORAGE\Images\Evernix\Mainline\","",1,-1,vbTextCompare) 
  localpath=Replace(localpath,"\\ai2s0017\TS01\group\BUILDSTORAGE\Images\ImageReader\mainline\","",1,-1,vbTextCompare)
    
  testname = patharry(UBound(patharry)-1)
  loc = loc&localpath 
  
  ResultDetails=loc&"\"&patharry(UBound(patharry))
  ResultDetails=Replace(ResultDetails,"D:\","\\ai2s-lab02\D\")
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set sh = CreateObject("Wscript.Shell")
  
  If Not fso.FileExists(buildpath) Then
    Stamp=Now
    action="result file was not found on AI2S0017 archive server"
    LogToFileFunction "Test ID: "&TestID,"No result file was found: "&buildpath,Stamp,action
  End If
  
  If Not(fso.FolderExists(loc)) Then 
    
    Do
      sh.Run "cmd /c mkdir "& loc,4,True
      WScript.Sleep(5000)
      If (fso.FolderExists(loc)) Then 
        Exit Do
      Else
        WScript.Sleep(65555)
        sh.Run "cmd /c mkdir "& loc,4,True
      End If
    Loop   
 End If 
  
   If fso.FolderExists(pathAssemble) Then 
    'If TestID<>"132" And TestID<>"167" And TestID<>"136" And TestID<>"163" And TestID<>"137" And TestID<>"131" Then
        'fso.CopyFolder pathAssemble, loc, True
        sh.Run "cmd /c robocopy " & pathAssemble & " " & loc & " /E",4,True
        WScript.Sleep(10000)
      
        I=1
        do 
      
          if fso.FileExists(loc&"%.html") or I>1 then 
            exit do
          else
            wscript.sleep(60000)
            I=I+1
          end if
        Loop
      'fso.CopyFolder pathAssemble, loc, True
      sh.Run "cmd /c robocopy " & pathAssemble & " " & loc & " /E",4,True
    'else
       WScript.Sleep(25000)
    'end If
  End If   
  
      WScript.Sleep(10000)
  						
End Function   

Function GetTestTypeFunc(TestID,TestType)

tablePathname = "d:\SharePoint\automatedtestdefinitions.accdb"

	tablename = "`Automated test definitions`"
	Field1 = "`ID`"
	Field2 = "`Test Type`"
	Field3 = "`Friendly name`"
	Field4 = "`TCDB number(s)`"
	Field5 = "`QA lead`"
	Field6 = "`PE Lead`"
	Field7 = "`Put into Production`"
	Field8 = "`Failure action`"
	FieldLastRun = "`Last run`"
	Field10 = "`Box`"
	Field11 = "`Platform`"
	
	connection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & tablePathname & ";Persist Security Info=False;" 

	Set AccessObj = CreateObject ("adodb.connection")
 	AccessObj.Open Connection
 	
 ''''' Test Type
  	SQL = "SELECT "&Field2&" FROM "&tablename&" WHERE "&Field1&" = "&TestID&""
	Set RSAccess = AccessObj.Execute (SQL)
 	TestType = RSAccess.GetString
 	WScript.Sleep(1000)
 	TestType=Replace(TestType,vbCrLf,"")
 	TestType=Replace(TestType,vbLf,"")
 	TestType=Replace(TestType,vbCr,"")
 	TestType=Trim(TestType)
 	Set RSAccess = Nothing 

End Function

Function ForceWriteFunc(TestID,Build,Result,ResultDetails2,testtable)
  On Error Resume Next
  path="D:\ShareThis\ManageInsert.txt"
  Do
  Set fso= CreateObject("Scripting.FileSystemObject")

  If Not fso.FileExists(path) Then 
    Set file=fso.OpenTextFile(path,8,True)
  Else
    Set file=fso.GetFile(path).OpenAsTextStream(8)
    If Err.Number <> 0 Then 
      WScript.Sleep (3100)
      Set fso=Nothing
      Set file=Nothing
      Err.Clear
      Stamp=Now
      action="Could not get file: "&path
      LogToFileFunction "Test ID: "&TestID,"Did not write "&build& "; Result: "&Result,Stamp,action
    Else
      file.Writeline(TestID&" "&Build&" "&Result&" "&ResultDetails2&" "&testtable)
      If Err.Number <> 0 Then 
        WScript.Sleep (3100)
        Set fso=Nothing
        Set file=Nothing
        Err.Clear
        Stamp=Now
        action="Could not write file: "&path
        LogToFileFunction "Test ID: "&TestID,"Did not write "&build& "; Result: "&Result,Stamp,action
      Else
        file.close
        Set file=Nothing
        Stamp=Now
        action="Success writing to file: "&path
        LogToFileFunction "Test ID: "&TestID,"FW was able to write "&build& "; Result: "&Result,Stamp,action
        Exit Do
      End If
    End If
  End If
 Loop

End Function 