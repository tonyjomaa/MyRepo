' ********************************************************************************
'
' Copyright (c) 2016, Algorithmic Implementations, Inc. (dba Ai Squared)
' All Rights Reserved.
'
' Developer:     Tony Jomaa
' 
' Description:   Evernix11Queue.vbs queues up testes on said machine
'
' Parameters:    needs build path and build version
'
' Keywords:      
'                
' Notes:         
'
' ********************************************************************************
'
On Error Resume Next 
Include "StartTestAndLog.vbs"
'************* declarations  *********************************************************//
'file="D:\ShareThis\EVXCommand11\AI2W0047TestTable.txt"
'machine = "AI2W0038"	
QueuePathFile = "D:\ShareThis\ZT11queue.txt"
emailaddress = "ajomaa@aisquared.com"  																	
'MachineTaskFolder = "\\ai2s-lab01\QALAB\ATS\ATM\atm_servers\" & machine & "\Tasks\"
InstallLocationFilePath = "\\ai2s-lab01\QALAB\ATS\ATM\atm_resources\InstallLocationZT11.txt"																				
' ******************************************************************************************//

Set WShell = CreateObject ("WScript.Shell") 
'  ****  master file object  ****
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set ObjQueueTextFile = objFSO.GetFile(QueuePathFile)
Set SendMail = CreateObject("CDO.Message")
Set objInstallLocationTextFile = objFSO.GetFile(InstallLocationFilePath)

Do 
WScript.Sleep (5000)

'If objfso.FolderExists ( MachineTaskFolder) Then
'	Set Taskfolder = objfso.GetFolder(MachineTaskFolder)
'	Set fileexist = Taskfolder.Files
'End If

Set ObjQueueTextstream = ObjQueueTextFile.OpenAsTextStream(1,-1) ' for reading
	
	done = "False"
	
'	SendMail.From = "EVX-Trigger-Server@aisquared.com"
'	SendMail.To = emailaddress
'	SendMail.Subject = "EVX Build Test Trigger"
'Do Until fileexist.Count=0 
'		WScript.Sleep (30000)    ' wait 30 seconds
'Loop 
	
'  ****  setup queue.txt file for reading  ****
'  ****  setup installLocation10.txt file  ****
	Do Until objQueueTextStream.AtEndOfStream

	line = objQueueTextStream.ReadLine
		
	 If objfso.FolderExists(line) Then 
    	strVerArry = Split(line,"\")
        Build = strVerArry(UBound(strVerArry))
		'If fileexist.Count=0 Then	

			Set objInstallLocationTextStream = objInstallLocationTextFile.OpenAsTextStream(2,-1)
				
				objInstallLocationTextStream.writeline(line)
				'WScript.Sleep (10000)
				objInstallLocationTextStream.Close
			Set objInstallLocationTextStream = Nothing 
   				''''''  *** This part is to send an email upon trigger *****
   				'SendMail.TextBody = "This path was written to Install Location file, and it triggered the test on machine name " & machine & ": " & line & "   " & "on: " & theDate & " " & thetime 
   				'SendMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
				'SendMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "AI2SMAIL01.aisquared.com"
				'SendMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = "25" 
				'SendMail.Configuration.Fields.Update
				'SendMail.Send
        	 	'WScript.Sleep (30000)			
   			    If objfso.FileExists(line&"\Installers\Trial\CD\MakeMe.vbs") Then wshell.Run line&"\Installers\Trial\CD\MakeMe.vbs silent",4,True
				If objfso.FileExists(line&"\Installers\Trial\FusionDVD\ComposeFusionDVDImage.vbs") Then wshell.Run line&"\Installers\Trial\FusionDVD\ComposeFusionDVDImage.vbs silent",4,True
				
				'Do Until fileexist.Count=0 
				'	WScript.Sleep (30000)    ' wait 30 seconds
				'Loop
		'wshell.Run "cmd /c svn.exe update D:\ShareThis\EVXCommand",4,True
		'wshell.Run "D:\ShareThis\EvernixScripts\EVXTestsAI2W0060.vbs "&line,4,False
		wshell.Run "D:\ShareThis\EvernixScripts\EVXTestsAI2WD0019.vbs "&line,4,False 
		WScript.Sleep(60000)
		wshell.Run "D:\ShareThis\EvernixScripts\ZT11TestsAI2WCYBER01.vbs "&line,4,False 				
		WScript.Sleep(120000)
		'WScript.Sleep(120000)
		'Process="EVXTestsAI2W-0013.vbs"
        'ProcessKill Process
	    wshell.Run "D:\ShareThis\EvernixScripts\ZT11TestsAI2W-0013.vbs "&line,4,False 
        WScript.Sleep(120000)
		WScript.Sleep(120000)
		wshell.Run "D:\ShareThis\EvernixScripts\ZT11TestsAI2W0047.vbs "&line,4,False
	    WScript.Sleep (120000)
	    ' Windows Eyes Tests
		'wshell.Run "D:\ShareThis\EvernixScripts\ZT11TestsAI2W-0018.vbs "&line,4,False
		WScript.Sleep(120000)
	    'WScript.Sleep(120000)
	    ' Windows Eyes Tests
		wshell.Run "D:\ShareThis\EvernixScripts\ZT11TestsAI2W0042.vbs "&line,4,False
	    WScript.Sleep (120000)	
	    wshell.Run "D:\ShareThis\EvernixScripts\ZT11TestsAI2W0038.vbs "&line,4,False
	    'WScript.Sleep (120000)				    	    
										    								
  line = ""
					
  'End If
  
  done = "True"
  line = ""
  
  Else
   Exit Do
  End If 

Loop
objQueueTextStream.Close
WScript.Sleep (5000) 
'*****   clear the queue contents
If done = "True" Then 
	Set objQueueTextStream = ObjQueueTextFile.OpenAsTextStream(2,-1)
	objQueueTextStream.Close
End If 

WScript.Sleep (5000)
buildpath=""
Result=""

Loop 

Function Include (vbsScriptFile)

  Set fso = CreateObject ("Scripting.FileSystemObject")
  if fso.FileExists(vbsScriptFile) then
    Set f = fso.OpenTextFile (vbsScriptFile)
    s = f.ReadAll()
    f.Close
    ExecuteGlobal s
  else
    wscript.echo wscript.ScriptFullName & ": Include - could not find file (" & vbsScriptFile & ")"
  end if

End Function