' ********************************************************************************
'
' Copyright (c) 2017, Abacus Automation
' All Rights Reserved.
'
' Developer:     Tony Jomaa
' 
' Description:   EVXTestsAI2WD0019.vbs queues up testes on said machine
'
' Parameters:    needs build path and build version
'
' Keywords:      
'                
' Notes:         
'
' ********************************************************************************
On Error Resume Next 
Include "StartTestAndLog.vbs"
'************* declarations  *********************************************************//
file="D:\ShareThis\EVXCommand\AI2WD0019TestTable.txt"
machine = "AI2WD0019"	
																		'	//
MachineTaskFolder =  "\\"&machine &"\QALAB\ATS\ATM\atm_servers\" & machine & "\Tasks\"		'   //	
InstallLocationFilePath = "\\"&machine &"\QALAB\ATS\ATM\atm_resources\InstallLocationEvernix.txt"
																				'	//
' ******************************************************************************************//

Set WShell = CreateObject ("WScript.Shell") 
'  ****  master file object  ****
Set objFSO = CreateObject("Scripting.FileSystemObject")
line=WScript.Arguments.Item(0)
'  ****  pause 5 seconds  ****
strVerArry = Split(line,"\")
Build = strVerArry(UBound(strVerArry))	

Set objInstallLocationTextFile = objFSO.GetFile(InstallLocationFilePath)

	'  ****  setup the file that determines if machine2 is done testing or not  ****
	If objfso.FolderExists ( MachineTaskFolder) Then
		Set Taskfolder = objfso.GetFolder(MachineTaskFolder)
		Set fileexist = Taskfolder.Files
	End If
		
	Do Until fileexist.Count=0
		WScript.Sleep (30000)    ' wait 30 seconds
	Loop 
		
If fileexist.Count=0 Then	
	
				Do Until fileexist.Count=0 
					WScript.Sleep (30000)    ' wait 30 seconds
				Loop
				Set objInstallLocationTextStream = objInstallLocationTextFile.OpenAsTextStream(2,-1)
				
				objInstallLocationTextStream.writeline(line)
				WScript.Sleep (10000)
				objInstallLocationTextStream.Close
				Set objInstallLocationTextStream = Nothing 
	
	Set fso=CreateObject("Scripting.FileSystemObject")
	Set open=fso.OpenTextFile(file,1)
	PrevCS=""
	copyCS=False
	Do While Not open.AtEndOfStream
	
	 str=open.ReadLine
	 If InStr(1,str,"'")=1 Then 
	   'it is a comment - do nothing
	 Else
	  
	   If str<>"" And str<>" " Then 
	     arry = Split(str," ")
	     TestID=arry(0)
	     CS=arry(1)
	     RF=arry(2)
	     If CS<>PrevCS Then
	       PrevCS=CS
	       'wait till previous test is done
	       copyCS=True
	       Do Until (fileexist.Count=0) And (fso.FolderExists(MachineTaskFolder))
		     WScript.Sleep (30000)    ' wait 30 seconds
		   Loop
	     Else
	       copyCS=False
	     End If
	     ' Start test 
	     StartTestFunc TestID,CS,RF,MachineTaskFolder,Build,copyCS
	   End If
	 End If 
	Loop
End If 				
WScript.Quit

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