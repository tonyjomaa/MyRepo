On Error Resume Next 

Set wshell = WScript.CreateObject ("Wscript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
CrashDumpFilePath = "D:\CrashDumps\"

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

Set objpathTextFile = objFSO.GetFile ("d:\Utils\path.txt")
Set objpathTextFileStream = objpathTextFile.OpenAsTextStream(1,0) ' for reading
SomeQAFolder = objpathTextFileStream.ReadLine
WScript.Sleep (500)
objpathTextFileStream.Close
Set objpathTextFileStream = Nothing 
Set objpathTextFile = Nothing 

Set objCrashDump = objFSO.GetFolder (CrashDumpFilePath).Files
File =""
NothingFound = "True"
'msgbox someqafolder
For Each File In objCrashDump
	
		'MsgBox File
		crashFileName = Replace (File,"D:\CrashDumps\","")
		'MsgBox crashFileName
	If InStr (crashFileName,"crash") > 0 Then 
		'copy to Report\QA folder
	objFSO.CopyFile File,SomeQAFolder&"\"
	WScript.Sleep (8000)
		'check if it was copied. If yes then delete local dump file. If no, then report and move on.
	'crashFileName = Replace (File,"D:\CrashDumps\","")
		If objFSO.FileExists (SomeQAFolder &"\"& crashFileName) Then 'returns "True or "False"
			 objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>Crash dump file was"&_
			 " generated. Copying crash dump file to "&SomeQAFolder&_
			 " was successful</font></th><th><font>PASSED</font></th></tr>")
			objFSO.DeleteFile File,True
			'NothingFound = "False"
			'Exit For 
		Else 
			objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>Crash dump file was"&_
			" generated. Copying crash dump file to "&SomeQAFolder&_
			" was NOT successful</font></th><th><font color=Red>FAILED</font></th></tr>")
			
		End If
		FlipColor color
		NothingFound = "False"  
	Else 
	
		If NothingFound = "False" Then 
			NothingFound = NothingFound
		Else 
	 		NothingFound = "True" 
		End If 
	'objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font> No crash dump file was"&_
	'		 "  generated</font></th><th><font>PASSED</font></th></tr>")
	
	'Exit For 
	End If 
Next

Set WshSysEnv = WShell.Environment("PROCESS")
userprofile = WshSysEnv("USERPROFILE")
DesktopPath = userprofile & "\Desktop"
'msgbox DesktopPath
If objfso.FileExists(DesktopPath&"\ZtCrash.dmp") Then 
  NothingFound="False"
  objFSO.MoveFile DesktopPath&"\ZtCrash.dmp",SomeQAFolder&"\"
  WScript.Sleep (8000)
End If

If NothingFound="True" Then 
	objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font> No crash dump file was"&_
			 "  generated</font></th><th><font>PASSED</font></th></tr>")
Else 
	'FlipColor color
	objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font> A crash dump file(s) was"&_
			 "  generated</font></th><th><font color=GoldenRod>WARNING</font></th></tr>")

End If 
FlipColor color	
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

Function FlipColor (color)
	
	If color = "3399FF" Then  
    	color = "99CCFF"
	Else 
		color = "3399FF"
	End If 
	
End Function 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''