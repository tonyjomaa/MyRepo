On Error Resume Next 

Dim strInfile (200), filename (200), strInfile2 (200)

Set WShell = CreateObject ("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

ZTVersion = WShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\" &_
     "Installer\UserData\S-1-5-18\Products\039E7F954EF3FC948A929C8C5110AA8C\InstallProperties\DisplayVersion")
ZTEVVersion = WShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Ai Squared\ZoomText Xtra\10.1\Settings\Reader\Software Version")
mainline=False

Set WshSysEnv = WShell.Environment("PROCESS")
userprofile = WshSysEnv("USERPROFILE") 
ZTlogfilepath = userprofile & "\AppData\Roaming\Ai Squared\ZoomText 10.1\" & ZTVersion 

folder = ZTlogfilepath
ParseAllFilesInFolder Folder,strInfile,mainline
Folder=userprofile & "\AppData\Roaming\Ai Squared\ZoomText 10.1\" & ZTEVVersion

If objFSO.FolderExists(Folder) Then ParseAllFilesInFolder Folder,strInfile,mainline

Set objparselogfile1stream1 = Nothing
Set objparselogfile1 = Nothing
Set WshSysEnv = Nothing
Set objInstallLocationTextFile = Nothing
Set objInstallLocationTextStream = Nothing
'Set objlogTextFile = Nothing 
'Set objlogTextFileStream = Nothing 
Set WShell = Nothing
Set objFSO = Nothing
ejectProcess()

'WScript.Quit()

'----------------------------------------
Function ejectProcess()
'----------------------------------------
	
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

'-------------------------------------------
Function strComputerName()
'-------------------------------------------
	Set objNetwork  = CreateObject("WScript.Network")
	strComputerName = objNetwork.ComputerName
	
	set objNetwork = Nothing
End Function


'-----------------------------------------
 Function ParseAllFilesInFolder(Folder,strInfile,mainline)
'-----------------------------------------
	Set objlogTextFile = objFSO.GetFile ("d:\Utils\Results.txt")
	Set objlogTextFileStream = objlogTextFile.OpenAsTextStream(8,-1)  'for appending
	
	Set wshell = CreateObject ("WScript.Shell")
	Set FSO = CreateObject("Scripting.FileSystemObject")
	
	Set logFiles = FSO.GetFolder(Folder).Files	
	
'''''''''''''''''''''''''''Get color code ''''''''''''''''''''''''''
	Set objcolorTextFile = objFSO.GetFile ("d:\Utils\color.txt")
	Set objcolorTextFileStream = objcolorTextFile.OpenAsTextStream(1,0)  ' for Reading
	color = objcolorTextFileStream.Readline '("3399FF")
	WScript.Sleep (50)
	objcolorTextFileStream.close
	Set objcolorTextFile = Nothing 
	Set objcolorTextFileStream = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	i = 1
	For Each File In logFiles
	
		filename(i) = File.Name
		newfolder = folder & "\" & filename(i)
		'look for ERROR in file
		If InStr(newfolder,".xml")=0 Then 
			strInfile(i) = wshell.Run ("%comspec% /c c:\windows\system32\find.exe /c ""ERROR"" """ & newfolder& """",8 ,True)
			strInfile2(i) = wshell.Run ("%comspec% /c c:\windows\system32\find.exe /c ""FATAL"" """ & newfolder& """",8 ,True)
									
			If strInfile(i) = 0 Then 
				strToFind = "Level=""ERROR"""
				count = 0
				getErrorCountFunc newfolder,strToFind,Count
				'If mainline Then 
					GradeIt ="GoldenRod>IGNORE"
				'Else
				'	GradeIt ="Red>FAILED"
				'End If 
				objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>Error(s) level = ""ERROR"" was found in """& filename(i) & """ file. Error count = "&Count&"</font></th><th><font color="&GradeIt&"</font></th></tr>")
				'MsgBox ("File name """& filename(i) & """ contains errors") ' strInFile = 0 fatal was found, =1 no fatal errors were found
				flipcolor color
			ElseIf strInfile(i) = 1 Then  
				objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>No level = ""ERROR"" errors were found parsing log file """& filename(i) & """</font></th><th><font>PASSED</font></th></tr>")
				'MsgBox ("File name """& filename(i) & """ contains no errors")
				flipcolor color
			End If 
			
			If strInfile2(i) = 0 Then 
				strToFind = "Level=""FATAL"""
				count = 0
				getErrorCountFunc newfolder,strToFind,Count
				objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>Error(s) level = ""FATAL"" was found in """& filename(i) & """ file. Fatal error count = "&Count&"</font></th><th><font color=Red>FAILED</font></th></tr>")
				'MsgBox ("File name """& filename(i) & """ contains errors") ' strInFile = 0 fatal was found, =1 no fatal errors were found
				flipcolor color
			ElseIf strInfile2(i) = 1 Then  
				objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>No level = ""FATAL"" errors were found parsing log file """& filename(i) & """</font></th><th><font>PASSED</font></th></tr>")
				'MsgBox ("File name """& filename(i) & """ contains no errors")
				flipcolor color
			End If 
			
		End If 	
		i = i + 1
		If i > 200 Then Exit For 
	Next
	
'''''''''''''''''''''''''''Write color code ''''''''''''''''''''''''''
	Set objcolorTextFile = objFSO.GetFile ("d:\Utils\color.txt")
	Set objcolorTextFileStream = objcolorTextFile.OpenAsTextStream(2,0)  ' for writing
'	'color = objcolorTextFileStream.Readline '("3399FF")

'	If color = "3399FF" Then  
'   	color = "99CCFF"
'	Else 
'		color = "3399FF"
'	End If 
	objcolorTextFileStream.Writeline (color)
	WScript.Sleep (50)
	objcolorTextFileStream.close
	Set objcolorTextFile = Nothing 
	Set objcolorTextFileStream = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	Set logFiles = Nothing
	
	Set fso = Nothing
	Set wshell = Nothing 
	Set objlogTextFile = Nothing 
	Set objlogTextFileStream = Nothing
	
End Function

Function flipcolor(color)
	'''''''''' Flip color code ''''''''''''''''''''''''
			If color = "3399FF" Then 
			  	color = "99CCFF"
			Else
			 	 color = "3399FF"
			End If
	''''''''''''''''''''''''''''''''''''''''''''''''''''
End Function

Function getErrorCountFunc(newfolder,strToFind,Count)

	Set fso =CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(newfolder) Then 
	
		Set file = fso.OpenTextFile(newfolder,1,False,0)
	
		str =file.ReadAll
		WScript.Sleep(1000)
		out = str 
		CharacterCount = (Len(out) - Len(Replace(out, strToFind, "")))/Len(strToFind)
		
		Count = CharacterCount
		Set file = Nothing
	Else
	    Count = -1 ' FAILED - folder not found
	End If
	
	Set fso =Nothing
	
End Function