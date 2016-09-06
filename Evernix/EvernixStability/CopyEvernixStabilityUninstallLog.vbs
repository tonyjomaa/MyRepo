On Error Resume Next 

Set WShell = CreateObject ("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Resultlogfile = "d:\utils\Results.txt"

buildfilepath = "\\ai2s0017\ts01\GROUP\BUILDSTORAGE\Images\Evernix\mainline\"
WinPath="c:\windows\"
Logg="Ai2Install_ZT11_0.log"
InstallLog=WinPath & logg
If Not objfso.FileExists(InstallLog) Then InstallLog=WinPath & "Ai2Install_ZT10_1.log"

'Get folder name for later use
Set pathfile=objfso.GetFile("d:\utils\path.txt")
Set pathfiletext=pathfile.OpenAsTextStream(1,0)
buildpath = pathfiletext.ReadLine
WScript.Sleep(500)
pathfiletext.Close
Set pathfile=Nothing
Set pathfiletext=Nothing

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
Set objlogTextFileStream = objlogTextFile.OpenAsTextStream(8,-1) 'For appending

'InstallLocationFilePath = "q:\ATS\ATM\atm_resources\InstallLocationEverest.txt"
'Set objInstallLocationTextFile = objFSO.GetFile(InstallLocationFilePath)
'Set objInstallLocationTextStream = objInstallLocationTextFile.OpenAsTextStream(1,-1)  'open for reading
'buildpath = objInstallLocationTextStream.ReadLine

'
WScript.Sleep(500)
Set files=objFSO.GetFolder(WinPath).Files
For Each file In files
  if (instr(file.name,".log")>0) and (instr(Lcase(file.name),"ai2")>0 or instr(lcase(file.name),"zoomtext")>0) then
    If InStr(file.Path,logg)>0 Then
      objFSO.CopyFile file.Path, buildpath & "\UnInstallLog.log", True
    Else
      objFSO.CopyFile file.Path, buildpath & "\", True
    End If 
  End If
Next

 
If objFSo.FileExists (buildpath & "\UnInstallLog.log") Then 
		objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>Copying ZoomText uninstall log to " & buildpath & "</font></th><th><font>TRUE</font></th></tr>")
		'objlogTextFileStream.writeline("<Action>Copying log files to " & logfilepath & " has PASSED</Action><br>" )
Else 
		objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>Copying ZoomText uninstall log to " & buildpath & "</font></th><th><font color=Red>FALSE</font></th></tr>")
		'objlogTextFileStream.writeline("<Action>Copying log files to " & logfilepath & " has FAILED</Action><br>" )
End If 
WScript.Sleep(5000)		

 
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
ejectProcess()
Set WshSysEnv = Nothing
Set objInstallLocationTextFile = Nothing
Set objInstallLocationTextStream = Nothing
Set objlogTextFile = Nothing 
Set objlogTextFileStream = Nothing 
Set WShell = Nothing
Set objFSO = Nothing


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

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function GetCurrentPathFunction(ZTVersion,buildpath,logfilepath)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Set objFSO1 = CreateObject ("Scripting.FileSystemObject")

	Set objlogTextFile = objFSO1.GetFolder(buildpath).SubFolders
		
	For Each Folder In objlogTextFile
	
		foldername = Folder.Name
		If InStr(foldername,ZTVersion) > 0 Then  
			'MsgBox foldername
			logfilepath = buildfilepath & foldername & "\Reports\"
			Exit For 
		End If 
		'MsgBox logfilepath
	Next 
	Set objlogTextFile = Nothing 
	Set objFSO1 = Nothing 
	
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
