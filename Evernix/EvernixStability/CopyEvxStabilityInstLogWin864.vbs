On Error Resume Next 

Set WShell = CreateObject ("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Resultlogfile = "d:\utils\Results.txt"

buildfilepath = "\\ai2s0017\ts01\GROUP\BUILDSTORAGE\Images\Evernix\mainline\"
WinPath="c:\windows\"
Logg="Ai2Install_ZT10_1.log"
logg11="Ai2Install_ZT11_0.log"
InstallLog=WinPath & logg
InstallLog11=WinPath & logg11

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


InstallLocationFilePath = "q:\ATS\ATM\atm_resources\InstallLocationEvernix.txt"
Set objInstallLocationTextFile = objFSO.GetFile(InstallLocationFilePath)
Set objInstallLocationTextStream = objInstallLocationTextFile.OpenAsTextStream(1,-1)  'open for reading
buildpath = objInstallLocationTextStream.ReadLine

If Not objFSO.FolderExists(buildpath & "\Reports") Then 
	objFSO.CreateFolder (buildpath & "\Reports")
	WScript.Sleep(100)
	objFSO.CreateFolder (buildpath & "\Reports\QA")
Else 
	If Not objFSO.FolderExists(buildpath & "\Reports\QA") Then
		objFSO.CreateFolder (buildpath & "\Reports\QA")
	End If
End If 
If Not objFSO.FolderExists(buildpath & "\Reports\QA\Win8164") Then
	objFSO.CreateFolder (buildpath & "\Reports\QA\Win8164")
	WScript.Sleep(100)
End If

stryear = Year(Date)
strday = Day(Date)
strmonth = Month(Date)
thetime =Time
thetime= Replace(thetime," ","")
thetime= Replace(thetime,":",".")
If WScript.Arguments.Count=0 Then
  buildpath = buildpath & "\Reports\QA\Win8164\Stability" &"_"& stryear &"-"& strmonth &"-"& strday &"_"& thetime
Else
  buildpath = buildpath & "\Reports\QA\Win8164\"&WScript.Arguments.Item(0)&"_"& stryear &"-"& strmonth &"-"& strday &"_"& thetime
End If
objFSO.CreateFolder (buildpath)

WScript.Sleep(500)
If objfso.FileExists(InstallLog) Or objfso.FileExists(InstallLog11) Then
  Set files=objFSO.GetFolder(WinPath).Files
  For Each file In files
    if (instr(file.name,".log")>0) and (instr(Lcase(file.name),"ai2")>0 or instr(lcase(file.name),"zoomtext")>0) then
      objFSO.CopyFile file.Path, buildpath & "\", True
    End If
  Next
  WScript.Sleep(500)
Else
  objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>ZoomText install log was not found! " & buildpath & "</font></th><th><font color=Red>FALSE</font></th></tr>")
  If color = "FFFFFF" Then  
  	color = "99CCFF"
  Else 
	color = "FFFFFF"
  End If 
End If 		
 
If objFSo.FileExists (buildpath & "\" & logg) Or objFSo.FileExists (buildpath & "\" & logg11) Then 
		objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>Copying ZoomText install log to " & buildpath & "</font></th><th><font>TRUE</font></th></tr>")
		'objlogTextFileStream.writeline("<Action>Copying log files to " & logfilepath & " has PASSED</Action><br>" )
Else 
		objlogTextFileStream.writeline("<tr bgcolor = #" & color & " ALIGN=""LEFT""><th><font>Copying ZoomText install log to " & buildpath & "</font></th><th><font color=Red>FALSE</font></th></tr>")
		'objlogTextFileStream.writeline("<Action>Copying log files to " & logfilepath & " has FAILED</Action><br>" )
End If 
WScript.Sleep(5000)		

'Save folder name for later use
Set pathfile=objfso.GetFile("d:\utils\path.txt")
Set pathfiletext=pathfile.OpenAsTextStream(2,0)
pathfiletext.WriteLine(buildpath)
WScript.Sleep(500)
pathfiletext.Close
Set pathfile=Nothing
Set pathfiletext=Nothing
 
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


WScript.Quit()

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
