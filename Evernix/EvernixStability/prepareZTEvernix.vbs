On Error Resume Next 

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set WShell = CreateObject ("WScript.Shell")

Testtype = WScript.Arguments.Item(0)
If Testtype = "" Then 
 	Testtype = "Evernix Test"
	Testperformed = "a quick smoke test"
	Smoketest = " Smoke"
Else 
 	Testperformed = "an automated test"
 	Smoketest = ""
End If 


'''''''' get build path to extract version '''''
InstallLocationFilePath = "q:\ATS\ATM\atm_resources\installLocationEvernix.txt"
Set objInstallLocationTextFile = objFSO.GetFile(InstallLocationFilePath)
Set objInstallLocationTextStream = objInstallLocationTextFile.OpenAsTextStream(1,-1)  'open for reading
buildpath = objInstallLocationTextStream.ReadLine
WScript.Sleep (50)

strVerArry = Split(buildpath,"\")
strVer = strVerArry(UBound(strVerArry))

'''''''''''''clear out old log files '''''''''''
'Set WshSysEnv = WShell.Environment("PROCESS")
'userprofile = WshSysEnv("USERPROFILE") 
'ZTlogfilepath = userprofile & "\AppData\Roaming\Ai Squared" 
'objFSO.DeleteFolder ZTlogfilepath
'WScript.Sleep(2000)

'''''''''''''clear out installed old ZT folder/files '''''''''''

'If InStr(Testtype,"DocRdr-In-Notepad-BBox-Test") Then

	'' do nothing
'	WScript.Sleep (50)

'Else 
		OldZTFolder = "c:\Program Files\ZoomText 10"
	If objFSO.FolderExists (OldZTFolder) Then  
		objFSO.DeleteFolder OldZTFolder, True 
	End If 
'End If 

'''''''''''''' Freshen the variables ''''''''''''''''''''''''
score=0
lowest=""
CPUName = ""
LCache = 0
CPUMan = ""
VideoCard = ""
VideoRAM = 0
version = ""
DriverDate = ""
Res = ""
bits = ""
refresh = ""
WDDM = ""
DxDiag = ""
getUACStatus = ""
dpi = ""
DWM = ""
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

starttime = Time 
startdate = Date 


Set objlogTextFile = objFSO.CreateTextFile("d:\Utils\Results.txt")
Set objlogTextFile = objFSO.GetFile ("d:\Utils\Results.txt")
Set objlogTextFileStream = objlogTextFile.OpenAsTextStream(2,-1)  ' for writing
os = wshell.RegRead ("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProductName")
SP = wshell.RegRead ("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\CSDVersion")
OSBuild = wshell.RegRead ("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\CurrentBuildNumber")
strOSVer=objFSO.GetFileVersion("C:\Windows\System32\kernel32.dll")
strtemp=Split(strOSVer,".")
WinVersion = strtemp(0) & "." & strtemp(1)
If InStr(WinVersion,"6.1") Then WinVersion = "Windows 7"
If InStr(WinVersion,"6.2") Then WinVersion = "Windows 8.0"
If InStr(WinVersion,"6.3") Then WinVersion = "Windows 8.1"
'WinVersion = wshell.RegRead ("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\CurrentVersion")

'test = wshell.RegRead ("HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\(Default)")
'MsgBox (test)
strComputerName()

'''''''''''''''''''''''''''Set color code ''''''''''''''''''''''''''
Set objcolorTextFile = objFSO.GetFile ("d:\Utils\color.txt")
Set objcolorTextFileStream = objcolorTextFile.OpenAsTextStream(2,0)  ' for writing
objcolorTextFileStream.writeline ("FFFFFF")
WScript.Sleep (50)
Set objcolorTextFile = Nothing 
Set objcolorTextFileStream = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'*******  Read system memory   *****************************
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputerName & "\root\cimv2")
Set colSettings = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
For Each objComputer in colSettings
getTotalMemory = objComputer.TotalPhysicalMemory
Next
getTotalMemory = getTotalMemory /1000000
getTotalMemory = Left (getTotalMemory,4)
'MsgBox (getTotalMemory & "MB" )
getTotalMemory = getTotalMemory & "MB"
'  ***************************************************************
if instr(os,"XP") > 0 Or instr(os,"8.1") > 0 Or instr(os,"10") > 0 Then
score ="N/A"
lowest = "N/A"
getUACStatus ="N/A"
else
getwindowsexpindex score,lowest
UACStatus getUACStatus
end if

getCPUIfno CPUName,LCache,CPUMan
NumberOfCPUs
'getVideoInfo VideoCard,VideoRAM
'MoreVideoInfo version,DriverDate,Res,bits,refresh
GetDXInfo WDDM,DxDiag,VideoCard,VideoRAM,version,DriverDate,Res,dpi
CallAeroCheckFunction DWM
strUpdate=""
getUpdateDate strUpdate
 
If Is64Bit Then   ' get OS bits function call
	os =  os & " 64 bit"
Else
	os =  os & " 32 bit"
End If 
'MsgBox (os)

'  ************* This section is just to make it look pretty   ****************
'strComputerName()
taskfile = "q:\ATS\ATM\atm_servers\" & strComputerName & "\Init\CurrentTask.txt"
Set taskid = objFSO.GetFile (taskfile)
Set taskidstream = taskid.OpenAsTextStream (1,0)
 Do While Not taskidstream.AtEndOfStream
  id = taskidstream.ReadLine
 Loop 
  id = Int(id) 
   
    color = "3399FF" 
    id = id Mod 2
    If id = 1 Then 
    	color = "99CCFF"
    End If 
  Set taskidstream = Nothing 
 '********************************************************************************

TestType = Replace(Testtype,"-"," ")
TestType = Replace(Testtype,"_"," ")
objlogTextFileStream.writeline ("<!DOCTYPE HTML PUBLIC -//W3C//DTD HTML 3.2 Final//EN>")
objlogTextFileStream.writeline ("<html><body><p><b><font size = 3 color = Blue face= Verdana >Results for " & Testtype & "</font></b></p>")
objlogTextFileStream.writeline ("<b><font size = 2 color = Gray face= Verdana > Test Description = Perform " & Testperformed & "</font></b></p>")
objlogTextFileStream.writeline ("<table border =  0 cellspacing = 2 cellpading = 1 bordercolor = #D3D3D3 bgcolor = #FFFFFF width = 85% ALIGN=""CENTER""><tr bgcolor = #827B60 >")
objlogTextFileStream.writeline ("<th><font color = White >Run Time Status</font></th><th><font color = White>Description</font></th></tr><tr bgcolor = #ECE5B6><th><font color = Blue>Test Argument</font></th><th><font color = Blue>ZoomText " & Smoketest & " Test</font></th></tr>")
objlogTextFileStream.writeline ("<tr bgcolor = #C9C299><TD WIDTH=""25%"" Align=""Center""><font color = Blue Size=3><B>Build Version From Build Path</B></font></TD><TD WIDTH=""75%"" Align=""Center""><font color = Blue Size=3><B>" & strver & "</B></font></TD></tr>")
objlogTextFileStream.writeline ("<tr bgcolor = #ECE5B6><th><font color = Blue>Operating System</font></th><th><font color = Blue>" & os & " "&SP&", Build "&OSBuild&", Windows Version "&WinVersion&"; " & getTotalMemory & " of RAM ( UAC Status = "&getUACStatus&" ) </font></th></tr>") '<tr bgcolor = #C9C299><th><font color = Blue>Test Started</font></th>")
objlogTextFileStream.writeline ("<tr bgcolor = #C9C299><th><font color = Blue>Processor Specification</font></th><th><font color = Blue>" & CPUMan & " " & CPUName & " L2 Cache Size " & LCache & "; Number of Processors is "&NumberOfCPUs&"</font></th></tr>") 
objlogTextFileStream.writeline ("<tr bgcolor = #ECE5B6><th><font color = Blue>Video Graphics Information</font></th><th><font color = Blue>" & VideoCard & "; "&dpi&"; " & VideoRAM & "; "&version&"; "&DriverDate&"; "&Res&"; AeroGlass Status = "&DWM&"</font></th></tr>") 
objlogTextFileStream.writeline ("<tr bgcolor = #C9C299><th><font color = Blue>Windows Experience Index based on </font><font color = DarkBlue>" & lowest & "</font><font color = Blue> being the lowest score</font></th><th><font color = Blue>" & score & " </font></th></tr>") 
objlogTextFileStream.writeline ("<tr bgcolor = #ECE5B6><th><font color = Blue>Last time Windows Updates were successfully installed was on</font></th><th><font color = Blue>" & strUpdate & " </font></th></tr>") 
objlogTextFileStream.writeline ("<tr bgcolor = #C9C299><th><font color = Blue>WDDM version</font></th><th><font color = Blue>" & WDDM & " </font></th></tr>") 
objlogTextFileStream.writeline ("<tr bgcolor = #ECE5B6><th><font color = Blue>DirectX version</font></th><th><font color = Blue>" & DxDiag & " </font></th></tr>") 
objlogTextFileStream.writeline ("<tr bgcolor = #C9C299><th><font color = Blue>Test Machine Name</font></th><th><font color = Blue>" & strComputerName & " </font></th></tr>") 
objlogTextFileStream.writeline ("<tr bgcolor = #ECE5B6><th><font color = Blue>Test started at</font></th><th><font color = Blue>" & starttime & " " & startdate & "</font></th></tr>") '<tr bgcolor = #C9C299><th><font color = Blue>Test Started</font></th>")

If WScript.Arguments.Count = 2 Then 
If Testtype = "AppRdr In IE9 BBox Test" Or Testtype = "AppRdr Speed Analysis Test" Or Testtype = "DocRdr In IE9 BBox Test" Or Testtype = "DocRdr Speed Analysis Test" Or Testtype = "AppRdr In Firefox BBox Test"Then
	'strURLFile = "\\ai2s-lab02\D\SharePoint\URLList.txt"
	strURLFile = WScript.Arguments.Item(1)
	If objFSO.FileExists(strURLFile) Then
		objFSO.CopyFile strURLFile,"d:\qalab\ATS\URLList.txt",True 
		Testperformed = "an automated test" & vbCrLf& "URL list file: " & strURLFile
		Set openfile = objFSO.OpenTextFile("d:\qalab\ATS\URLList.txt",1,False,0)
		bgcolor = "ECE5B6"
		objlogTextFileStream.writeline ("<tr bgcolor = #"&bgcolor&"><th><font color = Blue>URL List File and Location  </font></th><th><font color = Blue><a href="&strURLFile&">"&strURLFile&"</a></font></th></tr>")
		FlipColor bgcolor
		Do While Not openfile.AtEndOfStream
		  line = openfile.ReadLine
		  If line <>"" And line <> " " Then
		   
		   objlogTextFileStream.writeline ("<tr bgcolor = #"&bgcolor&"><th><font color = Blue>Site URL  </font></th><th><font color = Blue><a href="&line&">"&line&"</a></font></th></tr>")
		   FlipColor bgcolor
		  End If
		Loop
		'strURLFile = "d:\qalab\ATS\URLList.txt"
	End If
End If
End If


'	objlogTextFileStream.writeline("<DESC>ZT10 Smoke Test</DESC><br><br>")
'	Get this info:
	'objlogTextFileStream.writeline ("<OS>Vista32 ( UAC = OFF System RAM = 3566.36 MB AeroGlass = ON )</OS>")
	'objlogTextFileStream.writeline ("<OS>OS: Vista32</OS><br>")
	'objlogTextFileStream.writeline	("<TEST_DESCRIPTION>Install ZoomText10</TEST_DESCRIPTION><br>")
	
ejectProcess()

Set objlogTextFile = Nothing 
Set objlogTextFileStream = Nothing 
Set objInstallLocationTextStream = Nothing
Set objInstallLocationTextFile = Nothing
Set WShell = Nothing
Set objFSO = Nothing
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

   Set objFSO        = Nothing
End Function

Function strComputerName()

	Set objNetwork  = CreateObject("WScript.Network")
	strComputerName = objNetwork.ComputerName
	
	set objNetwork = Nothing

End Function

Function getwindowsexpindex(score,lowest)

Set objWMIservices = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputerName & "\root\cimv2")
Set colWSA = objWMIservices.ExecQuery("Select * From Win32_WinSAT")


For Each objItem in colWSA

	score1 = objitem.cpuscore
	score2 = objItem.MemoryScore
	score3 = objItem.GraphicsScore
	score4 = objItem.D3DScore
	score5 = objItem.DiskScore

Next

If score1 > score2 Then 
	score = score2
	lowest = "Memory"
Else 
	score = score1
	lowest = "Processor"
End If 
If score > score3 Then 
	score = score3
	lowest = "Graphics"
End If 
If score > score4 Then 
	score = score4
	lowest = "Gaming graphics"
End If 
If score > score5 Then 
	score = score5
	lowest = "Primary hard disk "
End If 
'MsgBox lowest & score
Set colWSA = Nothing 
Set objWMIservices = Nothing 

End Function 

Function getCPUIfno(CPUName,LCache,CPUMan)
	strComputerName()
	Set objWMIService = GetObject("winmgmts:\\" & strComputerName & "\root\cimv2")
	Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor")
	For Each objItem in colItems
		 
		LCache = objItem.L2CacheSize  	'"L2 Cache Size"
		CPUName = objItem.Name			' "CPU name"
		CPUMan = objItem.Manufacturer	' "Manufacturer"
	Next 
	
	Set colItems = Nothing 
	Set objWMIService = Nothing 
	
End Function 

Function getVideoInfo (VideoCard,VideoRAM)

	strComputerName()

	Set objWMIService = GetObject("winmgmts:\\" & strComputerName & "\root\cimv2")
	Set colItems = objWMIService.ExecQuery("Select * from Win32_VideoController",,48)

	For Each objItem in colItems    
	
		VideoCard = objItem.Caption  ' "Video Card Name" it finds all but ends by the last one listed in Device Manager
		VideoRAM = FormatNumber(objItem.AdapterRAM /1024\1024, 0) & " MB" ' "Adapter RAM"
	Next 
	Set colItems = Nothing
	Set objWMIService = Nothing 
	
End Function 

Function Is64Bit

  Is64Bit = False 

  Set WshShell = WScript.CreateObject("WScript.Shell")
  Set WshSysEnv = WshShell.Environment("SYSTEM")
  If "AMD64" = WshSysEnv("PROCESSOR_ARCHITECTURE") Then
    Is64Bit = True 
  End If 

End Function

Function UACStatus(getUACStatus)

	Set wshell2 = CreateObject ("Wscript.shell")
	flag = wshell2.RegRead ("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\EnableLUA")
	'MsgBox flag
	flag2 = wshell2.RegRead ("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\ConsentPromptBehaviorAdmin")
	os = wshell2.RegRead ("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProductName")
	
	'MsgBox os
	If InStr(os,"Vista")>0 Then 
	 	If flag =0 Then
	 		getUACStatus = "UAC is OFF"
	 	End If 
	 	If flag =1 Then 
	 		getUACStatus = "UAC is ON"
	 	End If 	
	End If 
	
	If InStr(os,"Windows 7")>0 Or InStr(os,"Windows 8")>0 Or InStr(os,"Windows 10")>0 Then
		If flag2 = 0 Then 
		getUACStatus = "UAC is set to ""Never notify"""
		End If 
		If flag2 = 5 Then 
		getUACStatus = "UAC is set to medium"
		End If 
		If flag2 = 2 Then 
		getUACStatus = "UAC is set to ""Always notify"""
		End If
	End If 
	Set wshell2 = Nothing  

End Function 

Function NumberOfCPUs

	strComputerName()
	Set objWMIService = GetObject("winmgmts:\\" & strComputerName & "\root\CIMV2") 
	Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_PerfFormattedData_PerfOS_Processor WHERE Name = '_Total'") 
		
	Set colProcs = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
	
	For Each objItem in colProcs
		NumberOfCPUs = objItem.NumberOfLogicalProcessors
	Next
	
	Set objWMIService = Nothing 
	Set colItems = Nothing
	Set colProcs = Nothing
	
End Function 

Function MoreVideoInfo(version,DriverDate,Res,bits,refresh)

	strComputerName()
	
	Set objWMIService = GetObject("winmgmts:\\" & strComputerName & "\root\cimv2")
	Set colItems = objWMIService.ExecQuery("Select * from Win32_VideoController",,48)
	
	For Each objItem in colItems    
		
		'VideoCard = objItem.Caption  ' "Video Card Name" it finds all but ends by the last one listed in Device Manager
		'VideoRAM = FormatNumber(objItem.AdapterRAM /1024\1024, 0) & " MB" ' "Adapter RAM"
		version = objItem.DriverVersion
		DriverDate = objItem.DriverDate
		hres = objItem.CurrentHorizontalResolution
		vres = objItem.CurrentVerticalResolution
		bits = objItem.CurrentBitsPerPixel
		refresh = objItem.CurrentRefreshRate
		If hres<>"" And vres<>"" And bits<>"" And refresh<>"" Then Exit For
	Next 
	driverdate1= CStr(DriverDate)
	driverdate1= Left(DriverDate1,8)
	driverdate2= Mid(DriverDate1,5,2)&"-"&Right(DriverDate1,2)&"-"& Left(DriverDate1,4)
	
	DriverDate = DriverDate2
	Res = hres &"x"&vres
	
	Set objWMIService = Nothing 
	Set colItems = Nothing
	Set colProcs = Nothing
	
End Function

Function FlipColor(bgcolor)

	If bgcolor = "ECE5B6" Then
		 bgcolor = "C9C299"
	Else
		 bgcolor = "ECE5B6"
	End If 

End Function

Function GetDXInfo(WDDM,DxDiag,VideoCard,VideoRAM,version,DriverDate,Res,dpi)

	Set sh = CreateObject("Wscript.shell")
    Set  fso = CreateObject("scripting.filesystemobject")

	windowsdir = sh.ExpandEnvironmentStrings("%windir%")
    Set shSysEnv = Sh.Environment("PROCESS")
	userprofile = shSysEnv("USERPROFILE")

	tmpFile = userprofile &"\DxDiag.txt"
	If fso.FileExists(tmpFile) Then fso.DeleteFile tmpFile,True
    cmd = windowsdir & "\System32\DxDiag.exe /t " & tmpFile

	sh.Run cmd,4,True
	found = False
	J=1

	Do Until found
		J = J +1
	WScript.Sleep(1000)
		If J >= 60 Then Exit Do
		If fso.FileExists(tmpFile) Then found =True

	Loop 

	Set file = fso.OpenTextFile(tmpFile,1,False,0)

   Do While Not file.AtEndOfStream
   
     line = file.ReadLine
     
       If InStr(line,"DirectX Version: ")>0 Then DxDiag = line
       If InStr(line,"Driver Model: ")>0 Then WDDM = line
       If InStr(line,"Card name: ")>0 Then VideoCard = line
       If InStr(line,"Dedicated Memory: ")>0 Then VideoRAM = line
       If version ="" Then 
         If InStr(line,"Driver File Version: ")>0 Then version = line
       End if
       If InStr(line,"Driver Date/Size: ")>0 Then DriverDate = line
       If InStr(line,"Current Mode: ")>0 Then res = line
       If InStr(line,"System DPI Setting: ")>0 Then dpi = line
       
       If DxDiag<>"" And WDDM<>"" And VideoCard<>"" And VideoRAM<>"" And version<>"" And DriverDate<>"" And res<>"" And dpi<>"" Then Exit Do
  Loop
  VideoRAM = Replace(VideoRAM,"Dedicated","Video")
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function CallAeroCheckFunction(DWM)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	Set WShell2 = CreateObject ("WScript.Shell")

	flag = WShell2.Run ("q:\ATS\AutoTasks\AeroCheck.exe", 4, True)

	If flag = 0 Then 
	 
		DWM="ON"
	Else 
	
		DWM="OFF" 
	End If 
	Set WShell2 = Nothing 
	
End Function 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function getUpdateDate(strUpdate)

  Set sh=CreateObject("WScript.Shell")

  str=sh.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\Results\Install\LastSuccessTime")
strUpdate= MonthName(Month(str))&" "&Day(str)&", "&Year(str)&" at "&Hour(str)&":"&Minute(str)
Set sh=nothing
End Function
