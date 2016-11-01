#region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseX64=n
#endregion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <Date.au3>

#cs ----------------------------------------------------------------------------

	1) Capture in execution log smi-colon delimited data
	- Date
	- Computer user
	- Computer name
	- OS ***[this needs to be updated to support windows 8]***
	- product version

	2) Merge TestSupp.ini and Update.ini ini's

	3) If nothing on the command line assume 10.0 installation otherwise
	determine zt version based on command-line

	4) this executable should be signed to avoid detection from virus software

#ce ----------------------------------------------------------------------------
Global $score, $lowest, $getUACStatus, $CPUName, $LCache, $CPUMan, $NumberOfCPUs, $VideoCard,$bFound,$ProgData,$bDisableCrashUpld
Global $DxDiag, $WDDM, $VideoRAM, $version, $DriverDate, $Res, $bits, $refresh, $Memory, $VideoDesc,$WinVersion, $szlocVer
Global $ZTEVVersion,$ZTPHVersion,$Versionstr,$DumpUploader,$ProcName,$strn,$str,$OSVersion,$SP,$OSBuild, $ztInstallDirS
Global $AwsCore, $AwsS3, $AwsCorePathfile, $AwsS3Pathfile

$rootDir = "\\ai2s-lab02\D\InternalPostInstall"
$DumpUploader = $rootDir&"\Ai2CrashDmpUploader\AiSquared.CrashDumpUploader.exe"
$AwsCore = "AWSSDK.Core.dll"
$AwsS3 = "AWSSDK.S3.dll"
$AwsCorePathfile = $rootDir&"\Ai2CrashDmpUploader\" & $AwsCore
$AwsS3Pathfile = $rootDir&"\Ai2CrashDmpUploader\" & $AwsS3
$EventsSubXML=$rootDir&"\eventsSubscription.xml"
$bFound=False
$bDisableCrashUpld=False

$I = $cmdline[0] ; get number of arguments -
;$ztInstallDir = $cmdline[1] - abandoned this method since it will not give me reliable results
;$ztVersion = $cmdline[2] - abandoned this method since it will not give me reliable results
$OSVersion = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProductName")
$SP = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "CSDVersion")
$OSBuild = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "CurrentBuildNumber")

if (StringInStr($OSVersion,"Windows 10")) Then
   $strn=RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion","CurrentMajorVersionNumber")

   $str=RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion","CurrentMinorVersionNumber")

   $WinVersion = $strn&"."&$str
EndIf

if @OSArch="x64" Then
  $UpdateDate = RegRead("HKLM64\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\Results\Install","LastSuccessTime")
  RegWrite("HKLM64\SOFTWARE\Microsoft\Windows\Windows Error Reporting\LocalDumps","DumpType", "REG_DWORD","2")
  RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\Windows Error Reporting\LocalDumps","DumpType", "REG_DWORD","2")
Else
  $UpdateDate = RegRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\Results\Install","LastSuccessTime")
  RegWrite("HKLM\SOFTWARE\Microsoft\Windows\Windows Error Reporting\LocalDumps","DumpType", "REG_DWORD","2")
EndIf
$UpdateDateNew=_DateTimeFormat ( $UpdateDate, 1 )
$UpdateDateNew2=_DateTimeFormat ( $UpdateDate, 3 )
$finalUpdateDate=$UpdateDateNew &" at "&$UpdateDateNew2



If StringInStr($OSVersion, "XP", 1) Then
	$score = "N/A"
	$lowest = "N/A"
	$getUACStatus = "N/A"
Else
	UACStatus()
	getwindowsexpindex()
	$score = Round($score, 1)
EndIf

getCPUIfno()
NumberOfCPUsfunc()
getVideoInfo()
GetDXInfo()

If $I = 0 Then ; no arguments -- assume Zt version 10.0

	$ZTEVVersion=RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Ai Squared\ZoomText Xtra\10.0\Settings\Reader", "Software Version")

	$logMsg = "Date = '" & _Now() & "; User " & @UserName & "; Computer " & @ComputerName & "; OS " & $OSVersion & " " & @OSArch & "; ZT version " & $ZTEVVersion & "; " & $SP & "; OS Build " & $OSBuild & "; Windows Version " & $WinVersion & "; System Memory " & $Memory & "MB; " & $getUACStatus & "; WEI " & $score & " for " & $lowest & "; Processor " & $CPUName & " " & $LCache & " " & $CPUMan & "; Number of CPUs " & $NumberOfCPUs & "; Video info " & $VideoCard & "; Video RAM " & $VideoRAM & "MB; Driver Version " & $version & "; Driver Date " & $DriverDate & "; Res " & $Res & "; Color depth " & $bits & "; Refresh rate " & $refresh &"; Last successful update was done on: "&$finalUpdateDate

	ConsoleWrite(@ScriptLineNumber & ": " & $logMsg & @CRLF)

	FileWriteLine($rootDir & "\Execution.log", $logMsg)

	$ztInstallDir = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Ai Squared\ZoomText Xtra\10.0\Settings\Reader", "Program Directory")

	If 0 = StringLen($ztInstallDir) Then Abort("Problem reading the registry")
	$ZTEVVersion=StringReplace($ZTEVVersion," ","")
	CopyConfig($ZTEVVersion)

	If FileExists($ztInstallDir&"\Aisquared.Magnification.ZoomText.exe") Then $ZTPHVersion=FileGetVersion($ztInstallDir&"\Aisquared.Magnification.ZoomText.exe")

	If $ZTPHVersion<>"" Then CopyConfig($ZTPHVersion)


Else ; arguments found

	$strArg = $CmdLineRaw ;take it all
	$strArgArr = StringSplit($strArg, """")

	$logMsg = "Date = '" & _Now() & "; User " & @UserName & "; Computer " & @ComputerName & "; OS " & $OSVersion & " " & @OSArch & "; ZT version " & $CmdLine[2] & "; " & $SP & "; OS Build " & $OSBuild & "; Windows Version " & $WinVersion & "; System Memory " & $Memory & "MB; " & $getUACStatus & "; WEI " & $score & " for " & $lowest & "; Processor " & $CPUName & " " & $LCache & " " & $CPUMan & "; Number of CPUs " & $NumberOfCPUs & "; DirectX Version: " & $DxDiag & "; WDDM Version: " & $WDDM & "; Video info " & $VideoDesc & " " & $VideoCard & "; Video RAM " & $VideoRAM & "MB; Driver Version " & $version & "; Driver Date " & $DriverDate & "; Res " & $Res & "; Color depth " & $bits & "; Refresh rate " & $refresh & "; Last successful update was done on: "&$finalUpdateDate

	ConsoleWrite(@ScriptLineNumber & ": " & $logMsg & @CRLF)

	FileWriteLine($rootDir & "\Execution.log", $logMsg)

	$ztInstallDir = $CmdLine[1]

	If 0 = StringLen($ztInstallDir) Then Abort("Problem - no ZT path argument was passed")
	$ZTEVVersion=StringReplace($strArgArr[2]," ","")
	CopyConfig($ZTEVVersion)

	If FileExists($ztInstallDir&"\Aisquared.Magnification.ZoomText.exe") Then $ZTPHVersion=FileGetVersion($ztInstallDir&"\Aisquared.Magnification.ZoomText.exe")

	If $ZTPHVersion<>"" Then CopyConfig($ZTPHVersion)

	If FileExists($EventsSubXML) Then FileCopy($EventsSubXML,$ztInstallDir&"\",1)

    If $I = 3 Then $bDisableCrashUpld = True

EndIf

If ShouldEnableIICE() Then SetIICEZtConfig()

If FileExists($ztInstallDir & "\TestSupp.ini") Then MergeIni($ztInstallDir, "TestSupp.ini") ; if file exists

If FileExists($ztInstallDir & "\") Then MergeIni($ztInstallDir, "Update.ini") ; if folder exists

If Not FileExists($ztInstallDir & "\AiSquared.CrashDumpUploader.exe") Or IsOlderCDVersionFound() Then
   ;Run($ztInstallDir & "\AiSquared.CrashDumpUploader.exe -au",$ztInstallDir, @SW_HIDE)
   FileCopy($DumpUploader, $ztInstallDir, 1)  ; 1 = overwrite, 8 = create path and copy FileChangeDir
   ;Run($ztInstallDir & "\AiSquared.CrashDumpUploader.exe -au",$ztInstallDir, @SW_HIDE)
EndIf

If Not FileExists($ztInstallDir & "\" & $AwsCore) Then
   FileCopy($AwsCorePathfile, $ztInstallDir, 0)
EndIf

If Not FileExists($ztInstallDir & "\" & $AwsS3) Then
   FileCopy($AwsS3Pathfile, $ztInstallDir, 0)
EndIf

If Not $bDisableCrashUpld Then
   If FileExists($ztInstallDir & "\AiSquared.CrashDumpUploader.exe") Then
	  Run("""" & $ztInstallDir & "\AiSquared.CrashDumpUploader.exe"" -au",$ztInstallDir, @SW_HIDE) ;Run(@ComSpec & " /k """ & $ztInstallDir & "\AiSquared.CrashDumpUploader.exe"" -au", "", @SW_SHOW)
   EndIf
Else
   If FileExists($ztInstallDir & "\AiSquared.CrashDumpUploader.exe") Then
	  Run("""" & $ztInstallDir & "\AiSquared.CrashDumpUploader.exe"" -u",$ztInstallDir, @SW_HIDE) ;Run(@ComSpec & " /k """ & $ztInstallDir & "\AiSquared.CrashDumpUploader.exe"" -au", "", @SW_SHOW)
   EndIf
EndIf
;MsgBox(0, "", "done")

Func MergeIni($ztInstallDirFunc, $ini)

	ConsoleWrite(@ScriptLineNumber & ": zt dir read from registry: " & $ztInstallDirFunc & @CRLF)

	$sectionNames = IniReadSectionNames($rootDir & "\" & $ini)

	If @error Then Abort("Error getting section names for " & $ini & " file")

	For $section = 1 To $sectionNames[0]

		$sectionData = IniReadSection($rootDir & "\" & $ini, $sectionNames[$section])

		If @error Then Abort("Error getting data from " & $sectionNames[$section] & " for " & $ini & " file")

		For $data = 1 To $sectionData[0][0]

			$key = $sectionData[$data][0]

			$value = $sectionData[$data][1]

			ConsoleWrite(@ScriptLineNumber & ": attempting to update " & $ini & " ['" & $sectionNames[$section] & "'] with" & "Key: " & $key & " Value: " & $value & @CRLF)

			If (1 <> IniWrite($ztInstallDirFunc & "\" & $ini, $sectionNames[$section], $key, $value)) Then Abort("Unable to update '" & $ini & "'")

			ConsoleWrite(@ScriptLineNumber & ": " & "success" & @CRLF)

		Next

	Next

EndFunc   ;==>MergeIni

Func Abort($logMsg)

	ConsoleWrite(@ScriptLineNumber & ": " & $logMsg & @CRLF)

	FileWriteLine($rootDir & "\Execution.log", $logMsg)

	;MsgBox(0, "ABORT", $logMsg)

	Exit

EndFunc   ;==>Abort

Func getwindowsexpindex()

	$objWMIservices = ObjGet("winmgmts:" & "{impersonationLevel=impersonate}!\\" & @ComputerName & "\root\cimv2")
	$colWSA = $objWMIservices.ExecQuery("Select * From Win32_WinSAT")


	For $objItem In $colWSA

		$score1 = $objItem.cpuscore
		$score2 = $objItem.MemoryScore
		$score3 = $objItem.GraphicsScore
		$score4 = $objItem.D3DScore
		$score5 = $objItem.DiskScore

	Next

	If $score1 > $score2 Then
		$score = $score2
		$lowest = "Memory"
	Else
		$score = $score1
		$lowest = "Processor"
	EndIf
	If $score > $score3 Then
		$score = $score3
		$lowest = "Graphics"
	EndIf
	If $score > $score4 Then
		$score = $score4
		$lowest = "Gaming graphics"
	EndIf
	If $score > $score5 Then
		$score = $score5
		$lowest = "Primary hard disk "
	EndIf

EndFunc   ;==>getwindowsexpindex

Func UACStatus()

	$flag = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "EnableLUA")

	$flag2 = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "ConsentPromptBehaviorAdmin")
	$os = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProductName")

	If StringInStr($os, "Vista", 1) Then
		If $flag = 0 Then
			$getUACStatus = "UAC is OFF"
		EndIf
		If $flag = 1 Then
			$getUACStatus = "UAC is ON"
		EndIf
	EndIf

	If StringInStr($os, "Windows 7", 1) Or StringInStr($os, "Windows 8", 1) Then
		If $flag2 = 0 Then
			$getUACStatus = "UAC is set to ""Never notify"""
		EndIf
		If $flag2 = 5 Then
			$getUACStatus = "UAC is set to medium"
		EndIf
		If $flag2 = 2 Then
			$getUACStatus = "UAC is set to ""Always notify"""
		EndIf
	EndIf
	;MsgBox (0,"",$flag &@CRLF & $flag2& @CRLF& $os&@CRLF&$getUACStatus)
EndFunc   ;==>UACStatus

Func getCPUIfno()

	$objWMIService = ObjGet("winmgmts:\\" & @ComputerName & "\root\cimv2")
	$colItems = $objWMIService.ExecQuery("Select * from Win32_Processor")
	For $objItem In $colItems

		$LCache = $objItem.L2CacheSize
		$CPUName = $objItem.Name
		$CPUMan = $objItem.Manufacturer
	Next

EndFunc   ;==>getCPUIfno

Func getVideoInfo()

	$objWMIService = ObjGet("winmgmts:\\" & @ComputerName & "\root\cimv2")
	$colItems = $objWMIService.ExecQuery("Select * from Win32_VideoController")

	For $objItem In $colItems
		$VideoDesc = $objItem.Description
		$VideoCard = $objItem.Caption
		$VideoRAM = $objItem.AdapterRAM ;/1024 & " MB"
		$version = $objItem.DriverVersion
		$DriverDate = $objItem.DriverDate
		$hres = $objItem.CurrentHorizontalResolution
		$vres = $objItem.CurrentVerticalResolution
		$bits = $objItem.CurrentBitsPerPixel
		$refresh = $objItem.CurrentRefreshRate
		If $hres <> "" And $vres <> "" And $bits <> "" And $refresh <> "" Then ExitLoop
	Next
	;$driverdate1= CStr(DriverDate)
	$driverdate1 = StringLeft($DriverDate, 8)
	$driverdate2 = StringMid($driverdate1, 5, 2) & "-" & StringRight($driverdate1, 2) & "-" & StringLeft($driverdate1, 4)

	$DriverDate = $driverdate2
	$Res = $hres & "x" & $vres
	$VideoRAM = Round($VideoRAM / 1024000)

EndFunc   ;==>getVideoInfo

Func NumberOfCPUsfunc()

	$objWMIService = ObjGet("winmgmts:\\" & @ComputerName & "\root\CIMV2")
	$colItems = $objWMIService.ExecQuery("SELECT * FROM Win32_PerfFormattedData_PerfOS_Processor WHERE Name = '_Total'")

	$colProcs = $objWMIService.ExecQuery("Select * from Win32_ComputerSystem")

	For $objItem In $colProcs
		$NumberOfCPUs = $objItem.NumberOfLogicalProcessors
		$Memory = $objItem.TotalPhysicalMemory
	Next
	$Memory = $Memory / 1024000
	$Memory = Round($Memory)

EndFunc   ;==>NumberOfCPUsfunc

Func GetDXInfo()

	$sh = ObjCreate("Wscript.shell")
	$fso = ObjCreate("scripting.filesystemobject")

	$windowsdir = $sh.ExpandEnvironmentStrings("%windir%")
	$shSysEnv = $sh.Environment("PROCESS")
	$userprofile = $shSysEnv("USERPROFILE")

	$tmpFile = $userprofile & "\DxDiag.txt"
	If $fso.FileExists($tmpFile) Then $fso.DeleteFile($tmpFile, True)
	$cmd = $windowsdir & "\System32\DxDiag.exe /t " & $tmpFile

	$sh.Run($cmd, 4, True)
	$found = False
	$J = 1

	Do
		$J = $J + 1
		Sleep(1000)
		If $J >= 60 Then ExitLoop
		If $fso.FileExists($tmpFile) Then $found = True

	Until $found

	$file = $fso.OpenTextFile($tmpFile, 1, False, 0)

	$allofit = $file.ReadAll
	Sleep(500)
	If StringInStr($allofit, "DirectX Version: ") Then
		$number = StringInStr($allofit, "DirectX Version: ")
		$DxDiag = StringMid($allofit, $number + 17, 10)
	Else
		$DxDiag = "N/A"
	EndIf

	If StringInStr($allofit, "Driver Model: ") Then
		$number2 = StringInStr($allofit, "Driver Model: ")
		$WDDM = StringMid($allofit, $number2 + 14, 8)
	Else
		$WDDM = "N/A"
	EndIf

EndFunc   ;==>GetDXInfo

Func CopyConfig($Versionstr)

   if StringInStr($Versionstr,".") Then
    $szHolder = StringSplit($Versionstr,".")
	$szMajor = $szHolder[1]
	$szMinor = $szHolder[2]
	if ($szMajor = "11") Then
	   $szlocVer = $szMajor
    else
	   $szlocVer = $szMajor & "." & $szMinor
    EndIf
   EndIf

	DirCreate(@AppDataDir&"\Ai Squared\ZoomText " & $szlocVer & "\"&$Versionstr&"\LoggerConfig")
	Sleep(50)
	FileCopy($rootDir&"\AiSquared.Logging.Logger.dll.config",@AppDataDir&"\Ai Squared\ZoomText " & $szlocVer & "\"&$Versionstr&"\LoggerConfig\",1)

EndFunc

Func StopProcess($ProcName)

   If ProcessExists($ProcName) Then ProcessClose($ProcName)

EndFunc  ;==>StopProcess

Func ShouldEnableIICE()

   Local $sz_User = @UserName
   Local $sz_MachineName = @ComputerName
   Local $sz_BannedUsers_List = "yfang ishchetinin jeckhardt igor.shchetinin ishch yurec ysimanovski"
   Local $sz_BannedComputers_List = "AI2W0065 AI2WD017WIN8X64 AI2W0034 HRK1-DHP-F30863 WIN8X64 WIN-FCLR8PBHR9A DESKTOP-3UOKUCU AI2WD0070 AI2W0069" & _
   "DESKTOP-J9FCDPR WIN-C6R3UPS8DO3 AI2W0069W10 DESKTOP-FFK72NP DESKTOP-HC2NUNU DESKTOP-U9TBER3 DESKTOP-4B9E983 DESKTOP-R7IKR6F DESKTOP-85C460T" & _
   "AI2D0099-WIN10 AI2D0099 WIN-T3M6K40M6JF AI2L0022W10X64 WIN-T3M6K40M6JF"

   If StringInStr($sz_BannedUsers_List, $sz_User) <> 0 Then
	  Return False
   EndIf
   If StringInStr($sz_BannedComputers_List, $sz_MachineName) <> 0 Then
	  Return False
   EndIf

   Return True

EndFunc

Func SetIICEZtConfig()

   Local $docpath = $ztInstallDir & "\ZoomTextConfig.xml"
   Local $fso, $root, $root2

   $fso=ObjCreate("Scripting.FileSystemObject")
   If $fso.FileExists($docpath) Then
	  $xml=$fso.OpenTextFile($docpath,1)
	  $root = $xml.ReadAll
	  $xml.Close

	  $root2 = StringReplace($root,"<Property name=""IICE"" value=""false"" />","<Property name=""IICE"" value=""true"" />")

	  $xml=$fso.OpenTextFile($docpath,2)
	  $xml.Write($root2)
	  Sleep(1000)
	  $xml.Close
   EndIf

EndFunc

Func IsOlderCDVersionFound()

   $InstalledCDVer = FileGetVersion ($ztInstallDir & "\AiSquared.CrashDumpUploader.exe")

   $strIarryver = StringSplit($InstalledCDVer,".")

   $StockCDVer = FileGetVersion ($DumpUploader)

   $strSarryver = StringSplit($StockCDVer,".")

   For $i =1 To 4

      If (Int($strIarryver[$i]) < Int($strSarryver[$i])) Then
	     Return True
      EndIf

   Next

   Return False

EndFunc