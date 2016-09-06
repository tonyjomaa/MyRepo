' ********************************************************************************
'
' Copyright (c) 2014, Algorithmic Implementations, Inc. (dba Ai Squared)
' All Rights Reserved.
'
' Developer:     Tony Jomaa
' 
' Description:   ManageSQLInsert.vbs queues up DB entries
'
' Parameters:    
'
' Keywords:      
'                
' Notes:         DB serializer   
'
' ********************************************************************************

On Error Resume Next 
path="D:\ShareThis\ManageInsert.txt"
Set WShell = CreateObject ("WScript.Shell")
linenum=1
oldlinenum=1
Set fso= CreateObject("Scripting.FileSystemObject")

If WScript.Arguments.Count=0 Then
   thenow=Now
   fso.CopyFile path,"D:\ShareThis\ManageInsert"&thenow&".txt",True
   Set file=fso.OpenTextFile(path,2,True)
   WScript.Sleep(10000)
   file.Close
   Set file=Nothing
End If

Do
	
	Do
	  Set file=fso.OpenTextFile(path,1)
	  If Err.Number=0 Or IsObject(file) Then 
	   Exit Do
	  End If
	  WScript.Sleep(15000)
	  Set file=nothing
	  Err.Clear
    Loop
	strall=file.ReadAll
	'MsgBox strall
	linenum=file.Line
	WScript.Sleep(150)
	file.close
	Set file=Nothing

    	strarry=Split(strall,vbCrLf)
    	'MsgBox oldlinenum & vbCrLf & linenum
 	    For i =(oldlinenum-1) To linenum
	     line=strarry(i)
	     ' this is where to write to SQL
	     'MsgBox line
	      If line<>"" Then 
	        linearry=Split(line," ")
	        TestID=linearry(0)
	        Build=linearry(1)
	        numb=UBound(linearry)
	        WScript.Sleep(5000)  ' for some reason numb is not computed right away!
	        
	        If numb>3 Then  
	          ' Update SQL
	          Result=linearry(2)
	          ResultDetails2=linearry(3)
	          testtable=linearry(4)
	          numb=0
	          'MsgBox ResultDetails2&" "&Result&" "&TestID&" "&Build&" "&linearry(3)
	          If testtable="IR" then
	            WShell.Run "D:\ShareThis\EvernixScripts\AccessDBSQL.exe "&ResultDetails2&" "&Result&" "&TestID&" "&Build&" "&testtable,4,True
	          Else
	            'WShell.Run "D:\ShareThis\EvernixScripts\AccessDBSQL.exe "&ResultDetails2&" "&Result&" "&TestID&" "&Build&" "&linearry(3),4,True
	            WShell.Run "D:\ShareThis\EvernixScripts\AccessDBSQL.exe "&ResultDetails2&" "&Result&" "&TestID&" "&Build,4,True
	          End If  
	          WScript.Sleep(2000)
	        ElseIf numb=2 Then
	          ' Insert into SQL IR
	          WShell.Run "D:\ShareThis\EvernixScripts\InsertIntoAccDB.exe "&TestID&" "&Build&" "&linearry(2),4,True
	          WScript.Sleep(2000)
	        Else
	          ' Insert into SQL non IR
	          WShell.Run "D:\ShareThis\EvernixScripts\InsertIntoAccDB.exe "&TestID&" "&Build,4,True
	          WScript.Sleep(2000)
	        End If
	      
	      End If
	    Next 
	    strarry=""
	    strall=""
	   
	   WScript.Sleep(40000)
	   oldlinenum=linenum

Loop 