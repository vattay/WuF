Option Explicit
'*******************************************************************************
'Wuf Advanced Controller
'Copyright (C) 2011 Anton Vattay

'This program is free software: you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation, either version 3 of the License, or
'(at your option) any later version.

'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.

'You should have received a copy of the GNU General Public License
'along with this program.  If not, see <http://www.gnu.org/licenses/>.
'*******************************************************************************

'CONST ---
Const WUFC_DEFAULT_RESULT_POSTFIX = ".result.txt"
Const WUFC_MASTER_AGENT_LOCATION = "agent\wuf_agent.vbs"
Const WUFC_DROPBOX_FILE_LOCATION = "dropbox.txt"
Const WUFC_PSEXEC_TEMP_FILE  = "C:\windows\temp\wuf.psexec.out.tmp"

Const WUFC_INPUT_ERROR = 10010

Const WUFC_USAGE_HEAD = "Usage: wuf.vbs [options] [/box:dropbox] /group:<groupfile> {actions}"

Const WUFC_USAGE_ACTIONS = "Actions: ( AUTO | SCAN | DOWNLOAD | INSTALL )"
Const WUFC_USAGE_OPTIONS = "Options: ( /r | /a | /i )"

'GLOBAL ---
Dim stdErr, stdOut, stdIn	'std stream access
Dim sh
Dim oFso

Dim gDropBoxRootLocation
Dim gDropBoxLocation
Dim gDropResultPostfix
Dim gDateTimeStamp
Dim gStrGroupFileLocation

Dim gBooRestart
Dim gBooAttached 
Dim gBooInteractive

Dim gBooActionAuto
Dim gBooActionScan
Dim gBooActionDownload
Dim gBooActionInstall

Dim gPassword
Dim gUsername

' main injection point
main()

'*******************************************************************************
Function main()

	If Not(isCscript()) Then
		WScript.echo  "Unsupported script host, this program must be run with cscript." 
		WScript.quit
	End If
	
	call init()
	call config()
	call run()
	
End Function

'*******************************************************************************
Function init()

	Set stdOut = WScript.StdOut
	Set stdErr = Wscript.StdErr
	Set stdIn = WScript.StdIn

	gDropBoxRootLocation = ""
	gDropBoxLocation = ""
	gDropResultPostfix = WUFC_DEFAULT_RESULT_POSTFIX
	
	gDateTimeStamp = getDateTimeStamp()
	
	gBooRestart = FALSE
	
	gBooInteractive = FALSE
	gBooActionAuto = FALSE
	gBooActionScan = FALSE
	gBooActionDownload = FALSE
	gBooActionInstall = FALSE
	
	Set oFso = CreateObject("Scripting.FileSystemObject")
	Set sh = WScript.CreateObject("WScript.Shell")
	
End Function

'*******************************************************************************
Function config()

	parseArgs()
	checkConfig()
	checkSystem()
	
	'Init dead server list.
	call sh.Run( "cmd /c echo. 2>dead.txt", 0, true )
	
	Dim strR
	If ( gBooRestart ) Then strR = "R" Else strR = ""
	
	gDropBoxLocation = gDropBoxRootLocation & "\" & _
		getDateTimeStamp() & "_" & _
		oFso.getBaseName(gStrGroupFileLocation) & "_" & _
		getActionString() & "_" & _
		strR
		
	oFso.CreateFolder gDropBoxLocation
	
	stdOut.writeLine( "Drop box location: " & gDropBoxLocation )
End Function

'*******************************************************************************
Function getActionString()
	getActionString = ""
	If ( gBooActionAuto ) Then getActionString = getActionString & "A"
	If ( gBooActionScan ) Then getActionString = getActionString & "S"
	If ( gBooActionDownload ) Then getActionString = getActionString & "D"
	If ( gBooActionInstall ) Then getActionString = getActionString & "I"
End Function


'*******************************************************************************
Function parseArgs()

	Dim arg
    Dim objArgs, objNamedArgs, objUnnamedArgs
	Dim booGroupGiven
	booGroupGiven = FALSE
	Dim booActionGiven
	booActionGiven = FALSE
	
	Set objArgs = WScript.Arguments
	
	Set objNamedArgs = WScript.Arguments.named
	Set objUnnamedArgs = WScript.Arguments.unnamed
	
	If (objArgs.Count > 0) Then
		Dim i
		For Each arg in objNamedArgs
			If ( strCompS(arg,"group") Or strCompS(arg,"g") ) Then
				gStrGroupFileLocation = Wscript.Arguments.Named( arg )
				booGroupGiven = TRUE
			ElseIf ( strCompS(arg,"box") Or strCompS(arg,"b")) Then
				gDropBoxRootLocation = Wscript.Arguments.Named( arg )
			ElseIf ( strCompS(arg,"user") ) Then
				gUsername = Wscript.Arguments.Named( arg )
			ElseIf ( strCompS(arg,"password") ) Then
				gPassword = Wscript.Arguments.Named( arg )
			ElseIf ( strCompS(arg,"auto") ) Then
				booActionGiven = TRUE
				gBooActionAuto = TRUE
			ElseIf ( strCompS(arg,"scan") ) Then
				booActionGiven = TRUE
				gBooActionScan = TRUE
			ElseIf ( strCompS(arg,"download") ) Then
				booActionGiven = TRUE
				gBooActionDownload = TRUE
			ElseIf ( strCompS(arg,"install") ) Then
				booActionGiven = TRUE
				gBooActionInstall = TRUE
			ElseIf ( strCompS(arg,"i") ) Then
				gBooInteractive = TRUE
			ElseIf ( strCompS(arg,"r") ) Then
				gBooRestart = TRUE
			ElseIf ( strCompS(arg,"a") ) Then
				gBooAttached = TRUE
			ElseIf ( strCompS(arg,"h") Or strCompS(arg,"?") ) Then
				printHelp()
				abort("Over and Out.")
			Else
				abort( "Unknown named argument: " & arg )
			End If
		Next
	End If
	
	For Each arg in objUnnamedArgs
		abort( "Unknown argument: " & arg )
	Next
	
	' Check Args
	If ( NOT booGroupGiven ) Then
		abort("The group argument is required.")
	End If
	
	If ( NOT booActionGiven ) Then
		abort( "No action argument, nothing to do." )
	End If
End Function

'*******************************************************************************
' Review configuration for logical problems
Function checkConfig()

	If ( gBooActionAuto ) Then
		If ( gBooActionScan OR gBooActionDownload OR gBooActionInstall ) Then
			abort( "Can't perform 'Auto' action with other actions" )
		End If
	End If
	
	If ( gBooRestart ) Then
		Dim choice
		StdOut.Write "Are you sure you want to restart all of them? (y/n): "
		choice = StdIn.ReadLine
		
		If ( NOT strCompI(choice, "y") ) Then abort("Quitting")
	End If
	
	If ( gDropBoxRootLocation = "" ) Then
		stdOut.WriteLine( "Dropbox root argument empty, checking: " &_
			WUFC_DROPBOX_FILE_LOCATION)
		If (NOT oFso.FileExists(WUFC_DROPBOX_FILE_LOCATION)) Then 
			abort("No dropbox location provided.")
		Else
			Dim fDrop
			Set fDrop = oFso.openTextFile(WUFC_DROPBOX_FILE_LOCATION,1 ,false, 0)
			'@@TODO Does this work with network paths?
			gDropBoxRootLocation = fDrop.ReadLine
			fDrop.close
		End If
	End If
	
	If (NOT oFso.FileExists(gStrGroupFileLocation)) Then abort("Group file doesn't exist.")
	
	If (NOT oFso.FileExists(WUFC_MASTER_AGENT_LOCATION)) Then abort("Agent master doesn't exist.")
	
	If (NOT oFso.FolderExists(gDropBoxRootLocation)) Then abort("Dropbox root doesn't exist: " & gDropBoxRootLocation)
	
End Function

'*******************************************************************************
Function checkSystem()

	Dim code
	code = sh.Run("psexec.exe -s echo.", 0, true)
	If ( code = 9009 ) 	Then abort("Psexec not available.")
	'If ( code = 5 )		Then abort("Insufficient permissions, try running as admin.")
End Function

'*******************************************************************************
Function run()

	Dim arrComputers
	'@@TODO Copy, remote exec, and delete agent
	arrComputers = splitTextFile( gStrGroupFileLocation )
	
	mapAgent(arrComputers)
	
End Function

'*******************************************************************************
Function splitTextFile(strFileLocation)

	Dim objGroupFile
	Dim strComputers
	Dim arrComputers
	
	Set objGroupFile = oFso.openTextFile( gStrGroupFileLocation,1,false,0 )
	
	strComputers = objGroupFile.ReadAll
	
	arrComputers = split( strComputers, VbCrLf )
	
	objGroupFile.close
	
	splitTextFile = arrComputers
	
End Function

'*******************************************************************************
Function mapAgent(arrComputers)

	Dim strComputer
	Dim strRemoteAgentName
	Dim i, intServerCount
	strRemoteAgentName = "local_" & oFso.getFileName( WUFC_MASTER_AGENT_LOCATION )
	
	WScript.echo strRemoteAgentName
	
	i = 0
	intServerCount = UBound(arrComputers)
	
	For Each strComputer In arrComputers
		If ( strComputer <> "" ) Then
			stdOut.writeline strComputer & " (" & i & "/" & intServerCount & ")"
			Dim booDeployed
			booDeployed = deployAgent(strComputer, strRemoteAgentName)
			call createStub(strComputer, booDeployed)
			If ( booDeployed ) Then
				call executeAgent(strComputer, gBooAttached, gBooRestart, true, strRemoteAgentName)
				'call deleteAgent(strComputer, strRemoteAgentName)
			Else
				stdErr.writeLine ( "Could not copy agent to remote host: " & strComputer )
			End If
			i = i + 1
		End If
	Next
	
End Function

'*******************************************************************************
Function createStub(strComputer, booSuccess)

	Dim exitCode
	Dim strCmd
	Dim strPostFix
	
	If ( booSuccess ) Then
		strPostFix = WUFC_DEFAULT_RESULT_POSTFIX
	Else
		strPostFix = ".fail"
	End If
	
	strCmd = "cmd /c echo. 2>" & gDropBoxLocation &_
		"\" & strComputer & strPostFix
		
	exitCode = sh.Run( strCmd, 0, true )
	
End Function

'*******************************************************************************
Function deployAgent(strComputer, strRemoteAgentName) 'return bool
	
	Dim exitCode
	Dim strCmdCopy, strCmdUse
	Dim strUserArg
	
	deployAgent = True
	
	stdOut.writeLine( "Deploying agent to: " & strComputer )
	
	strCmdCopy = "cmd /c copy /V " & WUFC_MASTER_AGENT_LOCATION &_
		" \\" & strComputer & "\ADMIN$\temp\" & strRemoteAgentName
	
	stdOut.writeLine( "Attempting copy." )
	exitCode = sh.Run( strCmdCopy, 1, true )
	
	If ( exitCode = 0 ) Then
		stdOut.writeLine "Copy successful"
		Exit Function
	ElseIf ( gBooInteractive ) Then
		stdErr.writeLine "Copy exit code: " & exitCode
		stdErr.writeLine "Initial copy failed, attemping to use network share."
	Else
		stdErr.writeLine "Copy exit code: " & exitCode
		stdErr.writeLine "Copy Failed."	
		deployAgent	= False	
		Exit Function
	End If
	
	If ( gUsername <> "" ) Then
		strUserArg = "/USER:" & gUsername
	End If
	
	'WScript.echo strCmdUse
	
	strCmdUse = "net use \\" & strComputer & "\ADMIN$ " & _
		gPassword & _
		" " & strUserArg
	exitCode = sh.Run( strCmdUse, 1, True )
	
	If ( exitCode <> 0 ) Then
		stdErr.writeLine( "NET USE Failed, trying copy anyway..." )
	End If
	
	stdOut.writeLine "NET USE exit code: " & exitCode
	
	exitCode = sh.Run( strCmdCopy, 1, true )
	
	stdOut.writeLine "Copy exit code: " & exitCode
	
	If ( exitCode <> 0 ) Then 
		stdErr.writeLine( "Copy failed to: " & strComputer )
		deployAgent = False
		call sh.Run( "cmd /c echo " & strComputer & " >> dead.txt", 0, True )
	End If
	
	strCmdUse = "net use \\" & strComputer & "\ADMIN$ " & _
		" /DELETE"
	exitCode = sh.Run( strCmdUse, 1, True )
	
	stdOut.writeLine "Net Use Delete exit code: " & exitCode
	
	If ( exitCode <> 0 ) Then
		stdErr.writeLine( "NET USE DELETE Failed, residual shares may still be in use." )
	End If
	
End Function

'*******************************************************************************
Function executeAgent(strComputer, booAttached, booRestart, booSystem, strRemoteAgentName)

	Dim exitCode
	Dim strCmd
	Dim strArgAttached
	Dim strArgRestart
	Dim strTempFileRedirect
	DIm strSysArg
	Dim strArgCmdCloseArg

	strTempFileRedirect = ""
	strArgAttached = "-d"
	strArgCmdCloseArg = "/k"
	
	If ( booAttached ) Then 
		strArgAttached = ""
		strTempFileRedirect = " 2>&1>" & WUFC_PSEXEC_TEMP_FILE
		strArgCmdCloseArg = "/c"
	End If
	
	If ( booSystem ) Then strSysArg = "-s" Else strSysArg = ""
	
	If ( booRestart )  Then strArgRestart = " /sR " Else strArgRestart = "" 

	stdOut.writeLine( "Remote executing wuf agent on: " & strComputer )
	
	strCmd = "cmd " & strArgCmdCloseArg & " psexec.exe -accepteula " & strArgAttached & " " & strSysArg  & _
		" \\" & strComputer & _
		" -w C:\Windows\Temp c:\windows\system32\cscript.exe //I //NoLogo c:\windows\temp\" & _
		strRemoteAgentName & " " & translateAction() & _
		" /oN:" & gDropBoxLocation & "\" & strComputer & WUFC_DEFAULT_RESULT_POSTFIX & _
		" /pS:" & gDropBoxLocation & _
	    strArgRestart '& strTempFileRedirect
		
	WScript.echo strCmd
	

	If (booAttached) Then
		exitCode = sh.Run ( strCmd, 1, booAttached )
		'Dim tempFile
		'Set tempFile = oFso.OpenTextFile(WUFC_PSEXEC_TEMP_FILE, 1 ,false, 0)
		'stdOut.WriteLine( tempFile.ReadAll )
		'tempFile.close
		
		If ( exitCode = 0 ) Then
			stdOut.writeLine "Remote executed agent."
		Else
			If ( exitCode = 5 ) Then abort("Insufficient access, try running as admin.")
			stdErr.writeLine("Unable to remote execute agent on:" & strComputer)
		End If
	Else
		exitCode = sh.Run ( strCmd, 7, booAttached )
		stdOut.writeLine "Psexec exit code: " & exitCode
	End If
	
	'oFso.DeleteFile( WUFC_PSEXEC_TEMP_FILE )
	
End Function

'*******************************************************************************
Function deleteAgent(strComputer, strAgentName)

	Dim strCmd
	Dim exitCode
	
	strCmd = "psexec.exe -d -s \\" & strComputer & _
		" cmd /c del c:\windows\temp\" & strAgentName
		
	WScript.echo strCmd
		
	exitCode = sh.Run ( strCmd, 0, true )
	stdOut.writeLine "Agent delete exit code: " & exitCode
End Function

'*******************************************************************************
Function translateAction()
	translateAction = ""
	If ( gBooActionAuto ) Then translateAction = translateAction & "/aA "
	If ( gBooActionScan ) Then translateAction = translateAction & "/aS "
	If ( gBooActionDownload ) Then translateAction = translateAction & "/aD "
	If ( gBooActionInstall ) Then translateAction = translateAction & "/aI "
	
End Function

'*******************************************************************************
Function abort(strMsg)

	stdErr.WriteLine("Wuf: " & strMsg)
	cleanup()
	WScript.Quit
	
End Function

'*******************************************************************************
Function cleanup()
	Set oFso = nothing
	Set sh = nothing
	' Cleanup anything that should be destroyed on exit
End Function


'*******************************************************************************
Function getDateTimeStamp()
	getDateTimeStamp =  getDateStamp() & "_" & getTimeStamp() & "_" & genRandId(100)
End Function

'*******************************************************************************
Function printHelp()

	stdOut.writeLine( WUFC_USAGE_HEAD )
	stdOut.writeLine( "" )
	stdOut.writeLine( WUFC_USAGE_ACTIONS )
	stdOut.writeLine( WUFC_USAGE_OPTIONS )
	
End Function


'Util

'*******************************************************************************
Function strCompS(strA, strB) 'returns boolean
	If (strComp(strA, strB, 0) = 0) Then
		strCompS = true
	Else
		strCompS = false
	End If
End Function

'*******************************************************************************
Function strCompI(strA, strB) 'returns boolean
	If (strComp(strA, strB, 1) = 0) Then
		strCompI = true
	Else
		strCompI = false
	End If
End Function

'**************************************************************************************
Function isCscript() 
	If inStr(ucase(WScript.FullName),"CSCRIPT.EXE") Then
		isCscript = TRUE
	Else
		isCScript = FALSE
	End If
End Function

'*******************************************************************************
Function genRandId(intMax)
	Dim fltRandNum, intRunId
	Randomize
	fltRandNum = Rnd * intMax
	intRunId = CInt (fltRandNum)
	genRandId = intRunId
End Function

'*******************************************************************************
Function getDateStamp()
	Dim someDate
	Dim thisMonth, thisDay, thisYear
	
	thisMonth = right("0" & month(Now()),2)
	thisDay = right("0" & day(Now()),2)
	thisYear = right("0" & year(Now()),2)
	
	someDate = thisMonth & thisDay & thisYear
	getDateStamp = someDate
End Function

'*******************************************************************************
Function getTimeStamp()
	Dim someTime
	Dim sec, min, hr
	
	sec = right("0" & second(time),2)
	min = right("0" & minute(time),2)
	hr = right("0" & hour(time),2)
	someTime = hr & min & sec
	getTimeStamp = someTime
End Function