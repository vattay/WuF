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

Const WUFC_DROPBOX_PILL_DIR = "pills"
Const WUFC_DROPBOX_ORIGIN_DIR = "origin"

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
Dim gPillLocation
Dim gOriginLocation
Dim gDropResultPostfix
Dim gDateTimeStamp
Dim gStrGroupFileLocation

Dim gBooRestart
Dim gBooAttached 
Dim gBooInteractive
Dim gBooImpersonate

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
	
	gBooImpersonate = TRUE
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
	autoConfig()
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
		
	gPillLocation = gDropBoxLocation & "\" & WUFC_DROPBOX_PILL_DIR
	gOriginLocation = gDropBoxLocation & "\" & WUFC_DROPBOX_ORIGIN_DIR
		
	createFolderChecked( gDropBoxLocation )
	createFolderChecked( gPillLocation ) 
	createFolderChecked( gOriginLocation )
	
	stdOut.writeLine( "Drop box location: " & gDropBoxLocation )
End Function

'*******************************************************************************
Function createFolderChecked( strPath )
	On Error Resume Next
		oFso.CreateFolder strPath
	If (Err.Number <> 0 ) Then
		On Error GoTo 0
		abort( "Unable to write folder at: " & strPath & ", quitting" )
	End If
	On Error GoTo 0
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
				gBooImpersonate = FALSE
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
			ElseIf ( strCompS(arg,"i") Or strCompS( arg, "interactive" ) ) Then
				gBooInteractive = TRUE
			ElseIf ( strCompS( arg, "r" ) Or strCompS( arg, "restart" ) ) Then
				gBooRestart = TRUE
			ElseIf ( strCompS(arg,"a") Or strCompS( arg, "attached" ) ) Then
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
Function autoConfig()

	If ( ( gBooActionDownload) Or (gBooActionInstall) ) Then
		gBooActionScan = TRUE
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
	
	If ( gBooImpersonate ) Then
		If (gPassword <> "" ) Then
			abort( "Can't use a password when impersonating current account" )
		End If
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
	On Error Resume Next
		code = sh.Run("psexec.exe -s echo.", 1, true)
	If ( Err.number <> 0 ) Then abort( "Could not run psexec, errorcode = " & Err.number )
	On Error GoTo 0

End Function

'*******************************************************************************
Function run()

	Dim arrComputers
	
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
		
			Dim booDeployed
		
			stdOut.writeline strComputer & " (" & i & "/" & intServerCount & ")"
			
			If (gBooInteractive) Then getAdminAccess( strComputer )
			
			booDeployed = deployAgent( strComputer, strRemoteAgentName )
			
			call createStub( strComputer, booDeployed)
			call createStat( strComputer, booDeployed, "copy" )
			
			If ( booDeployed ) Then
				Dim exitCode
				exitCode =  executeAgent(strComputer, gBooAttached, gBooRestart, _
					true, strRemoteAgentName)
					
				If ( gBooAttached ) Then
					If ( exitCode <> 0 ) Then call createStat( strComputer, false, "exec.attached.code_" & exitCode )
				Else 
					call createStat( strComputer, true, "exec.detached.pid_" & exitCode )
				End If
				
			Else
				stdErr.writeLine ( "Could not copy agent to remote host: " & strComputer )
			End If
			
			i = i + 1
			
		End If
		
	Next
	
	For Each strComputer In arrComputers
		
		call deleteAgent(strComputer, strRemoteAgentName)
		
		If (gBooInteractive) Then releaseAdminAccess( strComputer )
		
	Next
	
End Function

'*******************************************************************************
Function createStat( strComputer, booSuccess, strMsg )

	Dim exitCode
	Dim strCmd
	Dim strPostFix
	
	If (strMsg <> "") Then strMsg = "." & strMsg
	
	If ( booSuccess ) Then
		strPostFix = ".win" & strMsg
	Else
		strPostFix = ".fail" & strMsg
	End If
	
	strCmd = "cmd /c echo. 2>" & gOriginLocation & _
		"\" & strComputer & strPostFix
		
	exitCode = sh.Run( strCmd, 0, true )
	
End Function

'*******************************************************************************
Function createStub( strComputer , booSuccess)

	Dim exitCode
	Dim strCmd
	Dim strPostFix
	
	If ( booSuccess ) Then
		strPostFix = WUFC_DEFAULT_RESULT_POSTFIX
	Else
		strPostFix = ".fail"
	End If
	
	strCmd = "cmd /c echo. 2>" & gDropBoxLocation & _
		"\" & strComputer & strPostFix
		
	exitCode = sh.Run( strCmd, 0, true )
	
End Function

'*******************************************************************************
Function getAdminAccess( strComputer )

	Dim strAdminShare
	
	strAdminShare = "\\" & strComputer & "\ADMIN$"
	
	If ( gBooImpersonate ) Then
		getAdminAccess = useShareImpersonate( strAdminShare )
	Else
		getAdminAccess = useShare( strAdminShare, gUsername, gPassword )
	End If
	
End Function

'*******************************************************************************
Function releaseAdminAccess( strComputer )
	releaseAdminAccess = deleteShare( "\\" & strComputer & "\ADMIN$" )
End Function

'*******************************************************************************
Function useShare( strShare, strUsername, strPassword )

	Dim exitCode
	Dim strUserArg
	Dim strCmdUse
	
	strUserArg = "/USER:" & strUsername

	strCmdUse = "net use " & strShare & " " & _
		strPassword & _
		" " & strUserArg
	exitCode = sh.Run( strCmdUse, 1, True )

	useShare = exitCode
End Function

'*******************************************************************************
Function useShareImpersonate( strShare )

	Dim strCmdUse
	Dim exitCode

	strCmdUse = "net use " & strShare 
	exitCode = sh.Run( strCmdUse, 1, True )

	useShareImpersonate = exitCode
	
End Function

'*******************************************************************************
Function deleteShare( strShare )
	
	Dim strCmdUse
	Dim exitCode
	
	strCmdUse = "net use " & strShare & _
		" /DELETE"
	exitCode = sh.Run( strCmdUse, 1, True )
	
	stdOut.writeLine "Net Use Delete exit code: " & exitCode
	
	If ( exitCode <> 0 ) Then
		stdErr.writeLine( "NET USE DELETE Failed, residual shares may still be in use." )
	End If
	
	deleteShare = exitCode
	
End Function

'*******************************************************************************
Function deployAgent( strComputer, strRemoteAgentName ) 'return bool
	
	Dim exitCode
	Dim strCmdCopy, strCmdUse
	Dim strUserArg
	
	deployAgent = True
	
	stdOut.writeLine( "Deploying agent to: " & strComputer )
	
	strCmdCopy = "cmd /c copy /V " & WUFC_MASTER_AGENT_LOCATION & _
		" \\" & strComputer & "\ADMIN$\temp\" & strRemoteAgentName
	
	stdOut.writeLine( "Attempting copy." )
	exitCode = sh.Run( strCmdCopy, 1, true )
	
	If ( exitCode = 0 ) Then
		stdOut.writeLine "Copy Successful."
		Exit Function
	Else
		stdErr.writeLine "Copy exit code: " & exitCode
		stdErr.writeLine "Copy Failed."
		deployAgent	= False	
		Exit Function
	End If
	
End Function

'*******************************************************************************
Function executeAgent(strComputer, booAttached, booRestart, booSystem, strRemoteAgentName)

	Dim exitCode
	Dim strCmd
	Dim strArgAttached
	Dim strArgRestart
	Dim strSysArg
	Dim strArgCmdCloseArg
	Dim strRemoteExecutor
	

	strArgAttached = "-d"
	strArgCmdCloseArg = "/c"
	
	If ( booAttached ) Then
		strArgAttached = ""
		strArgCmdCloseArg = "/k"
	End If
	
	If ( booSystem ) Then strSysArg = "-s" Else strSysArg = ""
	
	If ( booRestart )  Then strArgRestart = " /sR " Else strArgRestart = "" 
	
	strRemoteExecutor = " psexec.exe -accepteula " & strArgAttached & " " & strSysArg  & _
		" \\" & strComputer & _
		" -w C:\Windows\Temp "

	stdOut.writeLine( "Remote executing wuf agent on: " & strComputer )
	
	strCmd = "cmd " & strArgCmdCloseArg & _
		strRemoteExecutor & _
		" c:\windows\system32\cscript.exe //I //NoLogo c:\windows\temp\" & _
		strRemoteAgentName & " " & translateAction() & _
		" /oN:" & gDropBoxLocation & "\" & strComputer & WUFC_DEFAULT_RESULT_POSTFIX & _
		" /pS:" & gPillLocation & _
	    strArgRestart

	'WScript.echo strCmd
	
	If (booAttached) Then
		exitCode = sh.Run ( strCmd, 1, true )
		
		If ( exitCode = 0 ) Then
			stdOut.writeLine "Remote executed agent."
		Else
			If ( exitCode = 5 ) Then abort("Insufficient access, try running as admin.")
			stdErr.writeLine("Unable to remote execute agent on:" & strComputer)
		End If
	Else
		exitCode = sh.Run ( strCmd, 1, true )
		stdOut.writeLine "PSexec exit code: " & exitCode
	End If
	
	executeAgent = exitCode
	
	
End Function

'*******************************************************************************
Function deleteAgent(strComputer, strAgentName)

	Dim strCmd
	Dim exitCode
	
	strCmd = "psexec.exe -d -s \\" & strComputer & _
		" cmd /c del c:\windows\temp\" & strAgentName
		
	'WScript.echo strCmd
		
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