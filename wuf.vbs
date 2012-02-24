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

Const WUFC_INPUT_ERROR = 10010

Const WUFC_USAGE_HEAD = "Usage: wuf.vbs [options] [/box:dropbox] /group:<groupfile> {actions}"

Const WUFC_USAGE_ACTIONS = "Actions: ( AUTO | SCAN | DOWNLOAD | INSTALL )"
Const WUFC_USAGE_OPTIONS = "Options: ( /r | /a )"

'GLOBAL ---
Dim stdErr, stdOut, stdIn	'std stream access
Dim gDropBoxRootLocation
Dim gDropResultPostfix
Dim gDateTimeStamp
Dim gStrGroupFileLocation

Dim gBooRestart
Dim gBooAttached 

Dim gBooActionAuto
Dim gBooActionScan
Dim gBooActionDownload
Dim gBooActionInstall

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
	
End Function

'*******************************************************************************
Function init()

	Set stdOut = WScript.StdOut
	Set stdErr = Wscript.StdErr
	Set stdIn = WScript.StdIn

	gDropBoxRootLocation = ""
	gDropResultPostfix = WUFC_DEFAULT_RESULT_POSTFIX
	
	gDateTimeStamp = getDateTimeStamp()
	
	gBooActionAuto = FALSE
	gBooActionScan = FALSE
	gBooActionDownload = FALSE
	gBooActionInstall = FALSE
	
End Function

'*******************************************************************************
Function config()
	parseArgs()
	checkConfig()
	checkSystem()
	run()
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
			If ( strCompS(arg,"group") ) Then
				gStrGroupFileLocation = Wscript.Arguments.Named( arg )
				booGroupGiven = TRUE
			ElseIf ( strCompS(arg,"box") ) Then
				gDropBoxRootLocation = Wscript.Arguments.Named( arg )
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
			ElseIf ( strCompS(arg,"r") ) Then
				gBooRestart = TRUE
			ElseIf ( strCompS(arg,"a") ) Then
				gBooAttached = TRUE
			ElseIf ( strCompS(arg,"h") ) Then
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
	Dim oFso
	Set oFso = CreateObject("Scripting.FileSystemObject")

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

	Dim WshShell, code
	Set WshShell = WScript.CreateObject("WScript.Shell")
	code = WshShell.Run("psexec.exe", true)
	If ( code = 9009 ) Then abort("Psexec not available.")
	
End Function

'*******************************************************************************
Function run()
	'@@TODO Copy, remote exec, and delete agent
End Function

'*******************************************************************************
Function abort(strMsg)

	stdErr.WriteLine("Wuf: " & strMsg)
	cleanup()
	WScript.Quit
	
End Function

'*******************************************************************************
Function cleanup()
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
'******************************************************************************
Function attachedExec(strCommand)
	Set oExec = WshShell.Exec( strCommand )
	Do While ( oExec.Status = 0 )
		WScript.Sleep( 100 )
		consumeIo( oExec )
		WScript.Sleep( 100 )
	Loop
End Function

'******************************************************************************
Function consumeIO(e)
	Do While ( Not e.SrdOut.AtEndOfStream ) OR ( NOT e.StdErr.AtEndOfStream)
		WScript.StdOut.WriteLine(">>" & e.StdOut.ReadLine )
		WScript.StdErr.WriteLine(">>" & e.StdErr.ReadLine )
	Loop
End Function

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