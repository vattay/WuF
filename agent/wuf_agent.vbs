Option Explicit
'*******************************************************************************
'Wuf Agent
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

'@@TODO:  + Fix installed update count being wrong after /aI
'         + Update Impact sort logic
'		  + Fix async status output weirness for install (too many newlines)
'		  + Add error wrap for async install and download, returns raw WU error

'Settings------------------------------
Const LOG_LEVEL = 3
Const WUF_CATCH_ALL_EXCEPTIONS = 0
Const WUF_ASYNC = True
Const WUF_DEFAULT_SHUTDOWN_DELAY = 15
Const ASYNC_REFRESH_RATE = 1000
Const ASYNC_REFRESH_MODERATION = 10
'--------------------------------------
  
Const LOG_LEVEL_DEBUG = 3
Const LOG_LEVEL_INFO = 2
Const LOG_LEVEL_WARN = 1
Const LOG_LEVEL_ERROR = 0

Const APP_NAME = "Wuf Agent"
Const APP_VERSION = "1.8"
Const WU_AGENT_VERSION = "7.0.6000.374"
Const WU_AGENT_LOCALE_DELIM = "."
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const ForAppending = 8
Const ForWriting = 2
Const ForReading = 1
Const REG_OBJECT_LOCAL = "winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv"
Const WSUS_REG_KEY_PATH = "SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
Const WSUS_REG_KEY_WUSERVER = "WUServer"
Const WSUS_REG_KEY_TARGET_GROUP = "TargetGroup"

Const WUF_INPUT_ERROR = 10001
Const WUF_FEEDBACK_ERROR = 10002
Const WUF_NO_UPDATES = 10003
Const WUF_INVALID_CONFIGURATION = 10004
Const WUF_SEARCH_ERROR = 10005
Const WUF_DOWNLOAD_ERROR = 10006
Const WUF_INSTALL_ERROR = 10007
Const WUF_GENERIC_ERROR = 10008
Const WUF_VERIFY_ERROR = 10009
Const WUF_STREAM_ERROR = 10010
Const WUF_COMMAND_ERROR = 10011

Const WUF_ACTION_UNDEFINED = 0
Const WUF_ACTION_AUTO = 	1
Const WUF_ACTION_SEARCH = 	2
Const WUF_ACTION_DOWNLOAD = 4
Const WUF_ACTION_INSTALL = 	8

Const WUF_SHUTDOWN_UNDEFINED = -1
Const WUF_SHUTDOWN_DONT = 	0
Const WUF_SHUTDOWN_RESTART = 1
Const WUF_SHUTDOWN_SHUTDOWN = 2

Const WUF_DEFAULT_SEARCH_CRITERIA = "IsAssigned=1 and IsHidden=0 and IsInstalled=0 and Type='Software'"
Const WUF_DEFAULT_FORCE_SHUTDOWN_ACTION = false
Const WUF_DEFAULT_ACTION = 1 
Const WUF_DEFAULT_SHUTDOWN_OPTION = 0

Const WUF_DEFAULT_LOG_LOCATION = "." '@TODO change this to use . for default and add log location command line arg

Const WUF_USAGE = "wuf_agent.vbs [/aA | /aS | /aD | /aI] [/sN | /sR | /sH] [/c:<criteria> [/pS:<dir>] [/fS] [/oN:<location_name>] [t:<shut_delay>]"
Const WUF_USAGE2 = "/a* - action, /s* - shutdown action, /fS - force shutdown /c - update criteria, /pS - result pill, oN - result location, s - shutdown delay"

'Globals - avoid modification after initialize()
Dim stdErr, stdOut	'std stream access
Dim gWshShell		'Shell access
Dim gWshSysEnv		'Env access
Dim gLogLocation	'Log location
Dim gAction			'This applications action
Dim gShutdownOption	'Restart, shutdown, or do nothing
Dim gForceShutdown	'Do the shutdown option even if not required
Dim gFileLog 		'Wuf Agent Log object
Dim gRunId			'Unique id of run
Dim gObjUpdateSession 'The windows update session used for all wu operations
Dim gObjDummyDict	'Used for async wu operations
Dim gResOut			'Result writer
Dim gBooUsePill		'Whether to use the result pill or not
Dim gPillDir		'Result Pill Directory
Dim gSearchCriteria 'The Windows Update search criteria
Dim gShutDownDelay	'Time to wait before taking aciton
Dim e				'Exception manager

Set e = New ExceptionManager.init()

Class DummyClass 'set up dummy class for async download and installation calls
	Public Default Function DummyFunction()
	
	End Function		
End Class

' main injection point
main()

'*******************************************************************************
Function main()

	If Not(isCscript()) Then
		WScript.echo  "Unsupported script host, this program must be run with cscript." 
		WScript.quit
	End If
	
	initialize()
	
	If (WUF_CATCH_ALL_EXCEPTIONS = 1) Then
		On Error Resume Next
			core()
		e.catch() 'catch
		On Error GoTo 0
		If (e.isException()) Then
			Dim Ex
			Set Ex = e.getException()
			Dim strMsg
			If Ex.number = cLng("&H80240044") Then
				strMsg = "Insufficient access, try running as administrator." 
				gResOut.recordError( strMsg )
			ElseIf (Ex.number = WUF_INPUT_ERROR) Then
				gResOut.recordError( "Improper input, " & Ex.Description )
				gResOut.recordInfo( WUF_USAGE )
				gResOut.recordInfo( WUF_USAGE2 )
			ElseIf  (Ex.number = WUF_INVALID_CONFIGURATION ) Then
				gResOut.recordError( "Invalid Configuration, " & Ex.Description )
				gResOut.recordInfo( WUF_USAGE )
				gResOut.recordInfo( WUF_USAGE2 )
			ElseIf  (Ex.number = WUF_SEARCH_ERROR ) Then
				gResOut.recordError( "Search Problem, " & Ex.Description )
				logError( e.dump(Ex) )
			ElseIf  (Ex.number = WUF_DOWNLOAD_ERROR ) Then
				gResOut.recordError( "Search Problem, " & Ex.Description )
				logError( e.dump(Ex) )
			ElseIf  (Ex.number = WUF_INSTALL_ERROR ) Then
				gResOut.recordError( "Install Problem, " & Ex.Description )
				logError( e.dump(Ex) )
			ElseIf  (Ex.number = WUF_STREAM_ERROR ) Then
				gResOut.recordError( "Stream access problem, " & Ex.Description )
				logError( e.dump(Ex) )
			Else
				gResOut.recordError(Ex.Description)
				logError( e.dump(Ex) )
				'call logErrorEx( "Unhandled exception: ", Ex)
			End If
			cleanup()
		End If
	Else
		core()
	End If
	
	WScript.quit
End Function

'*******************************************************************************
Function core()
	configure()
	If (verify() = true) Then
		preAction()
		doAction(gAction)
		postAction()
	End If
	cleanup()
End Function

'*******************************************************************************
'Must be called first, configures globals needed for everything else.
Function initialize()

	setGlobalRunId()
	
	Set stdOut = WScript.StdOut
	Set stdErr = Wscript.StdErr
	
	Set gWshShell = WScript.CreateObject("WScript.Shell")
	Set gWshSysEnv = gWshShell.Environment("PROCESS")
	
	Set gObjDummyDict = CreateObject("Scripting.Dictionary")
	
	Call gObjDummyDict.Add("DummyFunction", New DummyClass)
	
	gLogLocation = WUF_DEFAULT_LOG_LOCATION & "\wufa_" & gRunId & ".log"
	Set gResOut = New ResultWriter.init()
	gAction = WUF_ACTION_UNDEFINED
	gShutdownOption = WUF_DEFAULT_SHUTDOWN_OPTION
	gForceShutdown = WUF_DEFAULT_FORCE_SHUTDOWN_ACTION
	gSearchCriteria = WUF_DEFAULT_SEARCH_CRITERIA
	gShutDownDelay = WUF_DEFAULT_SHUTDOWN_DELAY

End Function

'*******************************************************************************
Function configure()

	configure = true
	configureLogFile(gLogLocation)
	
	
	logInfo( APP_NAME & " " & APP_VERSION & " " & Now() )
	logInfo( "Log system initialized." )
	
	logInfo( "Run Id: " & gRunId )
	
	logDebug( "Parsing Configuration" )
	
	call gResOut.writeTitle( APP_NAME, APP_VERSION )

	parseArgs()
	
	call gResOut.writeId( gRunId ) 

	logDebug( "Creating Update Session." )
	Set gObjUpdateSession = CreateObject( "Microsoft.Update.Session" )
	
End Function


'*******************************************************************************
Function parseArgs()
	Dim arg
    Dim objArgs, objNamedArgs, objUnnamedArgs
	Dim success
	
	Dim strOutputLocation
	Dim strPillLocation
	Dim booShutdownFlag
	Dim booUseResultFile
	
	strOutputLocation = ""
	booUseResultFile = false
	booShutdownFlag = false
	success = false
	
	Set objArgs = WScript.Arguments
	
	Set objNamedArgs = WScript.Arguments.named
	Set objUnnamedArgs = WScript.Arguments.unnamed
		
	
	If (objArgs.Count > 0) Then
		Dim i
		For Each arg in objNamedArgs
			Dim strArrTemp
			If ( headStrI(arg,"a") ) Then
				If Not ( parseAction(arg) ) Then
					Err.Raise WUF_INPUT_ERROR, "parseArgs()", "Invalid action " & arg
				End If
			ElseIf ( headStrI(arg,"s") ) Then 
				If (booShutdownFlag) Then 
					Err.Raise WUF_INPUT_ERROR, "parseArgs()", "More than one shutdown option."
				End If
				If Not( parseShutdownOption(arg) ) Then
					Err.Raise WUF_INPUT_ERROR, "parseArgs()", "Invalid shutdown option."
				Else
					booShutdownFlag = true
				End If
			ElseIf ( headStrI(arg,"f") ) Then
				If Not( parseForceShutdown(arg) ) Then
					Err.Raise WUF_INPUT_ERROR, "parseArgs()", "Invalid force option."
				End If
			ElseIf ( headStrI(arg,"o") ) Then
				booUseResultFile = true
				strOutputLocation = parseOutputOption(arg)
			ElseIf ( headStrI(arg,"p") ) Then
				gBooUsePill = true
				gPillDir = parsePillOption(arg)
			ElseIf ( headStrI(arg,"c") ) Then
				gSearchCriteria = Wscript.Arguments.Named( arg )
			ElseIf ( strCompI(arg,"t") ) Then
				Dim strDelay
				strDelay = Wscript.Arguments.Named( arg )
				If (isNumeric(strDelay)) Then
					gShutDownDelay = Wscript.Arguments.Named( cInt(strDelay) )
				Else
					Err.Raise WUF_INPUT_ERROR, "parseArgs()", "Shutdown delay must be numeric"
				End If
			Else
				success = false
				Err.Raise WUF_INPUT_ERROR, "parseArgs()", "Unknown named argument: " & arg	
			End If
		Next
		
		For Each arg in objUnnamedArgs
			success = false
			Err.Raise WUF_INPUT_ERROR, "parseArgs()", "Unknown argument: " & arg
		Next
		
		If (booUseResultFile) Then
			If strOutputLocation = "" Then
				call gResOut.addTeedFileStream(gRunId & ".result.txt", generateShadowLocation())
			Else
				call gResOut.addTeedFileStream(strOutputLocation, generateShadowLocation())
			End If
		Else 
			logInfo("Not using a result file.")
		End If
		
	Else
		' No Args
		success = false
		Err.Raise WUF_INPUT_ERROR, "parseArgs()", "No arguments."	
	End If
	
	checkConfig()
End Function

'*******************************************************************************
Function parseAction(strArgVal) 'return boolean
	parseAction = True
	If ( strCompI(strArgVal,"aA") ) Then
		gAction = gAction or WUF_ACTION_AUTO
	ElseIf ( strCompI(strArgVal,"aS") ) Then
		gAction = gAction or WUF_ACTION_SEARCH
	ElseIf ( strCompI(strArgVal,"aD") ) Then
		gAction = gAction or WUF_ACTION_DOWNLOAD
	ElseIf ( strCompI(strArgVal,"aI") ) Then
		gAction = gAction or WUF_ACTION_INSTALL
	Else
		gAction = gAction or WUF_ACTION_UNDEFINED
		parseAction = False
	End If
End Function


'*******************************************************************************
Function parseShutdownOption(strArgVal) 'return boolean
	parseShutdownOption = True
	If (strCompI(strArgVal,"sN")) Then
		gShutdownOption = WUF_SHUTDOWN_DONT
	ElseIf (strCompI(strArgVal, "sR")) Then
		gShutdownOption = WUF_SHUTDOWN_RESTART
	ElseIf (strCompI(strArgVal, "sH")) Then
		gShutdownOption = WUF_SHUTDOWN_SHUTDOWN
	Else
		gShutdownOption = WUF_SHUTDOWN_UNDEFINED
		parseShutdownOption = False
	End If
End Function

'*******************************************************************************
Function parseForceShutdown(strArgVal) 'return boolean
	parseForceShutdown = True
	If (strCompI(strArgVal,"fS")) Then
		gForceShutdown = True
	Else
		parseForceShutdown = False
	End If
End Function

'*******************************************************************************
Function parseOutputOption(strArgVal) 'return boolean
	parseOutputOption = ""
	If (strCompI(strArgVal,"oN")) Then
		parseOutputOption = Wscript.Arguments.Named( strArgVal )
	Else
		Err.Raise WUF_INPUT_ERROR, "parseArgs()", "Invalid output option."
	End If
End Function

'*******************************************************************************
Function parsePillOption(strArgVal) 'return boolean
	parsePillOption = ""
	If (strCompI(strArgVal,"pS")) Then 'Sync Pill
		parsePillOption = Wscript.Arguments.Named( strArgVal )
	Else
		Err.Raise WUF_INPUT_ERROR, "parseArgs()", "Invalid pill option."
	End If
End Function

'*******************************************************************************
Function checkConfig()

	If ( (gAction and WUF_ACTION_AUTO) <> 0 )  Then
		If ( (gAction and  WUF_ACTION_SEARCH) <> 0 ) OR _
			( (gAction and  WUF_ACTION_DOWNLOAD) <> 0 ) OR _
			( (gAction and  WUF_ACTION_INSTALL) <> 0 ) Then
			
			Err.Raise WUF_INVALID_CONFIGURATION, "checkConfig", "aA cannot be used with other actions"

		End If	
	End If
	
	If NOT ( (gAction and  WUF_ACTION_SEARCH) <> 0 ) Then
		If 	( (gAction and  WUF_ACTION_DOWNLOAD) <> 0 ) OR _
			( (gAction and  WUF_ACTION_INSTALL) <> 0 ) Then
			
			Err.Raise WUF_INVALID_CONFIGURATION, "checkConfig", "aS is required for aD or aI"
			
		End If
	End If

End Function

'*******************************************************************************
Function verify() 'return boolean
	Dim verified
	
	logInfo( "---Verifying Configuration...---" )
	verified = True
	
	If (checkUpdateAgent()) Then
		logInfo( "[+] Windows Update Agent is up to date." )
	Else
		logError( "[-] Windows Update Agent is out of date, failed check." )
		verified = False
	End If
	
	If (isShutdownActionPending()) Then
		logInfo("[?] There is a pending restart required.")
		If ((gAction and WUF_ACTION_INSTALL) = 1) Then
			logWarn("[-] Install requested with pending restart.")
			verified = False
		End If
	Else
		logInfo("[+] No pending restarts.")
	End If
	
	If (verified) then
		logInfo("[+] Verification passed.")
	Else
		logInfo("[-] Verification failed.")
	End If
	
	logInfo( "--- Verification Complete ---" )
	verify = verified
End Function 

'*******************************************************************************
Function preAction()
	logInfo("Performing Pre-Action.")
	logEnvironment()
	logLocalWuSettings()
	gResOut.recordPendingShutdown(isShutdownActionPending())
	logInfo("Pre-Action Complete.")
End Function

'*******************************************************************************
Function doAction(intAction)

	Dim searchResults

	logInfo("Performing Action.")
	
	Dim objUpdateResults
	
	If ((intAction and WUF_ACTION_AUTO) <> 0) Then
		logAutoUpdateSettings()
		autoDetect()
	ElseIf ((intAction and WUF_ACTION_SEARCH) <> 0) Then
		Set searchResults = manualAction(intAction)
	End If
	
	logInfo("Action Complete.")
	
End Function

'*******************************************************************************
Function wuDownloadOp(objSearchResult, booAsync)
		If (booAsync) Then
			Set wuDownloadOp = wuDownloadAsync(objSearchResult)
		Else
			Set wuDownloadOp = wuDownload(objSearchResult)
		End If
End Function

'*******************************************************************************
Function wuInstallOp(objSearchResult, booAsync)
		If (booAsync) Then
			Set wuInstallOp = wuInstallAsync(objSearchResult)
		Else
			Set wuInstallOp = wuInstall(objSearchResult)
		End If
End Function

'*******************************************************************************
Function wuDownloadWrapper(objSearchResult)

	Dim downloadResults
		
	If (WUF_CATCH_ALL_EXCEPTIONS = 0) Then
		On Error Resume Next
			Set downloadResults = wuDownloadOp(objSearchResult, WUF_ASYNC)
		e.catch() 'catch
		On Error GoTo 0
		If (e.isException()) Then
			Dim Ex, strMsg
			Set Ex = e.getException()
			
			If (Ex.number = cLng("&H80240024") ) Then
				strMsg = "No updates to download."
			ElseIf (Ex.number = cLng("&H80240044") ) Then
				strMsg = "Insufficient Access, try Run As Admin."
			Else 
				strMsg = "Unrecognized download exception."
			End If
			
			Dim newEx
			Set newEx = e.preRaise( New ErrWrap.initExM( WUF_DOWNLOAD_ERROR, _
				"wuDownloadWrapper()", strMsg , Ex) )
			Err.Raise newEx.number, newEx.Source, newEx.Description
			
		End If
	Else
		Set downloadResults = wuDownloadOp(objSearchResult, WUF_ASYNC)
	End If
		
	call logDownloadResult(objSearchResult.updates, downloadResults)
	call gResOut.recordDownloadResult(objSearchResult.updates, downloadResults)
	
End Function

'*******************************************************************************
Function wuInstallWrapper(objSearchResult)

	Dim installResult
	
	If (WUF_CATCH_ALL_EXCEPTIONS = 0) Then
		On Error Resume Next
			Set installResult = wuInstallOp(objSearchResult, WUF_ASYNC)
		e.catch() 'catch
		On Error GoTo 0
		If (e.isException()) Then
			Dim Ex, strMsg
			Set Ex = e.getException()
			If (Ex.number = cLng("&H80240024") ) Then
				strMsg = "No updates to install"
			ElseIf (Ex.number = cLng("&H80240044") ) Then
				strMsg = "Insufficient Access, try Run As Admin."
			ElseIf (Ex.number = cLng("&H80240016") ) Then
				strMsg = "Install not allowed due to pending restart or other installation."
			Else 
				strMsg = "Unexpected install problem." 
			End If
			Dim newEx
			Set newEx = e.preRaise( New ErrWrap.initExM( WUF_INSTALL_ERROR, _
				"wuDownloadWrapper()", strMsg , Ex) )
			Err.Raise newEx.number, newEx.Source, newEx.Description
		End If		
	Else
		Set installResult = wuInstallOp(objSearchResult, WUF_ASYNC)
	End If

	call logInstallationResult(objSearchResult.updates,installResult)
	call gResOut.recordInstallationResult(objSearchResult.updates, installResult )
	
End Function

'*******************************************************************************
Function manualAction(intAction)
	Dim searchResults
	Dim intUpdateCount
	
	logDebug("Starting Manual Action.")
	
	intUpdateCount = 0
	
	Set searchResults = wuSearch( gSearchCriteria )
	intUpdateCount = searchResults.Updates.Count
	
	logSearchResult( searchResults )
	gResOut.recordSearchResult( searchResults )
		
	Dim rs
	Set rs = new ResultSummary.init(searchResults)
	
	gResOut.recordInfo("Pre-op=" & rs.generateSummary())

	'If (intUpdateCount > 0) Then
		If ( (intAction and WUF_ACTION_DOWNLOAD) <> 0 ) Then
			wuDownloadWrapper(searchResults)
			gResOut.recordInfo("Post-op=" & rs.generateSummary())
		End If
		
		If ( (intAction and  WUF_ACTION_INSTALL) <> 0 ) Then
			acceptEulas(searchResults)
			wuInstallWrapper(searchResults)
			gResOut.recordInfo("Post-op=" & rs.generateSummary())
		End If
	'End If
	
	If ( gBooUsePill ) Then
		Set searchResults = wuSearch( gSearchCriteria )
		Set rs = new ResultSummary.init(searchResults)
		gResOut.recordInfo("post-check=" & rs.generateSummary())
		Dim resultPill
		Set resultPill = New ResultPill.initS(rs,gPillDir)
		resultPill.write( getComputerName() )
	End If
	
	logDebug("Manual Action completed.")
	
	Set manualAction = searchResults
	
End Function

'*******************************************************************************
Function postAction()
	logInfo("Performing post-actions")
	If (rebootPlanned()) Then
		logInfo("System shutdown action will occur.")
		call shutDownActionDelay(gShutdownOption, WUF_DEFAULT_SHUTDOWN_DELAY)
	End If
	gResOut.recordShutdownPlan(rebootPlanned())
	gResOut.recordComplete()
	logInfo("Completed post-actions")
End Function

'*******************************************************************************
Function cleanup()
	logInfo("Cleaning up")

	Set gWshShell = nothing
	Set gWshSysEnv = nothing
	
	Set gObjDummyDict = nothing
	
	Set gResOut = nothing
	
	logInfo("WUF finished.")
	gFileLog.close
	
	Set stdOut = nothing
	Set stdErr = nothing
End Function

'**************************************************************************************
Function wuSearch(strFilter) 'return ISearchResult
	Dim searchResult
	Dim updateSearcher 
	
	logDebug("Creating Update Searcher.")
	Set updateSearcher = gObjUpdateSession.CreateUpdateSearcher()
	
	logDebug("Update Server Selection = " & updateSearcher.serverSelection)
	'logDebug("Update Server Service ID = " & updateSearcher.serviceID)
	
	logInfo("Starting Update Search.")
	
	On Error Resume Next
		Set searchResult = updateSearcher.Search(strFilter)
	e.catch() 'catch
	On Error GoTo 0
	If ( e.isException() ) Then
		Dim Ex
		Set Ex = e.getException()
		Dim strDsc
		Dim strMsg
		If (Ex = cLng("&H80072F78") ) Then
			strDsc = "ERROR_HTTP_INVALID_SERVER_RESPONSE - The server response could not be parsed."
			strMsg = "The server response could not be parsed."
		ElseIf (Ex = cLng("&H8024402C") ) Then
			strDsc = "WU_E_PT_WINHTTP_NAME_NOT_RESOLVED - Winhttp SendRequest/ReceiveResponse failed with 0x2ee7 error. Either the proxy " _
			& "server or target server name can not be resolved. Corresponds to ERROR_WINHTTP_NAME_NOT_RESOLVED. " 
			strMsg = "Update server name could not be resolved."
		ElseIf (Ex = cLng("&H80072EFD") ) Then 
			strDsc = "ERROR_INTERNET_CANNOT_CONNECT - The attempt to connect to the server failed."
			strMsg = "Unable to connect to udpate server"
		ElseIf (Ex = cLng("&H8024401B") ) Then 
			strDsc = "SUS_E_PT_HTTP_STATUS_PROXY_AUTH_REQ - Http status 407 - proxy authentication required" 
			strMsg = "407 Proxy Authentication is required."
		ElseIf (Ex = cLng("&H8024002B") ) Then 
			strDsc = "WU_E_LEGACYSERVER - The Sus server we are talking to is a Legacy Sus Server (Sus Server 1.0)"
			strMsg = "Legacy SUS servers are not supported."
		ElseIf (Ex = cLng("&H80244018") ) Then 
			strDsc = "SUS_E_PT_HTTP_STATUS_FORBIDDEN HttProxy Status 403"
			strMsg = "Server returned 403 Forbidden"
		ElseIf (Ex = cLng("&H80072F8F") ) Then 
			strDsc = "ERROR_INTERNET_SECURE_FAILURE ErrorClockWrong"
			strMsg = "Unable to establish secure connection due to clock sync issue."
		ElseIf (Ex = cLng("&H80240032") ) Then 
			strDsc = "WU_E_INVALID_CRITERIA - Invalid Criteria String"
			strMsg = "Invalid Criteria String"
		ElseIf (Ex = cLng("&H8024001F") ) Then 
			strDsc = "SUS_E_NO_CONNECTION"
			strMsg = "No network connection available."
		ElseIf (Ex = cLng("&H80070002") ) Then 
			strDsc = "	ERROR_FILE_NOT_FOUND"
			strMsg = "Software Distribution folder likley needs to be cleared out."
		ElseIf (Ex = 7) Then 
			strDsc = "Out of memory - In most cases, this error will be resolved by rebooting the client." 
			strMsg = "Out of Memory"
		Else
			Dim strAddr
			strMsg = "Unknown problem searching for updates." 
		End If
		Dim newEx
		Set newEx = e.preRaise( New ErrWrap.initExM(WUF_SEARCH_ERROR, _
			"wuSearch()", strMsg, Ex) )
		Err.Raise newEx.number , newEx.Source, newEx.Description
	End If
	
	If ( isObject( searchResult ) ) Then
		Set wuSearch = searchResult
	Else
		wuSearch = null
	End If
	
End Function

'*******************************************************************************
Function wuDownload(objSearchResult)

	Dim downloader
	Dim objDownloadResult
	
	logDebug("Creating update downloader.")
	
	Set downloader = gObjUpdateSession.CreateUpdateDownloader() 
	downloader.Updates = objSearchResult.Updates
	
	logInfo("Downloading updates")
	
	Set objDownloadResult = downloader.Download()

	If Not( isObject(objDownloadResult) ) Then
		wuDownload = null
	Else
		Set wuDownload = objDownloadResult
	End If
	
End Function



'*******************************************************************************
Function wuDownloadAsync(objSearchResult)

	Dim downloader, dlJob, dlProgress
	Dim objDownloadResult
	Dim count
	Dim updates
	
	logDebug("Creating update downloader.")
	
	Set downloader = gObjUpdateSession.CreateUpdateDownloader() 
	Set updates = objSearchResult.Updates
	downloader.Updates = updates
	'@@TODO Try to avoid downloading updates already available. Problem is the IUpdate.isDownloaded method is
	'not reliable. No solution yet, it attempts to download everything searched.
	
	logInfo("Downloading Updates Asynchronously")

	Set dlJob = downloader.beginDownload(gObjDummyDict.Item("DummyFunction"),gObjDummyDict.Item("DummyFunction"),vbNull)

	Set dlProgress = dlJob.getProgress()
	
	Dim outputModerator
	outputModerator = ASYNC_REFRESH_MODERATION ' slow file output by factor of...
	Dim i
	i = 0
	While Not getAsyncWuOpJoinable(updates, dlJob)  
		Set dlProgress = dlJob.getProgress()
		Call gResOut.refreshDownloadStatus(dlProgress, updates)
		If (i = 0) Then
			Call gResOut.recordDownloadStatus(dlProgress, updates)
			logInfo( "Download Progress: " & dlProgress.percentcomplete & "%" )
		End IF
		i = (i+1) Mod outputModerator
		WScript.Sleep(ASYNC_REFRESH_RATE)
	Wend
	
	If (dlJob.isCompleted = TRUE) Then 
		logInfo("Asynchronous download completed." )
	Else
		logWarn( "Could not complete asynchronous download, forcing synchronous termination." )
	End If
	
	Set objDownloadResult = downloader.endDownload(dlJob)
	
	dlJob.CleanUp()
	Set dlJob = Nothing

	If Not( isObject(objDownloadResult) ) Then
		wuDownloadAsync = null
	Else
		Set wuDownloadAsync = objDownloadResult
	End If
	
End Function

'*******************************************************************************
Function countMissingUpdates(objSearchResult)
	
	countMissingUpdates = 0

	Dim i
	For i = 0 To objSearchResult.Updates.Count-1
		Dim update
	    Set update = objSearchResult.Updates.Item(i)
	    If (update.IsDownloaded) Then
			logInfo("Update has been downloaded: " & update.Title )
		Else
			logWarn("Update is not downloaded: " & update.Title )
			countMissingUpdates = countMissingUpdates + 1
	    End If
	Next
	
End Function

'*******************************************************************************
Function forceInstallerQuiet(objInstaller)

	On Error Resume Next
		objInstaller.ForceQuiet = True 
	e.catch() 'catch
	On Error GoTo 0
	If (e.isException()) Then
		Dim Ex
		Set Ex = e.getException()
		call logErrorEx("Could not force installer to be quiet.", Ex)
	End If
	
End Function

'*******************************************************************************
Function wuInstall(objSearchResult)

	Dim updatesToInstall
	Dim installationResult
	
	Set updatesToInstall = objSearchResult.Updates
	
	gResOut.recordMissingDownloads( countMissingUpdates(objSearchResult) )
	
	logDebug("Creating Update Installer.")
	
	Dim installer
	Set installer = gObjUpdateSession.CreateUpdateInstaller()
	installer.AllowSourcePrompts = False

	forceInstallerQuiet(installer)
	
	installer.Updates = updatesToInstall
	
	logInfo("Installing updates.")
	
	Set installationResult = installer.Install()

	If Not( isObject(installationResult) ) Then
		wuInstall = null
	Else
		Set wuInstall = installationResult
	End If
	
End Function

'*******************************************************************************
Function wuInstallAsync(objSearchResult)
	Dim installJob, installProgress
	Dim objInstallResult
	
	Dim updatesToInstall
	Set updatesToInstall = objSearchResult.Updates
	
	gResOut.recordMissingDownloads(countMissingUpdates(objSearchResult))
		
	logInfo ( "Number of updates to be installed that are downloaded: " & updatesToInstall.count )

	logDebug( "Creating Update Installer." )
	
	Dim installer
	Set installer = gObjUpdateSession.CreateUpdateInstaller()
	installer.AllowSourcePrompts = False

	forceInstallerQuiet(installer)
	
	installer.Updates = updatesToInstall
	
	logInfo("Installing updates.")
	
	Set installJob = installer.beginInstall(gObjDummyDict.Item("DummyFunction"),gObjDummyDict.Item("DummyFunction"),vbNull)

	set installProgress = installJob.getProgress()
	
	Dim outputModerator
	outputModerator = ASYNC_REFRESH_MODERATION ' slow file output by factor of...
	Dim i
	i = 0
	While Not getAsyncWuOpJoinable(installer.Updates, installJob) 
		set installProgress = installJob.getProgress()
		call gResOut.refreshInstallStatus(installProgress,updatesToInstall)
		If (i = 0) Then
			Call gResOut.recordInstallStatus(installProgress, updatesToInstall)
			logInfo( "Install Progress: " & installProgress.percentcomplete & "%" )
		End IF
		WScript.Sleep(ASYNC_REFRESH_RATE)
		i = (i+1) Mod outputModerator
	Wend
	
	If (installJob.isCompleted = TRUE) Then 
		logInfo("Asynchronous installation completed." )
	Else
		logWarn( "Could not complete asynchronous install, forcing synchronous completion." )
	End If
	
	Set objInstallResult = installer.endInstall(installJob)
	
	installJob.CleanUp()
	Set installJob = Nothing
	
	If Not( isObject(objInstallResult) ) Then
		wuInstallAsync = null
	Else
		Set wuInstallAsync = objInstallResult
	End If
	
End Function

'*******************************************************************************
' The only reason this is used is because the IDownloadJob::isCompleted
' and IInstallJob::isCompleted never returns true in rare situations,
' so they cannot be relied on for action completion.
Function getAsyncWuOpJoinable(objUpdates, objOperationJob)

	Dim i, intTotalResultCode
	intTotalResultCode = 15
	
	Dim objOperationProgress
	Set objOperationProgress = objOperationJob.getProgress()
	
	For i = 0 To objUpdates.count - 1
		intTotalResultCode = intTotalResultCode AND objOperationProgress.getUpdateResult(i).resultCode
	Next
	
	If (intTotalResultCode = 0) AND (objOperationJob.isCompleted = false) Then
		getAsyncWuOpJoinable = false
	Else 
		getAsyncWuOpJoinable = true
	End If
	
End Function

'**************************************************************************************
Function logSearchResult(objSearchResult)

	logInfo("Number of missing updates: " & objSearchResult.Updates.Count)
	
	Dim i
	For i = 0 To (objSearchResult.Updates.Count-1)
		Dim update, objCategories
		Set update = objSearchResult.Updates.Item(i)
		Set objCategories = objSearchResult.Updates.Item(i).Categories
		logInfo("Missing: " & objSearchResult.Updates.Item(i) )
	Next
	
End Function

'**************************************************************************************
Function logDownloadResult(objUpdates, objDownloadResult)

	If NOT (isObject(objDownloadResult) ) Then
		logInfo( "Download result not an object." )
		Exit Function
	End If

	'Output results of install
	logInfo( "Download Result Code: " & objDownloadResult.ResultCode )
	
	logInfo( "Indvidual Update Download Results..." )
	Dim i
	For i = 0 to objUpdates.Count - 1
		Dim strResult
		strResult = objDownloadResult.GetUpdateResult(i).ResultCode
		logInfo(objUpdates.Item(i).Title & ", " & objUpdates.Item(i).identity.updateId & ": " & strResult)
	Next
	
End Function

'**************************************************************************************
Function logInstallationResult(objUpdates, objInstallationResult)

	If NOT (isObject(objInstallationResult) ) Then
		logInfo("Installation result not and object.")
		Exit Function
	End If

	'Output results of install
	logInfo( "Installation Result Code: " & objInstallationResult.ResultCode )
	logInfo( "Reboot Required?: " & objInstallationResult.RebootRequired )
	
	logInfo( "Indvidual Update Installation Results..." )
	Dim i
	For i = 0 to objUpdates.Count - 1
		Dim strResult
		strResult = objInstallationResult.GetUpdateResult(i).ResultCode
		logInfo(objUpdates.Item(i).Title & ": " & strResult)
	Next
End Function

'*******************************************************************************
Function getAuScheduleText()
	Dim strDay
	Dim strTime
	Dim intTime
	Dim objAutoUpdate
	Dim objSettings

	Set objAutoUpdate = CreateObject("Microsoft.Update.AutoUpdate")
	Set objSettings = objAutoUpdate.Settings

	Select Case objSettings.ScheduledInstallationDay
	    Case 0
			strDay = "Every Day"
	    Case 1
			strDay = "Sunday"
	    Case 2
			strDay = "Monday"
	    Case 3
			strDay = "Tuesday"
	    Case 4
			strDay = "Wednesday"
	    Case 5
			strDay = "Thursday"
	    Case 6
			strDay = "Friday"
	    Case 7
			strDay = "Saturday"
	    Case Else
			strDay = "?"
	End Select

	intTime = objSettings.ScheduledInstallationTime

	If (intTime > 12) Then
		intTime = intTime - 12
		strTime = intTime & ":00 PM"
	Else
		If intTime = 0 Then intTime = 12
		strTime = intTime & ":00 AM"
	End If
	
	 getAuScheduleText = strDay & " at " & strTime
End Function

'*******************************************************************************
Function checkUpdateAgent() 'returns boolean (true if version is ok)

	Dim bUpdateNeeded
	Dim autoUpdateSettings
	Dim objAgentInfo
	
	logDebug("Checking version of Windows Update agent against version " _
	& WU_AGENT_VERSION & "...")
	
	On Error Resume Next
		Set objAgentInfo = CreateObject("Microsoft.Update.AgentInfo") 
	e.catch() 'catch
	On Error GoTo 0
	If (e.isException()) Then
		Dim Ex
		Set Ex = e.getException()
		Dim strMsg
		strMsg = "Unable to get Agent Info object, perhaps windows updates haven't been configured?"
		Dim newEx
		Set newEx = e.preRaise( New ErrWrap.initExM( WUF_VERIFY_ERROR, _
			"checkUpdateAgent()", strMsg , Ex) )
		Err.Raise newEx.number, newEx.Source, newEx.Description
	End If

	autoUpdateSettings = objAgentInfo.GetInfo("ProductVersionString")
	If replace(autoUpdateSettings,WU_AGENT_LOCALE_DELIM,"") = replace(WU_AGENT_VERSION,WU_AGENT_LOCALE_DELIM,"") then
		logInfo("File versions match (" & autoUpdateSettings & "). Windows Update Agent is up to date.")
		checkUpdateAgent = True 
	ElseIf (autoUpdateSettings > WU_AGENT_VERSION) Then 
		logDebug("Your installed version of the Windows Update Agent (" & autoUpdateSettings & _
			") is newer than the referenced version (" & WU_AGENT_VERSION & ").")
		checkUpdateAgent = True
	Else
		logError("File version (" & autoUpdateSettings & ") does not match. Windows Update Agent 2.0 required.") 
		checkUpdateAgent = False
	End If 

End Function

'*******************************************************************************
Function getDateStamp()
	Dim someDate
	someDate = Date()
	getDateStamp = Replace(someDate,"/","")
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

'*******************************************************************************
' A non zero delay is recommended so that this script can finish normally
Function shutDownActionDelay(intAction, intDelay)
	Dim strShutDown
	Dim strSysRt
	Dim objShell
	
	Set objShell = CreateObject("WScript.Shell")
	
	strSysRt = WScript.CreateObject("WScript.Shell").ExpandEnvironmentStrings("%SystemRoot%")
	
	If (intAction = WUF_SHUTDOWN_RESTART) Then
		strShutdown = strSysRt & "\system32\cmd.exe /c " & strSysRt & _
			"system32\shutdown.exe /r /t " & intDelay & " /f"
	ElseIf	(intAction = WUF_SHUTDOWN_SHUTDOWN) Then
		strShutdown = strSysRt & "\system32\cmd.exe /c " & strSysRt & _
			"system32\shutdown.exe /s /t " & intDelay & " /f"
	Else 
		Exit Function
	End If
	
	On Error Resume Next
	objShell.Run strShutdown, 0, FALSE
		e.catch() 'catch
	On Error GoTo 0
	If (e.isException()) Then
		Dim Ex, strMsg
		Set Ex = e.getException()
			strMsg = "Shutdown command could not complete."
		Dim newEx
		Set newEx = New ErrWrap.initExM(WUF_COMMAND_ERROR, _
			"shutDownActionDelay()", strMsg, Ex) 
		call logErrorEx("Shutdown action problem", newEx)
		gResOut.recordError( strMsg )
	End If
End Function

'*******************************************************************************
Function isShutdownActionPending() 'return boolean
	Dim computerStatus
	Set computerStatus = CreateObject("Microsoft.Update.SystemInfo") 
	isShutdownActionPending = computerStatus.RebootRequired
End Function

'*******************************************************************************
Function getComputerName() 'return computername
	Dim strLocalComputerName
	strLocalComputerName = WScript.CreateObject("WScript.Shell").ExpandEnvironmentStrings("%Computername%")
	getComputerName = strLocalComputerName
End Function

'**************************************************************************************
Function getAdComputerName() 'returns string (empty if name is not available) ' throws AD exception

	Dim strAdComputerName
	Dim objADInfo
	
	Set objADInfo = CreateObject("ADSystemInfo")
	
	strAdComputerName = objADInfo.ComputerName
	
	getAdComputerName = strAdComputerName
	
End Function

'**************************************************************************************
Function getComputerOU() 'returns string (empty if name is not available) ' throws AD exception

	Dim objComputer
	Dim strOu
	
	On Error Resume Next
		Set objComputer = GetObject("LDAP://" & getAdComputerName())
		If (Err <> 0) Then
			call logInfoEx("Could not get AD Computer name.", Err)
		End If
	On Error GoTo 0
	
	If isObject(objComputer) Then  
		strOU = replace(objComputer.Parent,"LDAP://","")
	Else
		strOU = ""
	End If
	
	getComputerOU = strOu
	
End Function

'**************************************************************************************
Function getUserName() 'return string
	getUsername = gWshSysEnv("username")
End Function

'**************************************************************************************
Function getDomain() 'return string
	getDomain = gWshSysEnv("userdomain")
End Function

'**************************************************************************************
Function isCscript() 
	If inStr(ucase(WScript.FullName),"CSCRIPT.EXE") Then
		isCscript = TRUE
	Else
		isCScript = FALSE
	End If
End Function

'**************************************************************************************
Function configureLogFile(strLogLocation) 'returns file stream
	' Side effect: sets global gFileLog
	Dim fso
	Set fso = WScript.CreateObject("Scripting.FileSystemObject")
	Set gFileLog = fso.OpenTextFile (strLogLocation, ForWriting, True)
	Set configureLogFile = gFileLog
End Function


'**************************************************************************************
Function forceShutdownMessage() 'return string
	If (gForceShutdown = true) Then 
		forceShutdownMessage = "Force"
	Else
		forceShutdownMessage = "Only if pending."
	End If
End Function

'**************************************************************************************
Function shutdownOptionMessage() 'return string
	Select Case gShutdownOption
	  Case WUF_SHUTDOWN_DONT
		shutdownOptionMessage = "Do nothing"
	  Case WUF_SHUTDOWN_RESTART 
		shutdownOptionMessage = "Restart"
	  Case WUF_SHUTDOWN_SHUTDOWN
		shutdownOptionMessage = "Shut down"
	  Case Else
		shutdownOptionMessage = "?"
	End Select
End Function

'**************************************************************************************
Function getWsusServer() ' returns String
	Dim regWSUSServer
	Dim oReg
	
	Set oReg=GetObject(REG_OBJECT_LOCAL)
	
	oReg.GetStringValue HKEY_LOCAL_MACHINE,WSUS_REG_KEY_PATH,WSUS_REG_KEY_WUSERVER,regWSUSServer
	
	If (regWSUSServer = "") Then 
		regWSUSServer = "Microsoft Windows Update Online"
	End If
	
	getWsusServer = regWSUSServer
End Function

'**************************************************************************************
Function getTargetGroup() 'returns String
	Dim regTargetGroup
	Dim oReg
	
	Set oReg=GetObject(REG_OBJECT_LOCAL)
	
	oReg.GetStringValue HKEY_LOCAL_MACHINE, WSUS_REG_KEY_PATH,WSUS_REG_KEY_TARGET_GROUP,regTargetGroup
	
	If (regTargetGroup = "") Then 
	  getTargetGroup = "Not specified"
	End If
	
	getTargetGroup = getTargetGroup
End Function

'**************************************************************************************
Function getAutoUpdateNotificationLevelText(intLevel) 'returns String
	Select Case intLevel
		Case 0 
		  getAutoUpdateNotificationLevelText = "WU agent is not configured."
		Case 1 
		  getAutoUpdateNotificationLevelText = "WU agent is disabled."
		Case 2
		  getAutoUpdateNotificationLevelText = "Users are prompted to approve updates prior to installing"
		Case 3 
		  getAutoUpdateNotificationLevelText = "Updates are downloaded automatically, and users are prompted to install."
		Case 4 
		  getAutoUpdateNotificationLevelText = "Updates are downloaded and installed automatically at a pre-determined time."
		Case Else
	End Select
End Function

'**************************************************************************************
Function logLocalWuSettings()
	
	logInfo("Update Server: " & getWsusServer() )
	
	logDebug("Target Group: " & getTargetGroup() )

End Function

'**************************************************************************************
Function logAutoUpdateSettings()

	Dim autoUpdateClient
	Dim autoUpdateSettings
	
	Set autoUpdateClient = CreateObject("Microsoft.Update.AutoUpdate")
	
	Set autoUpdateSettings = autoUpdateClient.Settings
	
	logInfo("WUA Mode: " & _
		getAutoUpdateNotificationLevelText(autoUpdateSettings.notificationlevel) )
	
	logInfo("WUA Schedule: " & getAuScheduleText() )
End Function

'**************************************************************************************
Function logEnvironment()
	Dim objArgs, strArguments
	Set objArgs = WScript.Arguments
	logInfo( "Computer Name: " & getComputerName() )
	logDebug( "OU: " & getComputerOU() )
	logInfo( "Executed by: " & getDomain() & "\" & getUserName() )
	Dim i
	If (LOG_LEVEL >= LOG_LEVEL_DEBUG) Then
		For i = 0 to (objArgs.Count - 1)
			strArguments = strArguments & " " & objArgs(i)
		Next
		logDebug( "Command arguments: " & strArguments )
	End If
	logInfo("Action: (" & gAction & ") " & getActionMessage(gAction))
	logInfo("Shutdown Option: " & shutdownOptionMessage() )
	logInfo("Force Shutdown Option: " & forceShutdownMessage() )
End Function

'**************************************************************************************
Function getActionMessage(intAction)
	getActionMessage = ""
	If ( (intAction and WUF_ACTION_AUTO) <> 0 ) Then
		getActionMessage = getActionMessage & "Auto "
	End If
	If ( (intAction and  WUF_ACTION_SEARCH) <> 0 ) Then
		getActionMessage = getActionMessage & "Scan "
	End If
	If ( (intAction and  WUF_ACTION_DOWNLOAD) <> 0 ) Then
		getActionMessage = getActionMessage & "Download "
	End If
	If ( (intAction and  WUF_ACTION_INSTALL) <> 0 ) Then
		getActionMessage = getActionMessage & "Install "
	End If
End Function

'**************************************************************************************
Function autoDetect()
	'Force WU Agent to detect
	Dim autoUpdateClient
	Set autoUpdateClient = CreateObject("Microsoft.Update.AutoUpdate")
	logDebug("Attempting to call Windows Auto Update DetectNow method.")

	On Error Resume Next' try
		autoUpdateClient.detectnow()
	e.catch() 'catch
	On Error GoTo 0
	If (e.isException()) Then
		Dim Ex, strMsg
		Set Ex = e.getException()
		If ( Ex.number = cLng("&H8024A000") ) Then
			strMsg = "WU Service Not Running"
		Else
			strMsg = "Unhandled WU Service Exception"
		End If
		Dim newEx
		Set newEx = e.preRaise( New ErrWrap.initExM(WUF_GENERIC_ERROR, _
		"autoDetect()", strMsg, Ex) )
		Err.Raise newEx.number, newEx.source, newEx.description
	End If
End Function

'**************************************************************************************
Function acceptEulas(objSearchResult) 'return ISearchResult
	logDebug("Accepting EULAS on each update...")
	Dim i
	For i = 0 to objSearchResult.Updates.Count-1
		Dim update
		Set update = objSearchResult.Updates.Item(i) 
		If Not update.EulaAccepted Then update.AcceptEula 
	Next 
	Set acceptEulas = objSearchResult
End Function

'*************************************************************************************************************
Function logInfo(strMsg) 
	If (LOG_LEVEL >= LOG_LEVEL_INFO) Then logEntry LOG_LEVEL_INFO, strMsg
End Function 

'*************************************************************************************************************
Function logError(strMsg) 
	If (LOG_LEVEL >= LOG_LEVEL_ERROR) Then LogEntry LOG_LEVEL_ERROR, strMsg
End Function 

'*************************************************************************************************************
Function logWarn(strMsg) 
	If (LOG_LEVEL >= LOG_LEVEL_WARN) Then LogEntry LOG_LEVEL_WARN, strMsg
End Function 

'*************************************************************************************************************
Function logDebug(strMsg) 
	If (LOG_LEVEL >= LOG_LEVEL_DEBUG) Then LogEntry LOG_LEVEL_DEBUG, strMsg
End Function 

'******************************************************************************
Function logEntry(intType, strMsg)
	Dim strLine
	
	strLine = "[" & time & "] - " & getLogTypeLabel(intType) & " - " & strMsg
	
	On Error Resume Next
		gFileLog.writeline strLine
	e.catch() 'catch
	On Error GoTo 0
	If (e.isException()) Then
		Dim Ex
		Set Ex = e.getException()
		stdErr.writeLine "{LOG ERROR} Unable to write to log file. " & getFormattedErrorMsg( Ex )
	End If
	
	If (isCScript()) Then 
		If (intType <= LOG_LEVEL_WARN ) Then
			stdErr.writeLine strLine
		End If
	End If
	
End Function

'******************************************************************************
Function getLogTypeLabel(intType) 'returns String
	Select Case intType
		Case LOG_LEVEL_DEBUG
			getLogTypeLabel = "DEBUG"
		Case LOG_LEVEL_INFO
			getLogTypeLabel = "INFO"
		Case LOG_LEVEL_WARN
			getLogTypeLabel = "WARN"
		Case LOG_LEVEL_ERROR
			getLogTypeLabel = "ERROR"
	End Select
End Function

'******************************************************************************
Function logDebugEx(strMsg, objErr)
	logDebug( strMsg & VbCrLf & _
		getFormattedErrorMsg(objErr) )
End Function

'******************************************************************************
Function logInfoEx(strMsg, objErr)
	logInfo( strMsg & VbCrLf & _
		getFormattedErrorMsg(objErr) )
End Function

'******************************************************************************
Function logWarnEx(strMsg, objErr)
	logWarn( strMsg & VbCrLf & _
		getFormattedErrorMsg(objErr) )
End Function

'******************************************************************************
Function logErrorEx(strMsg, objErr)
	logError( strMsg & VbCrLf & _
		getFormattedErrorMsg(objErr) )
End Function

'******************************************************************************
Function getFormattedErrorMsg( objErr ) 'returns string
	Dim strMessage
	strMessage = ""
	if ( Not(isNull(objErr) ) and ( Not(isEmpty(objErr)) ) ) then
		strMessage = "ExNum:[" & objErr.Number & " : 0x" & hex(objErr.Number) & "] " & VbCrLf & _
			"ExSource: " & objErr.Source & VbCrLf & _
			"ExDescription: " & objErr.Description & VbCrLf
	else
		strMessage = "Error object was empty or null."
	end if
	getFormattedErrorMsg = strMessage
End Function

'*******************************************************************************
Function strCompI(strA, strB) 'returns boolean
	If (strComp(strA, strB, 1) = 0) Then
		strCompI = true
	Else
		strCompI = false
	End If
End Function

'*******************************************************************************
Function headStrI(strHay, strNeedle) ' returns boolean
	If (inStr(1,strHay,strNeedle,VBTEXTCOMPARE) = 1) Then
		headStrI = true
	Else
		headStrI = false
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
Function genRunId()
	genRunId =  getComputerName() & "_" & getDateStamp() & "_" & getTimeStamp() & "_" & genRandId(100)
End Function

'*******************************************************************************
Function setGlobalRunId()
	gRunId = genRunId()
End Function

'*******************************************************************************
Function rebootPlanned()
	rebootPlanned = ( (isShutdownActionPending() _
		AND ( gShutdownOption >= WUF_SHUTDOWN_RESTART )) _
		OR ( gForceShutdown ) )
End Function

'*******************************************************************************
Function generateShadowLocation()
	
	Dim wshShell
	
	Set wshShell = CreateObject("WScript.shell")
	
	Dim tempDir 
	
	tempDir = wshShell.ExpandEnvironmentStrings("%temp%")
	
	generateShadowLocation = tempDir & "\" & gRunId & ".tmp"

End Function

'===============================================================================
'===============================================================================
Class psuedoTee
	Dim fStream
	Dim stdOut
	
	Function init(fStream)
		Set stdOut = WScript.StdOut
		Set me.fStream = fStream
		Set init = me
	End Function
	
	Function writeLine(strMessage) 'write Line
		stdOut.writeLine strMessage
		fStream.writeLine strMessage
	End Function
	
	Function write(strMessage) 'write
		stdOut.writeLine strMessage
		fStream.writeLine strMessage
	End Function
	
	Function close()
		If (isObject(fStream)) Then
			fStream.close()
		End If
	End Function
	
	Sub Class_Terminate
		close()
	End Sub
End Class

'===============================================================================
'===============================================================================
Class ResultPill

	Dim oFso 
	Dim objResultSummary
	Dim strDirectory
	Dim strLastPillName
	
	Function init(objSearchResult, strDirectory)
	
		Set oFso = CreateObject("Scripting.FileSystemObject")
		Set objResultSummary = New ResultSummary.init(objSearchResult)
		
		Me.strDirectory = strDirectory
		Me.strLastPillName = ""
		
		Set init = Me
		
	End Function
	
	'---------------------------------------------------------------------------------
	Function initS(objResultSummary, strDirectory)
	
		Set oFso = CreateObject("Scripting.FileSystemObject")
		Set Me.objResultSummary = objResultSummary
		
		Me.strDirectory = strDirectory
		Me.strLastPillName = ""
		
		Set initS = Me
		
	End Function
	
	'---------------------------------------------------------------------------------
	Function write( strPrefix )
	
		Dim strPillName
		
		If (strDirectory = "") Then
			strPillName = objResultSummary.generatePillName(strPrefix)
		Else
			strPillName = strDirectory & "\" & objResultSummary.generatePillName(strPrefix)
		End If
		
		If NOT ( strLastPillName = "" ) Then
			If oFso.fileExists( strLastPillName ) Then
				oFso.deleteFile( strLastPillName ) 
			End If
		End If
		
		call oFso.createTextFile( strPillName , True )
		
		strLastPillName = strPillName
		
	End Function
	
	'---------------------------------------------------------------------------------
	Function getComputerName() 'returns string
		getComputerName = WScript.CreateObject("WScript.Shell").ExpandEnvironmentStrings("%Computername%")
	End Function
	
End Class

'===============================================================================
'===============================================================================
Class ResultSummary
	Dim objSearchResult
	
	'---------------------------------------------------------------------------------
	Function init(objSearchResult)
	
		Set Me.objSearchResult = objSearchResult
		
		Set init = Me
		
	End Function

	'---------------------------------------------------------------------------------
	Function getSearchResult() 'returns ISearchResult
		getSearchResult = objSearchResult
	End Function
	
	'---------------------------------------------------------------------------------
	Function generateSummary() 'returns String
		
		generateSummary = "Searched=" & getUpdatesSearched() & _
			", Downloaded=" & getDownloadedCount() & _
			", Installed=" & getInstalledCount()
	End Function
	
	'---------------------------------------------------------------------------------
	Function generatePillName( strPrefix ) 'IUpdateSearchResult -> String
		
		generatePillName = strPrefix & _
			"_s" & getUpdatesSearched() & _
			"_d" & getDownloadedCount() & _
			"_i" & getInstalledCount() & ".pil"
	End Function
	
	'---------------------------------------------------------------------------------
	Function getUpdatesSearched()
		getUpdatesSearched = objSearchResult.Updates.Count
	End Function
	
	'---------------------------------------------------------------------------------
	Function getInstalledCount()
	
		Dim intInstalled
		
		intInstalled = 0
	
		Dim i
		For i = 0 To ( objSearchResult.Updates.Count-1 )
			Dim update
			Set update = objSearchResult.Updates.Item(i)
			
			If (update.isInstalled = True) Then
				intInstalled = intInstalled + 1
			End If
		Next
		
		getInstalledCount = intInstalled
	End Function
	
	'---------------------------------------------------------------------------------
	Function getDownloadedCount()
	
		Dim intDownloaded
		
		intDownloaded = 0
	
		Dim i
		For i = 0 To ( objSearchResult.Updates.Count-1 )
			Dim update
			Set update = objSearchResult.Updates.Item(i)
			
			If (update.isDownloaded = True) Then
				intDownloaded = intDownloaded + 1
			End If
		Next
		
		getDownloadedCount = intDownloaded
	End Function
	
End Class


'====================================================================================
'====================================================================================
Class ResultWriter
	Dim stdOut, stdErr
	Dim strResultLocation
	Dim stream
	Dim fStream
	
	Function init()
	
		Set Me.stdOut = WScript.StdOut
		Set Me.stdErr = Wscript.StdErr
		
		Set Me.stream = StdOut
		
		Me.fStream = NULL
		
		Set init = Me
		
	End Function
	
	'---------------------------------------------------------------------------------
	Function addTeedFileStream(strResultLocation, strShadowLocation)
	
		Dim shadowStream
	
		Set shadowStream = New ShadowedFileOutputStreamWriter.init(strResultLocation, _
			strShadowLocation)
		
		Set Me.fStream = shadowStream
		
		Set stream = New psuedoTee.init(shadowStream)
		
	End Function
	
	'---------------------------------------------------------------------------------
	Function setFileStream(strResultLocation, strShadowLocation) 'Only to file
	
		Dim shadowStream
	
		Set shadowStream = New ShadowedFileOutputStreamWriter.init(strResultLocation, _
			strShadowLocation)
			
		Set Me.fStream = shadowStream
		
		Set stream = shadowStream
		
	End Function
	
	'---------------------------------------------------------------------------------
	Function getUpdateListString( objSearchResult )
		
		Dim searchList
		searchList = ""
		
		Dim i
		For i = 0 To ( objSearchResult.Updates.Count-1 )
			Dim update, updateLine
			Set update = objSearchResult.Updates.Item(i)
			
			updateLine = "{" & update.title & _
				"|KB=" & update.KBArticleIDs.Item(0) & _
				"|impact=" & update.installationBehavior.impact & _
				"|isDl=" & update.isDownloaded & _
				"|isInst=" & update.isInstalled & _
				"}"
			If (i = 0 ) Then
				searchlist = updateLine
			Else
				searchlist = searchList & VBCRLF & "  ," & updateLine
			End If
		Next
		
		getUpdateListString = searchlist
	End Function
	
	'---------------------------------------------------------------------------------
	Function recordSearchResult( objSearchResult )
		stream.writeLine( getPair( "search.result.count", _
			objSearchResult.Updates.Count) )
			
		stream.writeLine( getPair( "search.result.code", _
			getOperationResultMsg( objSearchResult.ResultCode) ) )
		
		getUpdateListString(objSearchResult)
		call stream.writeLine( getPair("search.result.list",_
			getUpdateListString(objSearchResult)) )
	End Function
	
	'---------------------------------------------------------------------------------
	Function recordUpdateResult( objUpdates, objResults, strType )
		stream.writeLine( getPair( strType & ".result.count", _
			objUpdates.Count ) )
		stream.writeLine( getPair( strType & ".result.code", _
			getOperationResultMsg(objResults.ResultCode ) ) )
		stream.writeLine( getPair( strType & ".result.HResult", _
			hex( objResults.HResult ) ) ) 
		
		Dim dlList
		dlList = ""
		
		Dim i
		For i = 0 To ( objUpdates.Count-1 )
			Dim update, updateLine, dlLine
			Set update = objUpdates.Item(i)

			dlLine = "{" & update.title & _
				"|KB=" & update.KBArticleIDs.Item(0) & _
				"|Res=" & getOperationResultMsg(objResults.GetUpdateResult(i).ResultCode) & _
				"|HResult=0x" & hex( objResults.GetUpdateResult(i).HResult ) & "}"
			If (i = 0 ) Then
				dlList = dlLine
			Else
				dlList = dlList & VBCRLF & "  ," & dlLine
			End If
		Next
		
		call stream.writeLine( getPair( strType & ".result.list", dlList ) )
	End Function
	
	'---------------------------------------------------------------------------------
	Function recordDownloadResult( objUpdates, objDownloadResults )
		call recordUpdateResult( objUpdates,  objDownloadResults, "download" )
	End Function
	
	'---------------------------------------------------------------------------------
	Function recordDownloadFailure( strReason )
		stream.writeLine(getPair("download.failure!.reason",strReason))
	End Function
	
	'---------------------------------------------------------------------------------
	Function recordInstallFailure( strReason )
		stream.writeLine(getPair("install.failure!.reason",strReason))
	End Function
	
	'---------------------------------------------------------------------------------
	Function refreshDownloadStatus( objDlProgress, objUpdates )
		reWrite("                                                                                              ")
		reWrite( "download.status" & getTotalUpdateDownloadProgress(objDlProgress,objUpdates) & _
			":" & getCurrentUpdateDownloadProgress(objDlProgress,objUpdates) )
	End Function
	
	'---------------------------------------------------------------------------------
	Function refreshInstallStatus( objInstallProgress, objUpdates )
		reWrite("                                                                                              ")
		reWrite( "install.status" & getTotalUpdateInstallProgress(objInstallProgress,objUpdates) & _
			":" & getCurrentUpdateInstallProgress(objInstallProgress,objUpdates) )
	End Function

	'---------------------------------------------------------------------------------
	Function recordDownloadStatus( objDlProgress, objUpdates )
		reWrite("                                                                                              ")
		stream.writeLine("")
		stream.writeLine( "#download.status" & getTotalUpdateDownloadProgress(objDlProgress,objUpdates) & _
			":" & getCurrentUpdateDownloadProgress(objDlProgress,objUpdates) )
	End Function
	
	'---------------------------------------------------------------------------------
	Function recordInstallStatus( objInstallProgress, objUpdates )
		reWrite("                                                                                              ")
		stream.writeLine("")
		stream.writeLine( "#install.status" & getTotalUpdateInstallProgress(objInstallProgress,objUpdates) & _
			":" & getCurrentUpdateInstallProgress(objInstallProgress,objUpdates) )
	End Function
	
	'---------------------------------------------------------------------------------
	Function recordInstallationResult( objUpdates, objInstallationResults )
		stream.writeLine( getPair("install.reboot_required", _
			objInstallationResults.RebootRequired ) )
			
		call recordUpdateResult( objUpdates, objInstallationResults, "install" )
	End Function

	'---------------------------------------------------------------------------------
	Function recordMissingDownloads(intInstalled)
		stream.writeLine( getPair( "install.pre.dls.missing", intInstalled ) )
	End Function
	
	'---------------------------------------------------------------------------------
	Function recordShutdownPlan(booIsShutdownPlanned)
		stream.writeLine(getPair("post.rebootPlanned", booIsShutdownPlanned))
	End Function
	
	'---------------------------------------------------------------------------------
	Function recordComplete()
		stream.writeLine( getPair( "post.complete_time", Now() ) )
	End Function
	
	'---------------------------------------------------------------------------------
	Function recordError( strMessage )
		stream.writeLine( "#error:" & strMessage )
	End Function
	
	'---------------------------------------------------------------------------------
	Function recordInfo( strMessage )
		stream.writeLine( "#info:" & strMessage )
	End Function
	
	'---------------------------------------------------------------------------------
	Function writeTitle(strName, strVersion)
		stream.writeLine("#" & strName & " " & strVersion & " " & Now())
	End Function
	
	'---------------------------------------------------------------------------------
	Function writeId( strRunId )
		stream.writeLine( getPair("init.ruid", strRunId) )
	End Function
	
	'---------------------------------------------------------------------------------
	Function recordPendingShutdown(strPendingShutdownAction)
		stream.writeLine( getPair("pre.restart.required", strPendingShutdownAction ) )
	End Function
	
	'---------------------------------------------------------------------------------
	Function getPair(strKey, strVal)
		getPair = strKey & ":" & strVal
	End Function
	
	'---------------------------------------------------------------------------------
	Function getDownloadPhase(intPhase)
		Select Case intPhase
		  Case 1
			getDownloadPhase = "Initializing"
		  Case 2
			getDownloadPhase = "Downloading"
		  Case 3
			getDownloadPhase = "Verifying"
		  Case Else
			getDownloadPhase = "?"
		End Select
	End Function

	'---------------------------------------------------------------------------------
	Function getTotalUpdateDownloadProgress(downloadProgress,updates)
		Dim kbDown
		kbDown = (cLng(downloadProgress.TotalBytesDownloaded) / 1000)
		
		Dim kbTotal
		kbTotal = (cLng(downloadProgress.TotalBytesToDownload) / 1000)

		getTotalUpdateDownloadProgress = "(" & kbTotal & _
			"/" & kbDown & ")[" & downloadProgress.percentComplete & "]"
	End Function

	'---------------------------------------------------------------------------------
	Function getCurrentUpdateDownloadProgress(downloadProgress,updates)
		Dim dp
		Set dp = downloadProgress
		
		Dim currentUpdate
		Set currentUpdate = updates.item(dp.currentUpdateIndex)
		
		Dim currentUpdateKb
		'There is almost always just one KB
		currentUpdateKb = currentUpdate.KBArticleIDs.Item(0) 
		
		Dim dlSize
		dlSize = cLng(dp.currentUpdateBytesToDownload) / 1000
		
		Dim dlDone
		dlDone = cLng(dp.currentUpdateBytesDownloaded) / 1000
		
		Dim dlPhase
		dlPhase = getDownloadPhase(dp.CurrentUpdateDownloadPhase)
		
		Dim dlPct
		dlPct = dp.CurrentUpdatePercentComplete
		
		getCurrentUpdateDownloadProgress = "{" & currentUpdateKb & "-" & _
			dlPhase & "}(" & dlSize & "/" & dlDone & ")[" & dlPct & "]"
		
	End Function

	'---------------------------------------------------------------------------------
	Function getTotalUpdateInstallProgress(InstallProgress,updates)
		getTotalUpdateInstallProgress = "[" & InstallProgress.percentComplete & "]"
	End Function

	'---------------------------------------------------------------------------------
	Function getCurrentUpdateInstallProgress(InstallProgress,updates)
		Dim ip
		Set ip = InstallProgress
		
		Dim currentUpdate
		Set currentUpdate = updates.item(ip.currentUpdateIndex)
		
		Dim currentUpdateKb
		'There is almost always just one KB
		currentUpdateKb = currentUpdate.KBArticleIDs.Item(0) 
		
		Dim ipPct
		ipPct = ip.CurrentUpdatePercentComplete
		
		getCurrentUpdateInstallProgress = "{" & currentUpdateKb & "}[" & ipPct & "]"
		
	End Function
	
	'---------------------------------------------------------------------------------
	Function getOperationResultMsg(intResultCode)
		Dim strResult
		If intResultCode = 0 Then 
			strResult = "Not Started"
		ElseIf intResultCode = 1 Then 
			strResult = "In progress"
		ElseIf intResultCode = 2 Then 
			strResult = "Succeded"
		ElseIf intResultCode = 3 Then 
			strResult = "Succeeded with Errors"
		ElseIf intResultCode = 4 Then 
			strResult = "Failed"
		ElseIf intResultCode = 5 Then 
			strResult = "Aborted"			
		End If
		
		getOperationResultMsg = strResult
	End Function
	
	'---------------------------------------------------------------------------------
	' Do not output this to the tee, unless you want a huge pointless result file.
	Private Function reWrite(strMessage)
		stdOut.write chr(13) & strMessage
	End Function
	
	'---------------------------------------------------------------------------------
	Sub Class_Terminate()
		Set stream = nothing
	End Sub
End Class

'===============================================================================
'===============================================================================
Class ShadowedFileOutputStreamWriter
	Dim strOutputLocation
	Dim strShadowLocation
	Dim fStream
	Dim fShadow
	Dim fso
	
	'Custom Constructor - Take filename and shadow location
	Function init(strOutputLocation, strShadowLocation)
	
		
		Me.strOutputLocation = strOutputLocation
		Me.strShadowLocation = strShadowLocation
		
		Set fso = CreateObject("Scripting.FileSystemObject")
		
		Set fShadow = tryCreateFile(strShadowLocation)
		Set fStream = tryCreateFile(strOutputLocation)
		
		Set init = Me
	End Function
	
	Function tryCreateFile(strOutputLocation)
		On Error Resume Next
			Set tryCreateFile = fso.createTextFile( strOutputLocation, True )
		e.catch() 'catch
		On Error GoTo 0
		If (e.isException()) Then
			Dim Ex
			Set Ex = e.getException()
			Dim newEx
			Set newEx = e.preRaise( New ErrWrap.initExM( WUF_STREAM_ERROR, _
				"ShadowedFileOutputStreamWriter.tryCreateFile()",_
				"Unable to write to file: " & strOutputLocation , Ex) )
			Err.Raise newEx.number, newEx.Source, newEx.Description
		End If
	End Function
	
	Function writeLine(strMessage)
		fShadow.writeLine(strMessage)
		fStream.writeLine(strMessage)
	End Function
	
	Function write(strMessage) 
		fShadow.writeLine(strMessage)
		fStream.writeLine(strMessage)
	End Function
	
	Function close()
		If (isObject(fStream)) Then
			fStream.close
			fStream = NULL
		End If
		If (isObject(fShadow)) Then
			fShadow.close
			fShadow = NULL
		End If
		
		If ( checkFile(strOutputLocation) ) Then
			If ( fso.fileExists(strShadowLocation) ) Then
				fso.DeleteFile strShadowLocation
			End If
		Else
			logError ( "Unable to verify output file, leaving shadow at: " _
				& strShadowLocation )
		End If
	End Function
	
	Function getLocation()
		getLocation = strOutputLocation
	End Function
	
	Function isUsingFile()
		isUsingFile = Not booConsoleOnly
	End Function
	
	Function checkFile(strLocation)
		checkFile = fso.FileExists(strLocation)
	End Function
	
	Sub Class_Terminate()
		'close()
	End Sub
End Class

'===============================================================================
'===============================================================================
' Error Handling ---------------------------------------------------------------
' This section supports try-catch&throw functionality in vbscript.
' You should only surround one exception throwing command with this
' construct, otherwise you might lose the error.
' The usage idiom for a tr-catch-throw is:
' On Error Resume Next 'try
' 	... 'code that could throw exception
' Set Ex = e.catch() 'catch
' On Error GoTo 0 'catch part two
' If (Ex = <some_err_num>) Then
' 	... 'Handle error
'	Set newEx = New ErrWrap.initExM(<somenum>,"<source>", "<description>", Ex)'
'	e.preRaise(newEx)
'	Err.Raise newEx.number, newEx.source, newEx.description
' End If

' Note that code called within an error handler that re-throws (using Err.raise)
' must be "exception raise safe" all the way up the call chain.
' If your called function has an "On Error..." statement in it, that will reset
' The global Err object, thereby losing the exception the code was handling. When
' The raise is called at the end of the handling to re-throw, it will throw an
' "non-error" Err object with code 0, which will then slip by any upstream
' error handlers. A nightmare to debug if it happens.

Class ErrWrap
	Private pNumber
	Private pSource
	Private pDescription
	Private pHelpContext
	Private pHelpFile
	Private objReasonEx
	
	Public Function catch()
		init()		
		objReasonEx = NULL
		Set catch = Me
	End Function
	
	Public Function init()
		pNumber = Err.Number
		pSource = Err.Source
		pDescription = Err.Description
		pHelpContext = Err.HelpContext
		pHelpFile = Err.HelpFile
		objReasonEx = NULL
		Set init = Me
	End Function
	
	Public Function initM(intCode, strSource, strDescription)
		pNumber = intCode
		pSource = strSource
		pDescription = strDescription
		pHelpContext = ""
		pHelpFile = ""		
		Set initM = Me
	End Function
	
	Public Function initExM(intCode, strSource, strDescription, objEx)
		pNumber = intCode
		pSource = strSource
		pDescription = strDescription
		pHelpContext = ""
		pHelpFile = ""		
		Set objReasonEx = objEx
		Set initExM = Me
	End Function
	
	Public Function initEx(objEx)
		pNumber = Err.Number
		pSource = Err.Source
		pDescription = Err.Description
		pHelpContext = Err.HelpContext
		pHelpFile = Err.HelpFile
		Set objReasonEx = objEx
		Set initEx = Me
	End Function
	
	Public Function getReason() 'returns objEx
		If NOT isObject(objReasonEx) Then
			getReason = NULL
		Else
			Set getReason = objReasonEx
		End If
	End Function
	
	Public Function toString() 'returns string
		toString = ""
		toString = "ExNum:[" & pNumber & " : 0x" & hex(pNumber) & "] " & VbCrLf & _
			"ExSource: " & pSource & VbCrLf & _
			"ExDescription: " & pDescription & VbCrLf
	End Function
	
	Public Default Property Get Number
		Number = pNumber
	End Property
	
	Public Property Get Source
		Source = pSource
	End Property
	
	Public Property Get Description
		Description = pDescription
	End Property
	
	Public Property Get HelpContext
		HelpContext = pHelpContextl
	End Property
	
	Public Property Get HelpFile
		HelpFile = HelpFile
	End Property
	
End Class

'===============================================================================
'===============================================================================
' Usage: declare this at the glocal scope.
Class ExceptionManager
	Dim currentEx
	
	Function init()
		currentEx = NULL
		Set init = Me
	End Function
	
	Function catch()
		If ( isNull(currentEx) ) Then
			If ( Err.number <> 0 ) Then
				Set currentEx = New ErrWrap.catch()
			End If
		Else
			If (Err.number <> currentEx.number) Then
				'Exception mismatch, when the current exception
				'does not match the last recorded currentEx.
				'Happens when an exception is thrown in an 
				'exception handlerl
				If ( Err.number <> 0 ) Then
					Set currentEx = New ErrWrap.initEx(currentEx)
				End IF
			End If
			catch = true
		End If
		
	End Function
	
	Function isException()
		isException = true
		If NOT ( isObject(currentEx) )Then
			isException = false
		End If
	End Function
	
	Function getException()
		Set getException = currentEx
		currentEx = NULL
	End Function
	
	' Do not use if you care about localizing the line number of the error
	' This function will set the error line to the Raise called within
	Function throw(objEx)
		Set currentEx = objEx
		Err.Raise currentEx.number, currentEx.Source, currentEx.Description
	End Function
	
	Function preRaise(objEx)
		Set currentEx = objEx
		Set preRaise = currentEx
	End Function
	
	Function dump(objEx)
		dump = ""
		If NOT (isObject(objEx.getReason)) Then
			dump = objEx.toString() & VbCrLf
		Else
			dump = dump(objEx.getReason) & objEx.toString() & VbCrLf
		End If
	End Function
End Class