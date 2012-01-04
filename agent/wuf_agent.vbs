Option Explicit
'*******************************************************************************
Const LOG_LEVEL = 3
  
Const LOG_LEVEL_DEBUG = 3
Const LOG_LEVEL_INFO = 2
Const LOG_LEVEL_WARN = 1
Const LOG_LEVEL_ERROR = 0

Const WUF_CATCH_ALL_EXCEPTIONS = 0

Const WUF_AGENT_VERSION = "1.8"
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

Const WUF_ASYNC = 1

Const WUF_INPUT_ERROR = 10001
Const WUF_FEEDBACK_ERROR = 10002

Const WUF_ACTION_UNDEFINED = -1
Const WUF_ACTION_AUTO = 	0
Const WUF_ACTION_SCAN = 	1
Const WUF_ACTION_DOWNLOAD = 2
Const WUF_ACTION_INSTALL = 	3

Const WUF_SHUTDOWN_UNDEFINED = -1
Const WUF_SHUTDOWN_DONT = 	0
Const WUF_SHUTDOWN_RESTART = 1
Const WUF_SHUTDOWN_SHUTDOWN = 2
Const WUF_SHUTDOWN_DELAY = 60

Const WUF_DEFAULT_SEARCH_FILTER = "IsAssigned=1 and IsHidden=0 and IsInstalled=0 and Type='Software'"
'Const WUF_DEFAULT_SEARCH_FILTER = "IsAssigned=1 and IsHidden=0 and IsInstalled=1 and Type='Software'"
Const WUF_DEFAULT_FORCE_SHUTDOWN_ACTION = false
Const WUF_DEFAULT_ACTION = 0 
Const WUF_DEFAULT_SHUTDOWN_OPTION = 0

Const WUF_DEFAULT_LOG_LOCATION = "C:\Windows\Temp\wufa_local" 
Const WUF_DEFAULT_RESULT_LOCATION = "C:\Windows\Temp\wufa_local_result" 

Const WUF_DEFAULT_RESULT_DROPBOX = "c:\temp\wuf_dropbox"
Const WUF_DEFAULT_LOG_DROP_NAME = "wufa_drop_localhost.log"
Const WUF_DEFAULT_RESULT_DROP_NAME = "wufa_drop_result_localhost.txt"

Const WUF_USAGE = "wuf_agent [ action:(AUTO|SCAN|DOWNLOAD|INSTALL) shutdownoption:(0|1|2) force:(1|0) dropbox:(<unc_path>) resultdropname:(name)"

'Globals - avoid modification after initialize()
Dim stdErr, stdOut	'std stream access
Dim gWshShell		'Shell access
Dim gWshSysEnv		'Env access
Dim gLogLocation	'Log location
Dim gResultLocation	'Result location
Dim gAction			'This applications action
Dim gShutdownOption	'Restart, shutdown, or do nothing
Dim gForceShutdown	'Do the shutdown option even if not required
Dim gFileLog 		'Wuf Agent Log object
Dim gResultMap		'Holds results until ready to write to file
Dim gDropBox		'Where results are sent after completion
Dim gResultDropName	'Name of result file sent to dropbox
Dim gLogDropName	'Name of log file sent to dropbox
Dim gRunId			'Unique id of run
Dim gObjUpdateSession 'The windows update session used for all wu operations
Dim gObjDummyDict	'Used for async wu operations

Class DummyClass 'set up dummy class for async download and installation calls
	Public Default Function DummyFunction()
		WScript.echo "Dummy Function Called"
	End Function		
End Class

' main injection point
main()

'*******************************************************************************
Function main()
	
	If (WUF_CATCH_ALL_EXCEPTIONS = 1) Then
		On Error Resume Next
			core()
		Dim Ex
		Set Ex = New ErrWrap.catch() 'catch
		On Error GoTo 0
		If (Ex.number <> 0) Then
			call logErrorEx("Unexpected exception.", Ex)
		End If
	Else
		core()
	End If
	
	WScript.quit
End Function

'*******************************************************************************
Function core()
	initialize()
	If  (configure() = true) Then
		If (verify() = true) Then
			preAction()
			doAction(gAction)
			postAction()
			feedBack()
		End If
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
	Set gResultMap = CreateObject("Scripting.Dictionary")
	
	Set gObjDummyDict = CreateObject("Scripting.Dictionary")
	
	Call gObjDummyDict.Add("DummyFunction", New DummyClass)
	
	gLogLocation = WUF_DEFAULT_LOG_LOCATION & "_" & gRunId & ".log"
	gResultLocation = WUF_DEFAULT_RESULT_LOCATION  & "_" & gRunId & ".txt"
	gAction = WUF_DEFAULT_ACTION
	gShutdownOption = WUF_DEFAULT_SHUTDOWN_OPTION
	gForceShutdown = WUF_DEFAULT_FORCE_SHUTDOWN_ACTION
	gDropBox = WUF_DEFAULT_RESULT_DROPBOX
	gResultDropName = Replace(WUF_DEFAULT_RESULT_DROP_NAME, "localhost", getComputerName())
	gLogDropName = Replace(WUF_DEFAULT_LOG_DROP_NAME, "localhost", getComputerName())


End Function

'*******************************************************************************
Function configure()

	configure = true
	
	configureLogFile(gLogLocation)
	
	logInfo( "WUF Agent " & WUF_AGENT_VERSION )
	logInfo( "Log system initialized." )
	
	logInfo( "Run Id: " & gRunId )
	
	logDebug( "Parsing Configuration" )
	
	configureResultMap(gResultMap)
	
	logDebug("Creating Update Session.")
	Set gObjUpdateSession = CreateObject("Microsoft.Update.Session")
	
	On Error Resume Next
		parseArgs()
	Dim Ex
	Set Ex = New ErrWrap.catch() 'catch
	On Error GoTo 0
	If (Ex.number <> 0) Then
		If (Ex.number = WUF_INPUT_ERROR) Then
			call logErrorEx("Improper input.", Ex)
			logInfo("Usage: " & WUF_USAGE)
			configure = false
		Else
			call logErrorEx("Unknown error during configuration.", Ex)
			configure = false
			Exit Function
		End If
	End If
	
	logInfo( "DropBox: " & gDropBox)
	logInfo( "Result Drop Filename: " & gResultDropName)
	
End Function

'*******************************************************************************
Function configureResultMap(objResultMap)

	call writeGlobalResultMap( "init", "1" )
	call writeGlobalResultMap( "init.time", Now() )
	call writeGlobalResultMap( "init.ruid", gRunId )
	
End Function

'*******************************************************************************
Function parseArgs()

    Dim objArgs
	Dim success
	
	success = false
	
	Set objArgs = WScript.Arguments
	
	If (objArgs.Count > 0) Then
		Dim i
		For i = 0 to (objArgs.Count - 1)
			Dim strArrTemp
			If (objArgs.Count > 0) Then		  
				If ( headStrI(objargs(i),"action:") ) Then
					strArrTemp = split(objargs(i),":")
					If Not(parseAction(strArrTemp(1))) Then
						Err.Raise WUF_INPUT_ERROR, "Wuf.parseArgs()", "Invalid action: " & objargs(i)	
					End If
				ElseIf ( headStrI(objargs(i),"shutdownoption:") ) Then 
					strArrTemp = split(objargs(i),":")
					If Not( parseShutdownOption(strArrTemp(1)) ) Then
						Err.Raise WUF_INPUT_ERROR, "Wuf.parseArgs()", "Invalid shutdown option."
					End If
				ElseIf ( headStrI(objargs(i),"force:") ) Then
					strArrTemp = split(objargs(i),":")
					If Not( parseForceShutdown(strArrTemp(1)) ) Then
						Err.Raise WUF_INPUT_ERROR, "Wuf.parseArgs()", "Invalid force shutdown option."
					End If
				ElseIf ( headStrI(objargs(i),"dropbox:") ) Then
					strArrTemp = split(objargs(i),"box:")
					gDropBox = strArrTemp(1)
				ElseIf ( headStrI(objargs(i),"resultDropName:") ) Then
					strArrTemp = split(objargs(i),":")
					gResultDropName = strArrTemp(1)				
				Else
					success = false
					Err.Raise WUF_INPUT_ERROR, "Wuf.parseArgs()", "Unknown argument: " & objargs(i)	
				End If
			End If
		Next
	Else
		success = true
	End If
End Function

'*******************************************************************************
Function parseAction(strArgVal) 'return boolean
	parseAction = True
	If ( strCompI(strArgVal,"auto") ) Then
		gAction = WUF_ACTION_AUTO
	ElseIf ( strCompI(strArgVal,"scan") ) Then
		gAction = WUF_ACTION_SCAN
	ElseIf ( strCompI(strArgVal,"download") ) Then
		gAction = WUF_ACTION_DOWNLOAD
	ElseIf ( strCompI(strArgVal,"install") ) Then
		gAction = WUF_ACTION_INSTALL
	Else
		gAction = WUF_ACTION_UNDEFINED
		parseAction = False
	End If
End Function

'*******************************************************************************
Function parseShutdownOption(strArgVal) 'return boolean
	parseShutdownOption = True
	If (strCompI(strArgVal,"0")) Then
		gShutdownOption = WUF_SHUTDOWN_DONT
	ElseIf (strCompI(strArgVal, "1")) Then
		gShutdownOption = WUF_SHUTDOWN_RESTART
	ElseIf (strCompI(strArgVal, "2")) Then
		gShutdownOption = WUF_SHUTDOWN_SHUTDOWN
	Else
		gShutdownOption = WUF_SHUTDOWN_UNDEFINED
		parseShutdownOption = False
	End If
End Function

'*******************************************************************************
Function parseForceShutdown(strArgVal) 'return boolean
	parseForceShutdown = True
	If (strCompI(strArgVal,"0")) Then
		gForceShutdown = False
	ElseIf (strCompI(strArgVal,"1")) Then
		gForceShutdown = True
	Else
		parseForceShutdown = False
	End If
End Function

'*******************************************************************************
Function verify() 'return boolean
	Dim verified
	
	logInfo( "---Verifying Configuration...---" )
	verified = True
	
	If (isCscript()) Then
		logInfo( "[+] cscript used." )
	Else
		logError( "[-] Not using cscript to run wuf, failed." )
	End If
	
	If (checkUpdateAgent()) Then
		logInfo( "[+] Windows Update Agent is up to date." )
	Else
		logError( "[-] Windows Update Agent is out of date, failed check." )
		verified = False
	End If
	
	If (isShutdownActionPending()) Then
		logInfo("[?] There is a pending restart required.")
		If (gAction = WUF_ACTION_INSTALL) Then
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
	recordPreConditions()
	logInfo("Pre-Action Complete.")
End Function

'*******************************************************************************
Function doAction(intAction)

	logInfo("Performing Action.")
	
	Dim objUpdateResults
	
	If (intAction = WUF_ACTION_AUTO) Then
		autoDetect()
	Else
		Set objUpdateResults = manualAction(intAction)
	End If
	
	logInfo("Action Complete.")
	
End Function

'*******************************************************************************
Function manualAction(intAction)
	Dim searchResults
	Dim downloadResults
	Dim installResults
	Dim intUpdateCount
	
	logDebug("Starting Manual Action.")
	
	Set searchResults = wuSearch( WUF_DEFAULT_SEARCH_FILTER )
	intUpdateCount = searchResults.Updates.Count
	
	logUpdateSearchResult( searchResults )
	recordSearchResult( searchResults )
	
	If (intUpdateCount > 0) Then
		acceptEulas(searchResults)
		If (intAction = WUF_ACTION_DOWNLOAD) Then
			If (WUF_ASYNC = 1) Then
				Set downloadResults = wuDownloadAsync(searchResults)
			Else
				Set downloadResults = wuDownload(searchResults)
			End IF
			call logDownloadResult(searchResults.updates, downloadResults)
			call recordDownloadResult(searchResults.updates, downloadResults)
		ElseIf (intAction = WUF_ACTION_INSTALL) Then
			If (WUF_ASYNC = 1) Then
				Set downloadResults = wuDownloadAsync(searchResults)
			Else
				Set downloadResults = wuDownload(searchResults)
			End IF
			call logDownloadResult(searchResults.updates, downloadResults)
			call recordDownloadResult(searchResults.updates, downloadResults)
			If (WUF_ASYNC = 1) Then
				Set installResults = wuInstallAsync(searchResults)
			Else
				Set installResults = wuInstall(searchResults)
			End IF
			call recordInstallationResult(searchResults.updates, installResults )
			call logInstallationResult(searchResults.updates,installResults)
		End If
	End If
	
	logDebug("Manual Action completed.")
	Set manualAction = searchResults
End Function

'*******************************************************************************
Function recordPreConditions()
	call writeGlobalResultMap("pre.restart.required", _
		isShutdownActionPending())
End Function

'*******************************************************************************
Function recordSearchResult( objSearchResults )
	call writeGlobalResultMap("search.result.count", _
		objSearchResults.Updates.Count)
		
	call writeGlobalResultMap( "search.result.code", _
		getOperationResultMsg(objSearchResults.ResultCode) )
	
	Dim searchList
	searchList = ""
	
	Dim i
	For i = 0 To ( objSearchResults.Updates.Count-1 )
		Dim update, updateLine
		Set update = objSearchResults.Updates.Item(i)

		updateLine = update.title & "|" & update.uniqueid
		If (i = 0 ) Then
			searchlist = updateLine
		Else
			searchlist = searchList & VBCRLF & "  ," & updateLine
		End If
	Next
	
	call writeGlobalResultMap( "search.result.list", searchList)
End Function

'*******************************************************************************
Function recordDownloadResult( objUpdates, objDownloadResults )
	call writeGlobalResultMap("download.result.count", _
		objUpdates.Count)
	call writeGlobalResultMap( "download.result.code", _
		getOperationResultMsg(objDownloadResults.ResultCode) )
	
	Dim dlList
	dlList = ""
	
	Dim i
	For i = 0 To ( objUpdates.Count-1 )
		Dim update, updateLine, dlLine
		Set update = objUpdates.Item(i)

		dlLine = update.title & "|" & update.uniqueid & "|" & _
			getOperationResultMsg(objDownloadResults.GetUpdateResult(i).ResultCode)
		If (i = 0 ) Then
			dlList = dlLine
		Else
			dlList = dlList & VBCRLF & "  ," & dlLine
		End If
	Next
	
	call writeGlobalResultMap( "download.result.list", dlList)
End Function

'*******************************************************************************
Function recordInstallationResult( objUpdates, objInstallationResults )
	call writeGlobalResultMap("install.result.count", _
		objUpdates.Count)
		
	call writeGlobalResultMap( "install.result.code", _
		getOperationResultMsg(objInstallationResults.ResultCode) )
	call writeGlobalResultMap( "install.reboot_required", _
		objInstallationResults.RebootRequired )
	
	Dim instList
	instList = ""
	
	Dim i
	For i = 0 To ( objUpdates.Count-1 )
		Dim update, updateLine, instLine
		Set update = objUpdates.Item(i)

		instLine = update.title & "|" & update.uniqueid & "|" & _
			getOperationResultMsg(objInstallationResults.GetUpdateResult(i).ResultCode)
		If ( i = 0 ) Then
			instList = instLine
		Else
			instList = instList & VBCRLF & "  ," & instLine
		End If
	Next
	
	call writeGlobalResultMap( "install.result.list", instList)
End Function

'*******************************************************************************
Function recordPostResult()
	call writeGlobalResultMap("post.rebootPlanned", rebootPlanned())
	call writeGlobalResultMap("post.complete_time", Now())
End Function

'*******************************************************************************
Function postAction()
	logInfo("Performing post-actions")
	If (rebootPlanned()) Then
		logInfo("System shutdown action will occur.")
		call shutDownActionDelay(gShutdownOption, WUF_SHUTDOWN_DELAY)
	End If
	recordPostResult()
	logInfo("Completed post-actions")
End Function

'*******************************************************************************
Function feedBack()
	call writeResultMapFile(gResultMap, gResultLocation)
	call sendResults()
	'call sendLog()
End Function

'*******************************************************************************
Function cleanup()
	'@@TODO: Add info here.
	logInfo("Cleaning up")
	logInfo("WUF finished.")
	gFileLog.close
End Function

'**************************************************************************************
Function wuSearch(strFilter) 'return ISearchResult
	Dim searchResult
	Dim updateSearcher 
	Dim blnFatal
	
	logDebug("Creating Update Searcher.")
	Set updateSearcher = gObjUpdateSession.CreateUpdateSearcher()
	
	logInfo("Starting Update Search.")
	
	Dim caughtErr
	On Error Resume Next
		Set searchResult = updateSearcher.Search(strFilter)
	Set caughtErr = New ErrWrap.catch() 'catch
	On Error GoTo 0
	If (caughtErr.number <> 0) Then
		Dim strMsg
		If (caughtErr = cLng("&H80072F78") ) Then
			strMsg = "ERROR_HTTP_INVALID_SERVER_RESPONSE - The server response could not be parsed."
		ElseIf (caughtErr = cLng("&H8024402C") ) Then
			strMsg = "WU_E_PT_WINHTTP_NAME_NOT_RESOLVED - Winhttp SendRequest/ReceiveResponse failed with 0x2ee7 error. Either the proxy " _
			& "server or target server name can not be resolved. Corresponds to ERROR_WINHTTP_NAME_NOT_RESOLVED. " 
		ElseIf (caughtErr = cLng("&H80072EFD") ) Then 
			strMsg = "ERROR_INTERNET_CANNOT_CONNECT - The attempt to connect to the server failed."
		ElseIf (caughtErr = cLng("&H8024401B") ) Then 
			strMsg = "SUS_E_PT_HTTP_STATUS_PROXY_AUTH_REQ - Http status 407 - proxy authentication required" 
		ElseIf (caughtErr = cLng("&H8024002B") ) Then 
			strMsg = "WU_E_LEGACYSERVER - The Sus server we are talking to is a Legacy Sus Server (Sus Server 1.0)"
		ElseIf (caughtErr = cLng("&H80244018") ) Then 
			strMsg = "SUS_E_PT_HTTP_STATUS_FORBIDDEN HttProxy Status 403"
		ElseIf (caughtErr = cLng("&H80072F8F") ) Then 
			strMsg = "ERROR_INTERNET_SECURE_FAILURE ErrorClockWrong - Unable to establish secure connection due to clock sync issue"
		ElseIf (caughtErr = 7) Then 
			strMsg = "Out of memory - In most cases, this error will be resolved by rebooting the client." 
		Else
			Dim strAddr
			strAddr = "http://msdn.microsoft.com/en-us/library/windows/desktop/ms681381(v=vs.85).aspx "
			strMsg = "Unknown problem searching for updates, refer to " & strAddr & "to look up error number." 
			blnFatal = true
		End If
		call logErrorEx(strMsg,caughtErr)
		Err.Raise caughtErr.Number, caughtErr.Source, caughtErr.Description
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
	
	On Error Resume Next
		Set objDownloadResult = downloader.Download()
	Dim Ex
	Set Ex = New ErrWrap.catch() 'catch
	On Error GoTo 0
	If (Ex.number <> 0) Then
		If (Ex.number = cLng("&H80240024")) Then
			call logInfoEx( "No updates to download.", Ex )
		Else
			call logErrorEx( "Could not download updates.", Ex )
			Err.Raise Ex.Number, Ex.Source & "; wuDownload()", Ex.Description
		End If
	End if
	
	
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
	
	logDebug("Creating update downloader.")
	
	Set downloader = gObjUpdateSession.CreateUpdateDownloader() 
	downloader.Updates = objSearchResult.Updates
	
	logInfo("Downloading Updates Asynchronously")
	
	Set dlJob = downloader.beginDownload(gObjDummyDict.Item("DummyFunction"),gObjDummyDict.Item("DummyFunction"),vbNull)
	set dlProgress = dlJob.getProgress()
	
	While Not getAsyncWuOpComplete(downloader.Updates, dlProgress) 
		set dlProgress = dlJob.getProgress()
		WScript.Sleep(10000)
		logInfo( "Download Progress: " & dlProgress.percentcomplete & "%" )
	Wend
	If (dlJob.isCompleted = TRUE) Then 
		logInfo("Asynchronous download call completed." )
	Else
		logError( "Could not complete asynchronous call." )
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
Function wuInstall(objSearchResult)
	Dim caughtErr

	Dim updatesToInstall
	Set updatesToInstall = CreateObject("Microsoft.Update.UpdateColl")
	
	Dim i
	For i = 0 To objSearchResult.Updates.Count-1
		Dim update
	    Set update = objSearchResult.Updates.Item(i)
	    If (update.IsDownloaded) Then
			logInfo("Update has been downloaded: " & update.Title )
			updatesToInstall.Add(update)
		Else
			logWarn("Could not complete download of: " & update.Title )
	    End If
	Next
	
	logDebug("Creating Update Installer.")
	Dim installer
	Set installer = gObjUpdateSession.CreateUpdateInstaller()
	installer.AllowSourcePrompts = False

	On Error Resume Next
		installer.ForceQuiet = True 
	Set caughtErr = New ErrWrap.catch() 'catch
	On Error GoTo 0
	If (caughtErr.number <> 0) Then
		call logErrorEx("Could not force installer to be quiet.", Err)
	End If
	
	installer.Updates = updatesToInstall
	
	logInfo("Installing updates.")
	
	Dim installationResult
	On Error Resume Next	
		Set installationResult = installer.Install()
	Set caughtErr = New ErrWrap.catch() 'catch
	On Error GoTo 0	
	If (caughtErr.number <> 0) then
		If (err.number = cLng(&H80240024)) then
			logInfo( "No updates to install." )
		Else
			call logErrorEx("Could not install updates.", caughtErr)
			Err.Raise caughtErr.Number, caughtErr.Source & "; wuInstall()", caughtErr.Description
		End If
	End if
	
	If Not( isObject(installationResult) ) Then
		wuInstall = null
	Else
		Set wuInstall = installationResult
	End If
	
End Function

'*******************************************************************************
Function wuInstallAsync(objSearchResult)
	Dim caughtErr
	Dim installJob, installProgress
	Dim objInstallResult

	Dim updatesToInstall
	Set updatesToInstall = CreateObject("Microsoft.Update.UpdateColl")
	
	Dim i
	For i = 0 To objSearchResult.Updates.Count-1
		Dim update
	    Set update = objSearchResult.Updates.Item(i)
	    If (update.IsDownloaded) Then
			logInfo("Update has been downloaded: " & update.Title )
			updatesToInstall.Add(update)
		Else
			logWarn("Could not complete download of: " & update.Title )
	    End If
	Next
	
	logDebug("Creating Update Installer.")
	Dim installer
	Set installer = gObjUpdateSession.CreateUpdateInstaller()
	installer.AllowSourcePrompts = False

	On Error Resume Next
		installer.ForceQuiet = True 
	Set caughtErr = New ErrWrap.catch() 'catch
	On Error GoTo 0
	If (caughtErr.number <> 0) Then
		call logErrorEx("Could not force installer to be quiet.", Err)
	End If
	
	installer.Updates = updatesToInstall
	
	logInfo("Installing updates.")
	
	Set installJob = installer.beginInstall(gObjDummyDict.Item("DummyFunction"),gObjDummyDict.Item("DummyFunction"),vbNull)
	set installProgress = installJob.getProgress()
	
	While Not getAsyncWuOpComplete(installer.Updates, installProgress) 
		set installProgress = installJob.getProgress()
		WScript.Sleep(10000)
		logInfo( "Install Progress: " & installProgress.percentcomplete & "%" )
	Wend
	If (installJob.isCompleted = TRUE) Then 
		logInfo("Asynchronous install call completed." )
	Else
		logError( "Could not complete asynchronous call." )
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
Function getAsyncWuOpComplete(objUpdates, objOperationProgress)

	Dim i, intTotalResultCode
	intTotalResultCode = 15
	
	For i = 0 To objUpdates.count - 1
		intTotalResultCode = intTotalResultCode AND objOperationProgress.getUpdateResult(i).resultCode
	Next
	
	If (intTotalResultCode = 0) Then
		getAsyncWuOpComplete = false
	Else 
		getAsyncWuOpComplete = true
	End If
	
End Function

'**************************************************************************************
Function logDownloadResult(objUpdates, objDownloadResult)

	If NOT (isObject(objDownloadResult) ) Then
		logInfo("No download result recorded.")
		Exit Function
	End If

	'Output results of install
	logInfo("Download Result Code: " & _
		getOperationResultMsg(objDownloadResult.ResultCode) )
	
	logInfo("Indvidual Update Download Results")
	Dim i
	For i = 0 to objUpdates.Count - 1
		Dim strResult
		strResult = getOperationResultMsg(objDownloadResult.GetUpdateResult(i).ResultCode)
		logInfo(objUpdates.Item(i).Title & ", " & objUpdates.Item(i).identity.updateId & ": " & strResult)
	Next
	
End Function

'**************************************************************************************
Function logInstallationResult(objUpdates, objInstallationResult)

	If NOT (isObject(objInstallationResult) ) Then
		logInfo("No installation result recorded.")
		Exit Function
	End If

	'Output results of install
	logInfo("Installation Result Code: " & objInstallationResult.ResultCode )
	logInfo("Reboot Required?: " & objInstallationResult.RebootRequired )
	
	logInfo("Indvidual Update Installation Results")
	Dim i
	For i = 0 to objUpdates.Count - 1
		Dim strResult
		strResult = getOperationResultMsg(objInstallationResult.GetUpdateResult(i).ResultCode)
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
	
	Dim caughtErr
	On Error Resume Next
		Set objAgentInfo = CreateObject("Microsoft.Update.AgentInfo") 
	Set caughtErr = New ErrWrap.catch() 'catch
	On Error GoTo 0
	If (caughtErr.number <> 0) Then
		logError( "Unable to get Agent Info object, perhaps windows updates haven't been configured?" )
		Err.Raise caughtErr.Number, caughtErr.Source & "; checkUpdateAgent()", caughtErr.Description
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
Function sendLog()
	logDebug("Attempting to send log")
	Dim strFileName
	strFileName = gLogDropName
	call sendFile(gLogLocation, gDropBox, strFileName )
End Function

'*******************************************************************************
Function sendResults()
	logDebug("Attempting to send results")
	Dim strFileName
	strFileName = gResultDropName
	call sendFile(gResultLocation, gDropBox, strFileName )
End Function

'*******************************************************************************
Function sendFile(strSourceLocation, strDestFolder, strDestFileName)
	Dim objFSO, objDestFile
	Dim strPath, strFullName
	Dim objFolder
	Dim objSourceFile
	Dim strMessage
	Dim strMsg
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")

	sendFile = false

	Set objSourceFile = objFSO.OpenTextFile (strSourceLocation, FORREADING, False, -2)
	
	Dim Ex
	
	On Error Resume Next' try
		strMessage = objSourceFile.readAll 
	Set Ex = New ErrWrap.catch() 'catch
	On Error GoTo 0
	If (Ex.number <> 0) Then
		
		strMsg = "Unable to read local file for sending to dropbox: " & strSourceLocation
		If (Ex.number = 62) Then
			call logWarnEx(strMsg, Ex)
		Else
			call logErrorEx(strMsg, Ex)
			Err.Raise Ex.number,  Ex.Source & "; sendFile()", Ex.Description & "; " & strMsg
		End If
	End If
	'End try-catch
	
	If Not objFSO.FolderExists(strDestFolder) Then
		On Error Resume Next 'try
			Set objFolder = objFSO.CreateFolder(strDestFolder)
		Set Ex = New ErrWrap.catch() 'catch
		On Error GoTo 0
		If (Ex.number <> 0) Then
			strMsg = "Drop box did not exist and could not be created: " & strDestFolder
			call logErrorEx( strMsg, Ex )
			Err.Raise Ex.number,  Ex.Source & "; sendFile()", Ex.Description & "; " & strMsg
		End If
		'End try-catch
	End If

	strPath = strDestFolder
	strFullName = objFSO.BuildPath(strPath, strDestFileName)
	
	On Error Resume Next 'try
		Set objDestFile = objFSO.OpenTextFile(strFullName,FORWRITING,True)
	Set Ex = New ErrWrap.catch() 'catch
	On Error GoTo 0
	If (Ex.number <> 0) Then
		strMsg = "Could not write file to dropbox: " & strDestFolder
		call logErrorEx( strMsg, Ex )
		Err.Raise Ex.number,  Ex.Source & "; sendFile()", Ex.Description & "; " & strMsg
	End If
	
	objDestFile.writeLine(strMessage)
	objDestFile.close
	objSourceFile.close
	sendFile = true

End Function

'*******************************************************************************
Function getDateTimeStamp()
	getDateTimeStamp = getDateStamp() & "_" & getTimeStamp()
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
	someTime = TimeValue(Now())
	someTime = Replace(someTime,":","")
	someTime = Replace(someTime," ","")
	getTimeStamp = someTime
End Function

'*******************************************************************************
Function getTimeStamp2()
	Dim someTime
	Dim sec, min, hr
	
	sec = right("0" & second(time),2)
	min = right("0" & minute(time),2)
	hr = right("0" & hour(time),2)
	someTime = hr & min & sec
	getTimeStamp2 = someTime
End Function

'*******************************************************************************
' A non zero delay is recommended so that this script can finish normally
Function shutDownActionDelay(intAction, intDelay)
	Dim strShutDown
	Dim objShell
	
	If (intAction = WUF_SHUTDOWN_RESTART) Then
		strShutdown = "shutdown.exe /r /t " & intDelay & " /f"
	ElseIf	(intAction = WUF_SHUTDOWN_SHUTDOWN) Then
		strShutdown = "shutdown.exe /s /t " & intDelay & " /f"
	End If
	
	Set objShell = CreateObject("WScript.Shell")
	
	objShell.Run strShutdown, 0, FALSE
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
	strLocalComputerName = gWshShell.ExpandEnvironmentStrings("%Computername%")
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

	Dim autoUpdateClient
	Dim autoUpdateSettings
	
	logInfo("Update Server: " & getWsusServer() )
	
	logDebug("Target Group: " & getTargetGroup() )
	
	Set autoUpdateClient = CreateObject("Microsoft.Update.AutoUpdate")
	
	Set autoUpdateSettings = autoUpdateClient.Settings
	
	logInfo("WUA Mode: " & getAutoUpdateNotificationLevelText(autoUpdateSettings.notificationlevel) )
	
	logInfo("WUA Schedule: " & getAuScheduleText() )

End Function

'*******************************************************************************
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

'**************************************************************************************
Function logUpdateSearchResult(objSearchResults)
	logInfo("Number of missing updates: " & objSearchResults.Updates.Count)

	logInfo("--- Missing Update List ---")
	Dim i
	For i = 0 To (objSearchResults.Updates.Count-1)
		Dim update, objCategories
		Set update = objSearchResults.Updates.Item(i)
		Set objCategories = objSearchResults.Updates.Item(i).Categories
		logInfo("Missing: " & objSearchResults.Updates.Item(i) )
		Dim j
		For j = 0 to objCategories.Count-1
		  logDebug("--Category: " & objCategories.Item(j).Description)
		Next
	Next
	logInfo("------------------------")
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
	logInfo("Action: " & getActionMessage(gAction))
	logInfo("Shutdown Option: " & shutdownOptionMessage() )
	logInfo("Force Shutdown Option: " & forceShutdownMessage() )
End Function

'**************************************************************************************
Function getActionMessage(intAction)
	If (intAction = WUF_ACTION_AUTO) Then
		getActionMessage = "Auto"
	ElseIf (intAction = WUF_ACTION_SCAN) Then
		getActionMessage = "Scan"
	ElseIf (intAction = WUF_ACTION_DOWNLOAD) Then
		getActionMessage = "Download"
	ElseIf (intAction = WUF_ACTION_INSTALL) Then
		getActionMessage = "Install"
	Else
		getActionMessage = "Unknown"
	End If
End Function

'**************************************************************************************
Function autoDetect()
	'Force WU Agent to detect
	Dim autoUpdateClient
	Set autoUpdateClient = CreateObject("Microsoft.Update.AutoUpdate")
	logDebug("Attempting to call Windows Auto Update DetectNow method.")
	On Error Resume Next
		autoUpdateClient.detectnow()
		If (err <> 0) then 
			call errorHandlerErr("Windows Update Auto Detection failed: ", true, err)
		End If
	on error goto 0 
End Function

'**************************************************************************************
Function acceptEulas(objSearchResults) 'return ISearchResult
	logDebug("Accepting EULAS on each update...")
	Dim i
	For i = 0 to objSearchResults.Updates.Count-1
		Dim update
		Set update = objSearchResults.Updates.Item(i) 
		If Not update.EulaAccepted Then update.AcceptEula 
	Next 
	Set acceptEulas = objSearchResults
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
		If (err <> 0) Then
			stdErr.writeLine "{LOG ERROR} Unable to write to log file. " & getFormattedErrorMsg( Err )
		End If
	On Error GoTo 0
	
	If (isCScript()) Then 
		If (intType <= LOG_LEVEL_WARN ) Then
			stdErr.writeLine strLine
		Else
			stdOut.writeLine strLine
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
	genRunId = getDateStamp() & "_" & getTimeStamp2() & "_" & genRandId(100)
End Function

'*******************************************************************************
Function setGlobalRunId()
	gRunId = genRunId()
End Function

'*******************************************************************************
Function writeGlobalResultMap(key,value)
	gResultMap.Add key, value
End Function

'*******************************************************************************
Function writeResultMapFile(objResultMap, strResultFileLocation)

	Dim fso
	Set fso = WScript.CreateObject( "Scripting.FileSystemObject" )
	
	Dim fileResult
	Set fileResult = fso.createTextFile( strResultFileLocation, True, -2 )
	
	Dim key
	For Each key In objResultMap
		fileResult.writeLine( key & ":" & gResultMap.Item(key) & ";" )
	Next
	
	fileResult.close()
End Function

'*******************************************************************************
Function rebootPlanned()
	rebootPlanned = ( (isShutdownActionPending() _
		AND ( gShutdownOption >= WUF_SHUTDOWN_RESTART )) _
		OR ( gForceShutdown ) )
End Function

'=============================================================
'=================================================================
' Error Handling -------------------------------------------------
' This section supports try-catch&throw functionality in vbscript.
' You should only surround one exception throwing command with this
' construct, otherwise you might lose the error.
' You must use this format in your code to simulate a try catch
' On Error Resume Next 'try
' 	... 'code that could throw exception
' Set caughtErr = New ErrWrap.catch() 'catch
' On Error GoTo 0 'catch part two
' If (caughtErr = <some_err_num>) Then
' 	... 'Handle error
' End If

Class ErrWrap
	Private pNumber
	Private pSource
	Private pDescription
	Private pHelpContext
	Private pHelpFile
	
	Public Function catch()
		pNumber = Err.Number
		pSource = Err.Source
		pDescription = Err.Description
		pHelpContext = Err.HelpContext
		pHelpFile = Err.HelpFile		
		Set catch = Me
	End Function
	
	Public Function Newk(strSource, ErrWrap)
		pNumber = ErrWrap.Number
		pSource = strSource & "->" & ErrWrap.Source
		pDescription = ErrWrap.Description
		pHelpContext = ErrWrap.HelpContext
		pHelpFile = ErrWrap.HelpFile
		Set Newk = Me
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
		HelpContext = pHelpContext
	End Property
	
	Public Property Get HelpFile
		HelpFile = HelpFile
	End Property
	
End Class

' Re-throws an existing error
' Limitation: The error line number will be set to a line within this function
' rendering the line number useless.
Function throw(wErr) 'returns nothing
	On Error GoTo 0
	Err.Raise wErr.number, wErr.source, wErr.description, wErr.HelpContext, wErr.HelpFile
End Function
'-----------------------------------------------------------------