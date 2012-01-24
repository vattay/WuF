@echo off

REM *******************************************************************************
REM Wuf Controller
REM Copyright (C) 2011 Anton Vattay

REM This program is free software: you can redistribute it and/or modify
REM it under the terms of the GNU General Public License as published by
REM the Free Software Foundation, either version 3 of the License, or
REM (at your option) any later version.

REM This program is distributed in the hope that it will be useful,
REM but WITHOUT ANY WARRANTY; without even the implied warranty of
REM MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
REM GNU General Public License for more details.

REM You should have received a copy of the GNU General Public License
REM along with this program.  If not, see <http://www.gnu.org/licenses/>.
REM *******************************************************************************

setlocal
REM ========================================================
REM User Variables
REM set dropBoxRootLocation="\\moore\public\wuf_dropbox"
set dropBoxRootLocation=c:\temp\wuf_dropbox
set dropResultPostfix=.result.txt

REM ========================================================
REM Constants
set actionString="AUTO" or "SCAN" or "DOWNLOAD" or "INSTALL"
set master_agent=wuf_agent.vbs
set usage=Usage^: %0% groupFile ^{Action^} [RESTART] [ATTACHED]
set usage2=	Action : ^( AUTO, SCAN, DOWNLOAD, INSTALL ^)

REM ========================================================
REM Date and Time ------------------------------------------

rem Parse the date (e.g., Fri 02/08/2008)
set cur_yyyy=%date:~10,4%
set cur_mm=%date:~4,2%
set cur_dd=%date:~7,2%

rem Parse the time (e.g., 11:17:13.49)
set cur_hh=%time:~0,2%
if %cur_hh% lss 10 (set cur_hh=0%time:~1,1%)
set cur_nn=%time:~3,2%
set cur_ss=%time:~6,2%
set cur_ms=%time:~9,2%

rem Set the timestamp format
set timestamp=%cur_yyyy%%cur_mm%%cur_dd%_%cur_hh%%cur_nn%%cur_ss%%cur_ms%


REM ========================================================
REM Conf Log ------------------------------------------
if not exist log (
	mkdir log
)
set log_file=log\controller-%timestamp%.log

REM ========================================================
REM Parse Args ---------------------------------------------

echo WUF Controller

if not [%5]==[] (	
	echo Too many arguments. >> %log_file% 2>&1
	goto inputerr
)

if [%1]==[] (
	echo You must provide a groupfile. >> %log_file% 2>&1
	goto inputerr
) else (
	set groupFile=%1
	set groupName=%~n1
)

if [%2]==[] (
	echo You must specify an action ^(%actionString%^). 2>&1
	goto inputerr
) else (
	set action_tag=%2
	if /I "%2"=="auto" (
		set action=/aA
		echo Action is AUTO. >> %log_file% 2>&1
	) else (
		if /I "%2"=="scan" (
			set action=/aS
			echo Action is SCAN. >> %log_file% 2>&1
		) else (
			if /I "%2"=="download" (
				set action=/aS /aD
				echo Action is DOWNLOAD. >> %log_file% 2>&1
			) else (
				if /I "%2"=="install" (
					set action=/aS /aI
					echo Action is INSTALL. >> %log_file% 2>&1
				) else (
					if /I "%2"=="di" (
						set action=/aS /aD /aI
						echo Action is DOWNLOAD and INSTALL. >> %log_file% 2>&1
					) else (
						echo Unknown action requested. >> %log_file% 2>&1
						goto inputerr
					)
				)
			)
		)
	)
)

set restart=
set attached=-d
set restart_tag=
if NOT [%3]==[] (
	if /I "%3"=="restart" (
		set restart=/sR
		set restart_tag=r
		echo Restart was specified. >> %log_file% 2>&1
	) else (
		if /I "%3"=="attached" (
			set attached=
			echo Attached mode was specified. >> %log_file% 2>&1
		) else (
			echo Unknown argument for restart action. >> %log_file% 2>&1
			goto inputerr
		)
	)
)

if NOT [%4]==[] (
	if /I "%4"=="restart" (
		set restart=/sR
		set restart_tag=r
		echo Restart was specified. >> %log_file% 2>&1
	) else (
		if /I "%4"=="attached" (
			set attached=
			echo Attached mode was specified. >> %log_file% 2>&1
		) else (
			echo Unknown argument for restart action. >> %log_file% 2>&1
			goto inputerr
		)
	)
)

REM ========================================================
REM Check Config
if not exist %groupFile% (
	echo The specified group file could not be found. >> %log_file% 2>&1
	goto :fatalerror
)

if not exist agent\%master_agent% (
	echo The agent file could not be found: %master_agent%. >> %log_file% 2>&1
	goto :fatalerror
)

if not exist %dropBoxRootLocation% (
	echo The drop box root location %dropBoxRootLocation% does not exist. >> %log_file% 2>&1
	goto :fatalerror
)

set restartConfirm=""
:restartConfirmTag
if /I "%restart%"=="/sR" Set /P restartConfirm="Are you sure you want to restart all servers in group[y/n]">CON

if /I "%restart%"=="/sR" (
	if /I "%restartConfirm%"=="y" (
		echo Restart was confirmed. >> %log_file% 2>&1
	) else (
		if /I "%restartConfirm%"=="n" (
			echo Restart was not confirmed^, quitting. >> %log_file% 2>&1
			goto wufEnd
		) else (
			goto restartConfirmTag
		)
	)
)

REM !!Pattern for dropbox name.
set dropBoxLocation="%dropBoxRootLocation%\%timestamp%_%groupName%_%action_tag%_%restart_tag%"

if not exist %dropBoxLocation% (
	mkdir %dropBoxLocation%
)

set msg="Drop box instance %dropBoxLocation% does not exist and could not be created."
if not exist %dropBoxLocation% (
	echo %msg% >> %log_file% 2>&1
	goto :fatalerror
) else (
	echo Drop box instance: %dropBoxLocation% >> %log_file% 2>&1
)

REM ========================================================
REM Configure

set remote_agent=local_%master_agent%

REM ========================================================
REM Do your thing

REM add stub return files
for /F "eol=;" %%i in (%groupFile%) do ( echo. 2>%dropBoxLocation%\%%i%dropResultPostfix% )

REM deploy the script
for /F "eol=;" %%i in (%groupFile%) do ( 
  ( echo Copying agent ^(agent\%master_agent%^) to %%i >> %log_file% 2>&1 ) 
  ( copy agent\%master_agent% \\%%i\C$\windows\temp\%remote_agent% >> %log_file% 2>&1 )  
  if not errorlevel 1 ( 
    ( echo Remote executing wuf agent on %%i >> %log_file% >> %log_file% 2>&1)
	( psexec %ATTACHED% -s \\%%i -w C:\Windows\Temp cscript.exe //NoLogo c:\windows\temp\%remote_agent% %action% /oN:%dropBoxLocation%\%%i%dropResultPostfix% /pA:%dropBoxLocation% %restart% 2>&1) 
  )
)

REM Immediatley deleting the agent after the detached psexec is started has unpredictable results so this 
REM attempts to delay the deletion.
ping -n 5 127.0.0.1 >nul

REM delete script
for /F "eol=;" %%i in (%groupFile%) do ( 
	( echo "Deleting agent on %%i" >> %log_file% 2>&1 ) 
	( del \\%%i\C$\windows\temp\%remote_agent% >> %log_file% 2>&1 )
)


:wufEnd
echo Wuf Finished > CON
echo Finished >> %log_file% 2>&1
goto :eof

:inputerr
echo ERROR in input.
echo %usage% 2>&1
echo %usage2% 2>&1
goto :eof

:fatalerror
echo Fatal error, quitting
goto :eof
