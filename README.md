# WuF (Windows Update Forcer)

## Introduction

The WuF is a utility that forces the windows update agent to perform actions that normally must be performed via the Windows Update website or the Windows Update Agent GUI. It can be used locally or remotely. It is intended for use in an environment where many windows computers must be updated in discrete batches and where any useful scripting languages and executables are not permitted.

The WuF system is composed of two parts, the controller (wuf_controller.bat) and the agent (wuf_agent.vbs). The agent is a standalone command line wrapper of the WUA API, which allows you to search, download, and install windows updates on a single computer.

The controller runs the agent on more than one computer at a time. The controller allows you to update batches of windows computers simultaneously. Feedback from the agents can be provided via command line or results can be dropped to a single network accessible location.

## Usage
Simply drop the WuF folder on a computer you intend to use as a command center. You will find two files:

wuf_controller.bat
agent\wuf_agent.vbs

## Controller

### Pre-requisites:
- Psexec.exe available on command computer.
- Psexec on PATH.
- Folder accessible to all target computers, usuall a network share.
- "Run as Administrator" on Windows Vista and 7.

The controller will perform windows update operations on a batch of computers. It maps a Windows Update action to a list of computer names. 

The available actions are auto, scan, download, install, and di. Auto simply calls the computers Windows Automatic Updates configured behavior. Scan only checks for missing updates. Download downloads missing updates. Install installs missing updates. DI does both downloade and install actions. 

The optional "restart" argument is used to issue a restart command to the remote computer if it requires a restart after the action is complete.

The "attached" argument is used to synchronously run the remote agent and capture its standard output. This is much slower than the file based results as each server must be updated in serial, rather than parallel.


The results include the number and enumeration of missing updates, the status of the requested operations, start/finish time, etc.

### Usage
1. Create a group text file, a newline delimited file of server names that will be acted upon.
2. Set up a shared dropbox. This will usually be a network share. The network path must be writable by the SYSTEM account of every computer in the group. Be sure to check both sharing and file permissions.
3. Run `wuf_controller <dropbox_path> <group_file> <action>`
4. For example `wuf_controller \\nas\wufdrop group_1.txt SCAN`
	Or: `wuf_controller \\command_comp\public\wuf_drop group_2.txt DOWNLOAD ATTACHED`

## Agent

### Pre-requisites:
- Cscript.exe (not wscript.exe)
- "Run as Administrator" on Windows Vista and 7.

The agent can be run separate from the controller. In this mode, it simply operates as a command line wrapper around WUA. It will run configured automatic operation, scan, download, or install.

## Props
Lufi99 on technet who appears to have discovered the hack needed to get asynchronous WUA operations to work with VBScript.