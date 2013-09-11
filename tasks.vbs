'======================================================================================
'
'Name		  tasks.vbs
'Author		Josh Reichardt
'Email		josh.reichardt@gmail.com
'Date		  12/12/12
'
'Comments	Use this script to do a variety of different operational tasks behind
'			    the scenes based on which version of Windows is detected to be in use.  
'			    Place this script into the \\domain\netlogon folder to be run
'			    as a user profile logon script.
'
'		  The following is a list of things that this script is intended to handle:
'
'			-Updates DNS settings
'			-Syncs up local time to be current with its domain controller.
'			-Calls "deleteTempFiles.vbs" to clean up temp files upon login.
'			-Maps network drives.
'			-Create and Update a number of Windows scheduled tasks.
'			-Create and encrypt the "Encrypt" folder in My Documents.
'
'======================================================================================
'
'System level variables.
Set WshShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set networkdrive = CreateObject("Scripting.FileSystemObject")
Set tDrive = CreateObject("WScript.Network")
Set nDrive = CreateObject("WScript.Network")
Set yDrive = CreateObject("WScript.Network")
'Profile settings
UserProfile = WshShell.ExpandEnvironmentStrings("%USERPROFILE%")
UserProfileXP = WshShell.ExpandEnvironmentStrings(UserProfile & "\My Documents\Encrypt")
UserProfile7 = WshShell.ExpandEnvironmentStrings(UserProfile & "\Documents\Encrypt")
'Get computer properties.
strComputer = "."
'Get the WMI Operating System object and query results
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set oss = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")

'Refresh DNS Settings.
WshShell.Run "ipconfig /registerdns", 0, true
WshShell.Run "ipconfig /flushdns", 0, true

'Sync up local time to DC.
WshShell.Run "NET TIME %logonserver% /SET /Y", 0 , true

'Check if the machine is a server, quit if it is.
If (objFSO.FileExists("c:\server.dat")) = TRUE Then WScript.Quit

'Check if machine is a VDI, quit if it is.
If (objFSO.FileExists("c:\VDI.dat")) = TRUE Then WScript.Quit

'Get the OS version number.
For Each os in oss
	OSVersion = Left(os.Version,3)
Next

'Convert OS versions into names.
Select Case OSVersion
	Case "6.2"
		OSName = "Windows 8"
	Case "6.1"
		OSName = "Windows 7"
	Case "6.0" 
		OSName = "Windows Vista"
	Case "5.2" 
		OSName = "Windows 2003"
	Case "5.1" 
		OSName = "Windows XP"
	Case "5.0" 
		OSName = "Windows 2000"
	Case Else
		OSName = "Version not supported"
End Select

'Output the OS name.
'WScript.Echo OSName

If OSName = "Windows 2003" Then
	Wscript.Quit
end if

'XP specific tasks.
If OSName = "Windows XP" Then
	dim xp
	set xp = createobject("wscript.shell")
	
	'Create tasks.  Need to double check that this works.
	if not objFSO.FileExists("c:\windows\tasks\Defrag.job") then
		'objFSO.DeleteFile("c:\windows\tasks\Defrag.job")
		xp.run "schtasks /create /TN Defrag /SC WEEKLY /MO 1 /D SUN /ST 18:00:00 /RU SYSTEM /TR ""c:\windows\system32\defrag.exe C:""", 0, true
	end if
	
	if not objFSO.FileExists("c:\windows\tasks\Reboot.job") then
		'objFSO.DeleteFile("c:\windows\tasks\Reboot.job")
		xp.run "schtasks /create /TN Reboot /SC WEEKLY /MO 1 /D SUN /ST 01:00:00 /RU SYSTEM /TR ""c:\windows\system32\shutdown.exe -r -t 01 -f""", 0, true
	end if
	
	if not objFSO.FileExists("c:\windows\tasks\GPUpdate.job") then
		'objFSO.DeleteFile("c:\windows\tasks\GPUpdate.job")
		xp.run "schtasks /create /TN GPUpdate /SC DAILY /ST 00:00:00 /RU SYSTEM /TR ""c:\windows\system32\gpupdate.exe /force""", 0, true
	end if
	
	'Check if Encrypt folder exists.
	if not objFSO.FolderExists(UserProfileXP) then
		objFSO.CreateFolder(UserProfileXP)
	end if
	
	'Encrypt folder and files.
	WshShell.Run "cipher /e /a /s:" & chr(34) & UserProfileXP & chr(34), 0, true

	set xp = nothing

'Win 7 specific tasks.
ElseIf OSName = "Windows 7" Or "Windows 8" Then
	
	If WScript.Arguments.length = 0 Then 
		Set ObjShell = CreateObject("Shell.Application") 
		ObjShell.ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """" & " RunAsAdministrator", , "runas", 1 
	Else	
		dim seven
		set seven = createobject("wscript.shell")
		
		'Check if Encrypt folder exists, create if it doesn't.
		if not objFSO.FolderExists(UserProfile7) then
			objFSO.CreateFolder(UserProfile7)
		end if
		
		'Encrypt folder and files.
		WshShell.Run "cipher /e /a /s:" & chr(34) & UserProfileXP & chr(34), 0, true
		
		'Create tasks.
		if not objFSO.FileExists("C:\Windows\System32\Tasks\GPUpdate") then
			seven.run "schtasks /create /TN GPUpdate /SC DAILY /ST 00:00:00 /RU SYSTEM /RL HIGHEST /TR ""c:\windows\system32\gpupdate.exe /force"" /F", 0, true
		end if
		
		if not objFSO.FileExists("C:\Windows\System32\Tasks\Defrag") then
			seven.run "schtasks /create /TN Defrag /SC WEEKLY /MO 1 /D SUN /ST 18:00:00 /RU SYSTEM /RL HIGHEST /TR ""c:\windows\system32\defrag.exe C:"" /F", 0, true
		end if
		
		if not objFSO.FileExists("C:\Windows\System32\Tasks\Reboot") then
			seven.run "schtasks /create /TN Reboot /SC WEEKLY /MO 1 /D SUN /ST 01:00:00 /RU SYSTEM /RL HIGHEST /TR ""c:\windows\system32\shutdown.exe -r -t 01 -f"" /F", 0, true
		end if
		
		set seven = nothing

	End If
	
Else
	WScript.Echo "Couldn't add scheduled task.  Please contact your administrator."
End If

'Clean up a bit.
set objFSO = nothing
set WshSHell = nothing
set networkdrive = nothing
set tDrive = nothing
set nDrive = nothing
set yDrive = nothing
