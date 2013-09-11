win-maintenance
===============

Use this script to do a variety of different operational tasks behind the
scenes based on which version of Windows is detected to be in use.  Place this
script into the \\domain\netlogon folder to be run as a user profile logon
script.  '

The following is a list of things that this script is intended to handle:

* Updates DNS settings
* Syncs up local time to be current with its domain controller.
* Calls "deleteTempFiles.vbs" to clean up temp files upon login.
* Maps network drives.
* Create and Update a number of Windows scheduled tasks.
* Create and encrypt the "Encrypt" folder in My Documents.
