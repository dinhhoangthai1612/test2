'Need login first times then remember ID

Option Explicit

Dim Shell, WMI, wql, process

Set Shell = CreateObject("WScript.Shell")
Set WMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

wql = "SELECT ProcessId FROM Win32_Process WHERE Name = 'UiPath.Executor.exe'"

For Each process In WMI.ExecQuery(wql)
    Shell.AppActivate process.ProcessId	
	Shell.SendKeys "{TAB}"	
	Shell.SendKeys "{ENTER}"	
	WScript.Sleep 10000
	
    
	'Shell.SendKeys "{TAB}"
	'WScript.Sleep 500
	'Shell.SendKeys "+{TAB}"
	'Shell.SendKeys "roboticsnt@ykk.com"
	Shell.SendKeys "{TAB}"
	Shell.SendKeys "+{TAB}"
	Shell.SendKeys "Password@123"
	Shell.SendKeys "{ENTER}"	
	WScript.Sleep 15000
Next
