Option Explicit
dim wmiServices, computerName

computerName="."

set wmiServices = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & computerName & "\root\cimv2")

dim runningServices
set runningServices = wmiServices.execQuery("Select * from Win32_Service where state like 'running'")

dim service
For Each service In runningServices
   WScript.Echo  service.DisplayName & vbTab & service.state
Next 