strComputer = "."
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

For Each objclass in objWMIService.SubclassesOf()
    Wscript.Echo objClass.Path_.Class
Next
