Option Explicit

dim shell 
set shell = CreateObject("Wscript.Shell")

dim element
For Each element In shell.SpecialFolders
    WScript.Echo element
Next 

dim variable
For Each variable In shell.Environment
    WScript.Echo variable
Next ' variable

