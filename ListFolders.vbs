Option Explicit

dim fso
set fso = CreateObject("Scripting.FileSystemObject")    ' FileSystemObject instanziieren

dim drive
set drive = fso.Drives("C") ' Collection der Drives, daraus Element namens "C"

dim i 
i = fso.Drives.count

dim fld
set fld = drive.RootFolder  'fso.GetFolder("C:\")  ' Stamm-Verzeichnis

dim f
For Each f In fld.SubFolders    ' Für jeden Unterordner des Stammverzeichnisses
    dim d
    for each d in f.Files
        d.si
    Next
Next 


