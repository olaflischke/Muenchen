Option Explicit

dim fso
Set fso = CreateObject("Scripting.FileSystemObject")    ' FileSystemObject instanziieren

dim drives
set drives = fso.Drives ' Collection von Drive-Objekten

dim drive, s, n
For Each drive In drives
    s = s & drive.DriveLetter & ": - " & drive.DriveType  & " - "
    If drive.DriveType = 3 Then ' Netzlaufwerk
        n = drive.ShareName
    Else
        n = drive.VolumeName
    End If
    s = s & n & vbCrLf
Next 

WScript.Echo s

dim fld
set fld = fso.GetFolder("C:\")

WScript.Echo fld.Path