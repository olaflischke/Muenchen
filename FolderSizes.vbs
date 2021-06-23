Option Explicit

' Bei Fehler weitermachen - erfordert Fehlerbehandlung, siehe zb. Sub CheckOnError()
On Error Resume Next

dim fso
set fso = CreateObject("Scripting.FileSystemObject")    ' FileSystemObject instanziieren

dim shell
set shell = CreateObject("WScript.Shell")

dim root
set root = fso.GetFolder(shell.ExpandEnvironmentStrings("%WINDIR%")) 'fso.GetFolder("C:\tmp")  ' Stamm-Verzeichnis

dim fld
For Each fld In root.SubFolders    ' Für jeden Unterordner des Stammverzeichnisses
    
    dim size
    size = GetSumOfFilesInFolder(fld)   ' Summe der Dateien im Ordner ermitteln

    'Ausgabe der Folder-Größe
    WScript.Echo fld.Path & ": " & size

Next 

' Fehlerbehandlung abschalten
On Error Goto 0

Function GetSumOfFilesInFolder(folder)
    dim datei, sum
    ' Dateigrößen aufsummieren
    for each datei in folder.Files
        
        if err.Number <> 0 Then
            WScript.Echo datei & ": " & err.Description & " (" & err.Number & ")"
            err.Clear
        end if

        sum = sum + datei.Size
    Next

    if folder.SubFolders.Count > 0 then ' Gibt es Unterordner?
        ' Ordnergrößen der Unterordner aufsummieren
        dim sf
        For Each sf In folder.SubFolders
            sum = sum + GetSumOfFilesInFolder(sf)
            CheckOnError
        Next 
    end If 
    ' Ergebnis zurückgeben
    GetSumOfFilesInFolder = sum
End Function 

sub CheckOnError()
    if err.Number <> 0 Then
        WScript.Echo err.Description & " (" & err.Number & ")"
        err.Clear
    end if
end Sub