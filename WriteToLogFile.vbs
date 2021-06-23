Option Explicit

WriteToLog "Log gestartet.", "C:\tmp\vbScript.log"

function WriteToLog(meldung, logFile)
    ' Meldung in ein Logfile schreiben
    ' Wenn Log nicht existiert, anlegen, ansonsten anhängen ans bestehende Log
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    meldung = Now & " - " & meldung

    dim log
    If fso.FileExists(logFile) Then
        ' Datei existiert bereits? Reinschreiben!
        set log = fso.OpenTextFile(logFile, 8)
    Else
        ' Datei existiert noch nicht? Erstellen!
        set log = fso.CreateTextFile(logFile, True)
    End If

    log.WriteLine(meldung)
    log.Close

end function