Option Explicit
dim WSHShell
set WSHShell = CreateObject("WScript.Shell")

on error resume next
dim benutzer, frage, ueberschrift, meldung

benutzer = WSHShell.RegRead("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\RegisteredOwner")

frage = "Unter welchem Namen soll Windows registriert werden?"
ueberschrift = "Benutzer"

benutzer = InputBox(frage, ueberschrift, benutzer)

WSHShell.RegWrite "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\RegisteredOwner", benutzer

meldung = "Registrierter Benutzer ist " & "jetzt """ & benutzer & """!"

MsgBox meldung, vbInformation