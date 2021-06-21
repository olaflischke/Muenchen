Option Explicit

dim computerZahl, benutzerZahl

' Computer "denkt" sich Zahl zwischen 1 und 10 aus
Randomize ' Zufallsgenerator neu initialisieren
computerZahl = Int(rnd()*10)+1

' Benutzer hat 3 Versuche, sie zu erraten
dim versuch
versuch=1
Do
    benutzerZahl = CInt(InputBox("Bitte rate eine Zahl"))
    versuch = versuch + 1
Loop Until versuch > 3 or benutzerZahl = computerZahl

MsgBox("Die richtige Zahl war " & computerZahl)
