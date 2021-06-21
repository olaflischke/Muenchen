dim zahl

zahl = HoleWert(4, 2) ' Rechts vom = Klammern!
HoleWert 8, 8 ' Ohne = Keine Klammern! (Außer 1 Parameter)
MachEtwas 2, 4  ' Ohne = Keine Klammern!
Call MachEtwas(2, 4) ' Aufruf mit Call: Klammern!


WScript.Echo zahl

' Sub: Subroutine/Unterroutine - gibt keinen Wert zurück
Sub MachEtwas(faktor1, faktor2)
    zahl = zahl * (faktor1 + faktor2)
End Sub

' Function: Gibt einen Wert zurück
Function HoleWert(wert1, wert2)
    HoleWert = wert1 * wert2 ' Gibt das zurück, was dem Namen der Funktion zugewiesen wird
End Function 