Option Explicit

dim array(5)    ' Array mit 6 Elementen
dim i, j        ' Schleifenvariablen

For i=0 To UBound(array) ' UBound liefert die obere Grenze (Upper Boundary) des Arrays (Anzahl der Elemente)
    array(i) = i
Next 

dim meldung

For j=0 To UBound(array)
    meldung = meldung & array(j) & ", "
Next 

MsgBox(meldung)

WScript.Echo TypeName(array) & " array: " & UBound(array) &" Elemente."
WScript.Echo TypeName(meldung) & " meldung: " & meldung
WScript.Echo TypeName(i) & " i: " & i
WScript.Echo TypeName(j) & " j: " & j
