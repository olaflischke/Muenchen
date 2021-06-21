' Benutzer soll festlegen können, wieviele Lottoreihen er spielen möchte (max. 12)
' Lottoreihe hat 6 Zahlen zwischen 1 und 49 (nicht wiederholend!)
' Gezogene Lottoreihen sollen komma-getrennt ausgegeben werden


Option Explicit

dim anzahlReihen
anzahlReihen = CInt( InputBox("Wieviele Reihen sollen es sein?") )

if anzahlReihen > 12 Then
    MsgBox("Maximal 12 Stück!")
    Stop
End If

dim i, j, ausgabe
dim zahlen '(5)

For i=0 To anzahlReihen-1
    zahlen = ErstelleLottoreihe ' Erstellt ein Array mit 6 Zahlen zwischen 1 und 49 (nicht wiederholend)
    ausgabe = ausgabe & vbCrLf & KommaString(zahlen) ' Erzeugt aus den Elementen des gg. Arrays eine komma-getrennte Zeichenkette
Next

MsgBox(ausgabe)

Function IsInArray(byval element, byval array)
    dim k
    For k=0 To ubound(array)
        If array(k)=element Then
            IsInArray = True
            Exit Function
        End If
    Next

    IsInArray = False
End Function 

' Erstellt ein Array mit 6 Zahlen zwischen 1 und 49 (nicht wiederholend)
Function ErstelleLottoreihe()

    dim array(5), l, zahl
    For l=0 To Ubound(array)
        Do
            zahl = Int(rnd()*49)+1
        Loop While IsInArray(zahl, array)

        array(l)=zahl
    Next

    ErstelleLottoreihe = array
End Function

' Erzeugt aus den Elementen des gg. Arrays eine komma-getrennte Zeichenkette
Function KommaString(byval array) 
    dim m, ergebnis

    For m=0 To ubound(array)
        ergebnis = ergebnis & array(m) & ", "
    Next

    KommaString = ergebnis
End Function