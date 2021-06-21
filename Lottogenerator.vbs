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
dim zahlen(5)

For i=0 To anzahlReihen-1

    For j=0 To UBound(zahlen)
        dim zahl
        Do
            zahl = int(rnd() * 49) + 1
        Loop While IsInArray(zahl, zahlen)
        zahlen(i) = zahl
        ausgabe = ausgabe & zahlen(i) & ", "
    Next

    ausgabe = ausgabe & vbCrLf

Next

MsgBox(ausgabe)

Function IsInArray(element, array)
    dim k
    For k=0 To ubound(array)
        If array(k)=element Then
            IsInArray = True
            Exit Function
        End If
    Next

    IsInArray = False
End Function 