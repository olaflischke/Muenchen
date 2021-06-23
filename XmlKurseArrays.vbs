Option Explicit
dim ziel, betrag
' Benutzereingaben
ziel = InputBox("Welche Währung?")
betrag = CDbl( InputBox ("Betrag (Fremdwährung)?"))

dim waehrungen(), raten()
LiesXml 

dim index
index = HolePosition(ziel, waehrungen)

dim ergebnis
ergebnis = betrag / raten(index)

MsgBox betrag & " " & ziel & " sind " & FormatNumber(ergebnis, 2) & " EUR"

Function HolePosition(element, array)
    dim i
    For i=0 To ubound(array)
        If array(i)=element Then
            HolePosition=i
            Exit Function
        End If
    Next

    HolePosition = -1
End Function ' HolePosition

sub LiesXml()
    dim xmlFile
    xmlFile = "https://www.ecb.europa.eu/stats/eurofxref/eurofxref-daily.xml"

    dim xmlDoc
    set xmlDoc = CreateObject("Microsoft.XMLDOM") ' DocumentObjectModel - XML-Dokument im Arbeitsspeicher
    xmlDoc.Async = False
    xmlDoc.Load(xmlFile)

    dim root
    set root = xmlDoc.DocumentElement   ' Rootknoten des Dokuments

    dim cubeNodes
    set cubeNodes = root.GetElementsByTagName("Cube")   ' Alle Tags, die "Cube" heißen

    If cubeNodes.Length > 0 Then
        dim i
        i=0

        redim waehrungen(0)
        Redim raten(0)

        dim node
        For Each node In cubeNodes
            dim rate, iso
            rate = node.GetAttribute("rate")
            iso = node.GetAttribute("currency")

            WScript.Echo "i: " & i

            if iso <> "" AND rate <> "" then 
                'WScript.Echo iso & ": " & rate        
                WScript.Echo "waehrungen: " & UBound(waehrungen)


                If ubound(waehrungen) < i Then
                    Redim Preserve waehrungen(i)
                    Redim Preserve raten(i)
                End If

                waehrungen(i) = iso
                raten(i) = rate
                i=i+1
            end If 
        Next
    End If
end sub