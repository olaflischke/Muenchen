Option Explicit
dim ziel, betrag
' Benutzereingaben
ziel = InputBox("Welche Währung?")
betrag = InputBox ("Betrag (Fremdwährung)?")

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
    dim waehrungen, raten, i
    dim node
    For Each node In cubeNodes
        dim rate, iso
        rate = node.GetAttribute("rate")
        iso = node.GetAttribute("currency")

        if iso <> "" AND rate <> "" then 
            'WScript.Echo iso & ": " & rate        
            ' waehrungen(i) = iso
            ' raten(i) = rate
            if iso = ziel then
                dim ergebnis
                ergebnis = betrag / rate
                MsgBox(betrag & " " & ziel & " sind " & FormatNumber(ergebnis, 2) & " EUR")
                exit For
            End if
        end If 
    Next
End If
