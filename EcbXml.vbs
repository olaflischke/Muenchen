Option Explicit

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
    dim node
    For Each node In cubeNodes
        dim rate, iso
        rate = node.GetAttribute("rate")
        iso = node.GetAttribute("currency")

        if iso <> "" AND rate <> "" then 
            WScript.Echo iso & ": " & rate        
        end If 
    Next
End If
