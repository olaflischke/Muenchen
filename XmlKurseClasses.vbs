Option Explicit
dim ziel, betrag
' Benutzereingaben
ziel = InputBox("Welche Währung?")
betrag = InputBox ("Betrag (Fremdwährung)?")

dim waehrungen
set waehrungen = CreateObject("System.Collection.ArrayList") ' ActiveX-Fehler ?
LiesXml

dim ergebnis
ergebnis = betrag / waehrungen.Item(ziel)

MsgBox betrag & " " & ziel & " sind " & FormatNumber(ergebnis, 2) & " EUR"

Sub LiesXml()
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
                'WScript.Echo iso & ": " & rate        
                'waehrungen.Add iso, rate
                dim waehrung
                set waehrung = new Sorte
                waehrung.Zeichen=iso
                waehrung.Eurokurs=rate
                
                'WScript.Echo waehrung.Zeichen & ": " & waehrung.Eurokurs

               waehrungen.Add waehrung
            end If 
        Next
    End If

end sub

Class Sorte
    ' Backing Field für die Eurokurs-Property
    Private m_Eurokurs

    Public Property Get Eurokurs
        Eurokurs = m_Eurokurs
    End Property

    Public Property Let Eurokurs(Value)
        m_Eurokurs = Value
    End Property

    Private m_Zeichen

    Public Property Get Zeichen
         Zeichen = m_Zeichen
    End Property

    Public Property Let Zeichen(Value)
         m_Zeichen = Value
    End Property

    public Function Machwas()
        ' Beispiel für eine Methode
    end Function

    Private Sub Class_Initialize()
        
    End Sub

    Private Sub Class_Terminate()
        
    End Sub
End Class ' Sorte