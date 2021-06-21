' \b[\w\.-]+@[\w\.-]+\.\w{2,4}\b - Regulärer Ausdruck für eine Emailadresse

Option Explicit
dim mailAdresse

mailAdresse=InputBox("Gib Deine Emailadresse an:")

dim reg
set reg = new RegExp    ' Instanz der RegExp-Klasse
reg.Pattern = "\b[\w\.-]+@[\w\.-]+\.\w{2,4}\b"
reg.IgnoreCase = False

dim antwort
If reg.Test(mailAdresse) = True Then
    MsgBox mailAdresse & " ist eine formal gueltige Emailadresse", vbOKOnly & vbInformation, "Check ok"
Else
    antwort= MsgBox( mailAdresse & " ist so nicht ok.", vbYesNo , "Check fehlgeschlagen")
    If antwort=vbNo Then
        MsgBox("Gar nicht wahr, ist und bleibt ungültig!")
    End If
End If


If True Then
    ' Wahr
Else
    ' Falsch
End If