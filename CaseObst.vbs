dim obst

obst=InputBox("Was hast Du in der Hand?")

Select Case obst
    Case "Apfel", "Birne"
        WScript.Echo "Du hast ein Obst"
    case "Kartoffel"
        WScript.Echo "Kein Obst"
    Case Else
        WScript.Echo "Keine Ahnung, was das ist"
End Select

dim zahl
zahl = CInt( InputBox("Eine Zahl:"))

Select Case zahl
    Case zahl > 10
        WScript.Echo "Größer als 10"
    Case Else
        WScript.Echo "Keine Ahnung"
End Select

