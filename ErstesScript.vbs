Option Explicit ' Variablendeklaration erforderlich

Dim name, zahl1, zahl2
name = InputBox("Gib Deinen Namen ein")
zahl1 = CInt( InputBox("Zahl 1:"))
zahl2 = CInt( InputBox("Zahl 2:"))
MsgBox(zahl1 + zahl2)
