dim wert1, wert2
dim ergebnis, ergebnisAusFunction

wert1=1
wert2=2

ergebnisAusFunction=ByFunction(wert1, wert2)
ergebnis = wert1 + wert2

WScript.Echo ergebnisAusFunction
WScript.Echo ergebnis

Function ByFunction(zahl1, zahl2) ' ByRef ist Default!

    zahl1=zahl1+10
    zahl2=zahl2+25

    ByFunction=zahl1+zahl2
        
End Function 

Function ByValFunction(ByVal zahl1, ByVal zahl2) ' Wert(e) aus den übergebenen Variablen werden benutzt

    zahl1=zahl1+10
    zahl2=zahl2+25

    ByValFunction=zahl1+zahl2
        
End Function 

Function ByRefFunction(ByRef zahl1, ByRef zahl2) ' Ursprüngliche Variablen werden mit verändert, weil Speicheradresse übergeben wird

    zahl1=zahl1+10
    zahl2=zahl2+25

    ByRefFunction=zahl1+zahl2
        
End Function 