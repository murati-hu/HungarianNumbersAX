@ECHO OFF
ECHO Szamok helyesirasa ActiveX szolgaltatasainak
ECHo regisztralasa....
ECHO.
copy szamok.dll c:\windows\system32\szamok.dll
regsvr32.exe /i /s c:\windows\system32\szamok.dll
ECHO Kesz.
ECHO.
ECHO Osztalynev: Szamok.Irasa
ECHO Fuggvenyei:
ECHO 		szamot_szovegge(Szam as double) as string
ECHo 		szoveget_szamma(Szoveg as string) as double
ECHO Valtozoi:
ECHO		Ellenorzes as boolean
ECHO		VagolapraMasol as boolean
ECHO.
pause
