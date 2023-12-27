@ECHO OFF
ECHO Szamok helyesirasa ActiveX szolgaltatasainak
ECHo regisztralasa....
ECHO.
vbrun.exe /q
copy szamok.dll c:\windows\system\szamok.dll
c:\windows\system\regsvr32 /i /s c:\windows\system\szamok.dll
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
