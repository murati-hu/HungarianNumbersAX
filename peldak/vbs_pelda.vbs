Msgbox "Ez a VBScript a Sz�mok helyes�r�sa ActiveX k�pess�geit mutatja be.", vbinformation

dim iro, szam
set iro = createobject("Szamok.Irasa")
szam=inputbox("Els� l�p�sben n�zz�k meg egy sz�m le�r�s�t..." & vbcrlf & "�rja be a sz�mot sz�mjegyekkel:","1.) l�p�s")
msgbox "A " & szam & " le�rva: " & iro.szamot_szovegge(cdbl(szam))

szoveg=inputbox("M�sodik l�p�s n�zz�k meg egy le�rt sz�mr�l, hogy mi a sz�mk�pe." & vbcrlf & "A sz�veg vissza�r�s�n�l, ker�lj�k a 'k�t' kifejez�st, helyette  haszn�ljunk 'kett�'-t!" & vbcrlf & "�rja le a sz�mot bet�kkel:","2.) l�p�s")
iro.ellenorzes=false
msgbox "A sz�veg sz�mk�pe: " & iro.szoveget_szamma(cstr(szoveg))

msgbox "Bizony�ra �szrevette, ha kihagyja a k�t�jelet, a program akkor is visszaadja a sz�mot. Ennek elker�l�s�re van egy bels� ellen�rz�, ami csak akkor ad vissza �rt�ket, ha j�l van k�t�jelezve a sz�veg. Ha a program 0-t ad vissza, akkor helytelen�l lett le�rva a sz�m."
iro.ellenorzes=true
iro.vagolapramasol=true
msgbox "A sz�veg k�t�jelesen ellen�rz�tt sz�mk�pe: " & iro.szoveget_szamma(cstr(szoveg))
msgbox "Mellesleg az utols� eredm�ny a v�g�lapra lett m�solva..."
