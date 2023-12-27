Msgbox "Ez a VBScript a Számok helyesírása ActiveX képességeit mutatja be.", vbinformation

dim iro, szam
set iro = createobject("Szamok.Irasa")
szam=inputbox("Elsõ lépésben nézzük meg egy szám leírását..." & vbcrlf & "Írja be a számot számjegyekkel:","1.) lépés")
msgbox "A " & szam & " leírva: " & iro.szamot_szovegge(cdbl(szam))

szoveg=inputbox("Második lépés nézzük meg egy leírt számról, hogy mi a számképe." & vbcrlf & "A szöveg visszaírásánál, kerüljük a 'két' kifejezést, helyette  használjunk 'kettõ'-t!" & vbcrlf & "Írja le a számot betûkkel:","2.) lépés")
iro.ellenorzes=false
msgbox "A szöveg számképe: " & iro.szoveget_szamma(cstr(szoveg))

msgbox "Bizonyára észrevette, ha kihagyja a kötõjelet, a program akkor is visszaadja a számot. Ennek elkerülésére van egy belsõ ellenõrzõ, ami csak akkor ad vissza értéket, ha jól van kötõjelezve a szöveg. Ha a program 0-t ad vissza, akkor helytelenül lett leírva a szám."
iro.ellenorzes=true
iro.vagolapramasol=true
msgbox "A szöveg kötõjelesen ellenõrzött számképe: " & iro.szoveget_szamma(cstr(szoveg))
msgbox "Mellesleg az utolsó eredmény a vágólapra lett másolva..."
