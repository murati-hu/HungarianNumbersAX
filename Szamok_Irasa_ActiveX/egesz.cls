VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 2  'vbComplexBound
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "egesz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private egyes(0 To 9) As String, tizes(0 To 9) As String, kivetel(1 To 2) As String
Private ertek(1 To 20) As Byte, neve(1 To 20) As String
Private szorzo(1 To 4) As Long, szneve(1 To 4) As String
Public ellenor As Boolean

Private Sub Class_Initialize()
    Dim l
    ellenor = False
'Átalakítási értékek
    kivetel(1) = "tíz"
    kivetel(2) = "húsz"

    egyes(0) = ""
    egyes(1) = "egy"
    egyes(2) = "kettõ"
    egyes(3) = "három"
    egyes(4) = "négy"
    egyes(5) = "öt"
    egyes(6) = "hat"
    egyes(7) = "hét"
    egyes(8) = "nyolc"
    egyes(9) = "kilenc"
    
    tizes(0) = ""
    tizes(1) = "tizen"
    tizes(2) = "huszon"
    tizes(3) = "harminc"
    tizes(4) = "negyven"
    tizes(5) = "ötven"
    tizes(6) = "hatvan"
    tizes(7) = "hetven"
    tizes(8) = "nyolcvan"
    tizes(9) = "kilencven"
    
'Visszakeresési értékek
    
    For l = 1 To 9
        neve(l) = tizes(l)
        ertek(l) = l * 10
    Next l
    
    neve(10) = "tíz"
    ertek(10) = 10
    
    neve(11) = "húsz"
    ertek(11) = 20
    
    For l = 1 To 9
        neve(l + 11) = egyes(l)
        ertek(l + 11) = l
    Next l
    
    
    
    
'Szorzók
    szneve(1) = "száz"
    szorzo(1) = 100
    szneve(2) = "ezer"
    szorzo(2) = 1000
    szneve(3) = "millió"
    szorzo(3) = 1000000
    szneve(4) = "billió"
    szorzo(4) = 1000000000
    
End Sub

Public Function szamot_szovegge(szam As Long) As String
On Error GoTo hiba
    Dim i As Integer, szam_str As String, szoveg As String, kj As String
    Dim jegyek(1 To 10) As Byte, atalakitott(1 To 10) As String, minus As String
    
    szoveg = ""
    If szam < 0 Then minus = "mínusz " Else minus = ""
    szam_str = CStr(Abs(szam))

    'a számjegyek sorrendjének felcserélése
    For i = 1 To Len(szam_str)
        jegyek(i) = Mid(szam_str, Len(szam_str) - i + 1, 1)
    Next i

    'helyiértékes vizsgálat és átalakítás
    For i = 1 To Len(szam_str)
    
        If jegyek(i) = 0 And Len(szam_str) > 1 Then
            atalakitott(i) = ""
            GoTo kov
        End If
    
        Select Case i
            Case 1
                If jegyek(i) = 0 And Len(szam_str) = 1 Then
                        atalakitott(i) = "nulla"
                    Else
                        atalakitott(i) = egyes(jegyek(i))
                End If
            
            Case 2, 5, 8
                If atalakitott(i - 1) = "" And (jegyek(i) = 1 Or jegyek(i) = 2) Then
                        atalakitott(i) = kivetel(jegyek(i))
                    Else
                        atalakitott(i) = tizes(jegyek(i))
                End If
            
                Select Case i
                    Case 5
                        If atalakitott(i - 1) = "" Then atalakitott(i) = atalakitott(i) & "ezer-"
                    Case 8
                        If atalakitott(i - 1) = "" Then atalakitott(i) = atalakitott(i) & "millió-"
                End Select
            
            Case 3, 6, 9
                If jegyek(i) = 1 Then
                        atalakitott(i) = "száz"
                    Else
                        atalakitott(i) = egyes(jegyek(i)) & "száz"
                End If
            
                Select Case i
                    Case 6
                        If atalakitott(i - 1) = "" And atalakitott(i - 2) = "" Then atalakitott(i) = atalakitott(i) & "ezer-"
                    Case 9
                        If atalakitott(i - 1) = "" And atalakitott(i - 2) = "" Then atalakitott(i) = atalakitott(i) & "millió-"
                End Select
            
            Case 4
                If CDbl(szam_str) > 2000 Then kj = "-" Else kj = ""
                If jegyek(i) = 1 And Len(szam_str) = 4 Then
                        atalakitott(i) = "ezer"
                    Else
                        atalakitott(i) = egyes(jegyek(i)) & "ezer" & kj
                End If
            Case 7
                atalakitott(i) = egyes(jegyek(i)) & "millió" & "-"
            Case 10
                atalakitott(i) = egyes(jegyek(i)) & "billió"
        End Select
kov:
    Next i
    
    'szövegrészletek összeillesztése
    For i = 1 To 10
        szoveg = atalakitott(i) + szoveg
    Next i

    'a felesleges kötõjel levágása
    If Mid(szoveg, Len(szoveg), 1) = "-" Then
        szoveg = Mid(szoveg, 1, Len(szoveg) - 1)
    End If
    
    szamot_szovegge = minus & szoveg
    
    Exit Function
hiba:
    szamot_szovegge = "Túlcsorduló eredmény"
End Function

Public Function szoveget_szamma(szoveg As String) As Long
On Error GoTo hiba
    Dim k As Integer, i As Integer, osszeg As Long, resz As Long, tal As Boolean
    Dim minus As Integer
    
    szoveg = LCase(szoveg)
    If Mid(szoveg, 1, 6) = "mínusz" Then
            minus = -1
            szoveg = Mid(szoveg, 7, Len(szoveg) - 6)
        Else
            minus = 1
    End If
    tal = False
    osszeg = 0
    k = 0
be:
    If tal = False Then
            k = k + 1
        Else
            tal = False
    End If
    
        For i = 1 To 20
            If neve(i) = Mid(szoveg, k, Len(neve(i))) Then
                resz = resz + ertek(i)
                k = k + Len(neve(i))
                tal = True
                GoTo szorzo
            End If
        Next i
szorzo:
        For i = 1 To 4
            If szneve(i) = Mid(szoveg, k, Len(szneve(i))) Then
                If resz = 0 Then resz = 1
                resz = resz * szorzo(i)
                If i >= 2 Then
                    osszeg = osszeg + resz
                    resz = 0
                End If
                k = k + Len(szneve(i))
                tal = True
                GoTo ell
            End If
        Next i
ell:
    If k <= Len(szoveg) Then GoTo be
    osszeg = osszeg + resz
    If ellenor Then
        If szoveg = szamot_szovegge(osszeg) Then
            szoveget_szamma = minus * osszeg
        Else
            szoveget_szamma = 0
        End If
    Else
        szoveget_szamma = minus * osszeg
    End If
    
    Exit Function
hiba:
    szoveget_szamma = 0
End Function
Public Function nevjegy() As String
    nevjegy = "Magyar számok írása v" & App.Major & "." & App.Minor & "." & App.Revision & ". Freeware ActiveX - Muráti Ákos 2002."
End Function
