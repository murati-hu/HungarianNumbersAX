VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 2  'vbComplexBound
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Irasa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private nevezek(1 To 11) As String
Dim egesz As Object
Private vagolap As Boolean

Public Function szamot_szovegge(tort As Double) As String
On Error GoTo hiba
    Dim egesz_resz As String, tort_resz As String
    Dim tort_str As String, i As Integer
    
    tort_str = CStr(tort)
    
    For i = 1 To Len(tort_str)
        If Mid(tort_str, i, 1) = "." Or Mid(tort_str, i, 1) = "," Then
            egesz_resz = egesz.szamot_szovegge(CLng(Mid(tort_str, 1, i - 1)))
            tort_resz = egesz.szamot_szovegge(CLng(Mid(tort_str, i + 1, Len(tort_str) - i)))
            szamot_szovegge = egesz_resz & " egész " & tort_resz & " " & nevezek(Len(Mid(tort_str, i + 1, Len(tort_str) - i)))
            If CLng(Mid(tort_str, 1, i - 1)) = 0 And tort < 0 Then
                szamot_szovegge = "mínusz " & szamot_szovegge
            End If
            Vagolapra (szamot_szovegge)
            Exit Function
        End If
    Next i
    szamot_szovegge = egesz.szamot_szovegge(CLng(tort_str))
    Vagolapra (szamot_szovegge)
    Exit Function
hiba:
    szamot_szovegge = "Túl nagy számot adott meg."
    Vagolapra ("")
End Function

Private Sub Class_Initialize()
    Set egesz = New egesz
    nevezek(1) = "tized"
    nevezek(2) = "század"
    nevezek(3) = "ezred"
    nevezek(4) = "tízezred"
    nevezek(5) = "százezred"
    nevezek(6) = "milliomod"
    nevezek(7) = "tízmilliomod"
    nevezek(8) = "százmilliomod"
    nevezek(9) = "billiomod"
    nevezek(10) = "tízbilliomod"
    nevezek(11) = "százbilliomod"
    
    vagolap = False
End Sub
Public Function nevjegy() As String
    nevjegy = egesz.nevjegy
End Function

Public Function szoveget_szamma(szoveg As String) As Double
On Error GoTo hiba
    Dim egesz_resz As String, tort_resz As String, helyiertek As String
    Dim i As Integer, j As Integer, minus As Integer
    szoveg = LCase(szoveg)
    
    For i = 1 To Len(szoveg)
        If Mid(szoveg, i, 7) = " egész " Then
            egesz_resz = egesz.szoveget_szamma(Mid(szoveg, 1, i - 1))
            If Mid(szoveg, 1, 6) = "mínusz" And egesz_resz = 0 Then
                    minus = -1
                    'szoveg = Mid(szoveg, 7, Len(szoveg) - 6)
                Else
                    minus = 1
            End If
            tort_resz = Mid(szoveg, i + 7, Len(szoveg) - i - 6)
                For j = 1 To Len(tort_resz)
                    If Mid(tort_resz, j, 1) = " " Then
                        helyiertek = Mid(tort_resz, j + 1, Len(tort_resz) - j)
                        tort_resz = egesz.szoveget_szamma(Mid(tort_resz, 1, j - 1))
                        GoTo ki
                    End If
                Next j
ki:
            For j = 1 To 11
                If helyiertek = nevezek(j) Then
                    szorzo = 10 ^ j
                End If
            Next j
            
            If egesz_resz < 0 Then
                tort_resz = tort_resz * -1
            End If
            
            If egesz.ellenor Then
                If szoveg = szamot_szovegge(CLng(egesz_resz) + (CLng(tort_resz) / szorzo)) Then
                        szoveget_szamma = minus * (CLng(egesz_resz) + (CLng(tort_resz) / szorzo))
                    Else
                        szoveget_szamma = 0
                End If
            Else
                szoveget_szamma = minus * (CLng(egesz_resz) + (CLng(tort_resz) / szorzo))
            End If
            Vagolapra (CStr(szoveget_szamma))
            Exit Function
        End If
    Next i
    szoveget_szamma = egesz.szoveget_szamma(szoveg)
    Vagolapra (CStr(szoveget_szamma))
    Exit Function
hiba:
    szoveget_szamma = 0
    Vagolapra (CStr(szoveget_szamma))
End Function
Public Property Get Ellenorzes() As Boolean
    Ellenorzes = egesz.ellenor
End Property
Public Property Let Ellenorzes(uj As Boolean)
    egesz.ellenor = uj
End Property
Private Sub Vagolapra(mit As String)
    If vagolap Then Clipboard.SetText mit
End Sub

Private Sub Class_Terminate()
    Debug.Print "vége"
End Sub
Public Property Get VagolapraMasol() As Boolean
    VagolapraMasol = vagolap
End Property
Public Property Let VagolapraMasol(uj As Boolean)
    vagolap = uj
End Property
