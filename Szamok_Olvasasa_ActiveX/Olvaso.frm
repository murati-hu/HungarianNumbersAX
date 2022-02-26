VERSION 5.00
Begin VB.Form Olvaso 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Olvasás:"
   ClientHeight    =   255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   1920
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   238
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   255
   ScaleWidth      =   1920
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer szunet 
      Enabled         =   0   'False
      Interval        =   750
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "Olvaso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'DirectX és DS életrehívása
Dim Dx As New DirectX7
Dim Ds As DirectSound

Dim DsDesc As DSBUFFERDESC
Dim DsWave As WAVEFORMATEX

Dim hangok(0 To 50) As DirectSoundBuffer 'A 51 jelentése "semmi"
Dim jatszando(0 To 100) As Byte
Dim mutato As Byte

Private Sub Form_Load()
On Error GoTo hiba
    Olvaso.Show
    Olvaso.SetFocus
    Olvaso.ZOrder 0
    
    Set Ds = Dx.DirectSoundCreate("")

    Ds.SetCooperativeLevel Olvaso.hWnd, DSSCL_NORMAL

    DsDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
    DsWave.nFormatTag = WAVE_FORMAT_PCM
    DsWave.nChannels = 2
    DsWave.lSamplesPerSec = 22050
    DsWave.nBitsPerSample = 16
    DsWave.nBlockAlign = DsWave.nBitsPerSample / 8 * DsWave.nChannels
    DsWave.lAvgBytesPerSec = DsWave.lSamplesPerSec * DsWave.nBlockAlign

    'Hangok betöltése a memóriába
    Set hangok(0) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\nulla.wav", DsDesc, DsWave)
    Set hangok(1) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\egy.wav", DsDesc, DsWave)
    Set hangok(2) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\ketto.wav", DsDesc, DsWave)
    Set hangok(3) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\harom.wav", DsDesc, DsWave)
    Set hangok(4) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\negy.wav", DsDesc, DsWave)
    Set hangok(5) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\ot.wav", DsDesc, DsWave)
    Set hangok(6) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\hat.wav", DsDesc, DsWave)
    Set hangok(7) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\het.wav", DsDesc, DsWave)
    Set hangok(8) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\nyolc.wav", DsDesc, DsWave)
    Set hangok(9) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\kilenc.wav", DsDesc, DsWave)
    Set hangok(10) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\tiz.wav", DsDesc, DsWave)
    Set hangok(11) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\tizen.wav", DsDesc, DsWave)
    Set hangok(12) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\husz.wav", DsDesc, DsWave)
    Set hangok(13) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\huszon.wav", DsDesc, DsWave)
    Set hangok(14) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\harminc.wav", DsDesc, DsWave)
    Set hangok(15) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\negyven.wav", DsDesc, DsWave)
    Set hangok(16) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\otven.wav", DsDesc, DsWave)
    Set hangok(17) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\hatvan.wav", DsDesc, DsWave)
    Set hangok(18) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\hetven.wav", DsDesc, DsWave)
    Set hangok(19) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\nyolcvan.wav", DsDesc, DsWave)
    Set hangok(20) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\kilencven.wav", DsDesc, DsWave)
    Set hangok(21) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\szaz.wav", DsDesc, DsWave)
    Set hangok(22) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\ezer.wav", DsDesc, DsWave)
    Set hangok(23) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\millio.wav", DsDesc, DsWave)
    Set hangok(24) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\billio.wav", DsDesc, DsWave)
    Set hangok(25) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\minusz.wav", DsDesc, DsWave)
    Set hangok(26) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\egesz.wav", DsDesc, DsWave)
    Set hangok(27) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\tized.wav", DsDesc, DsWave)
    Set hangok(28) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\szazad.wav", DsDesc, DsWave)
    Set hangok(29) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\ezred.wav", DsDesc, DsWave)
    Set hangok(30) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\tizezred.wav", DsDesc, DsWave)
    Set hangok(31) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\szazezred.wav", DsDesc, DsWave)
    Set hangok(32) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\milliomod.wav", DsDesc, DsWave)
    Set hangok(33) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\tizmilliomod.wav", DsDesc, DsWave)
    Set hangok(34) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\szazmilliomod.wav", DsDesc, DsWave)
    Set hangok(35) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\billiomod.wav", DsDesc, DsWave)
    Set hangok(36) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\tizbilliomod.wav", DsDesc, DsWave)
    
Exit Sub
hiba:
    Select Case Err.Number
        Case 432
            MsgBox "Az egyik hang fájl nem található. Kérem telepítsen újra!", vbInformation, "Inicializálási hiba"
            Unload Me
        Case Else
            MsgBox Err.Description, vbInformation, "A rendszer hibaüzenete: " & Err.Number
            Unload Me
    End Select
    End
End Sub
Public Function Kovetkezo()
    If mutato > 0 Then
            hangok(jatszando(mutato)).Play DSBPLAY_DEFAULT
            mutato = mutato - 1
        Else
            szunet.Enabled = False
            Unload Me
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    For i = 0 To 30
        Set hangok(i) = Nothing
    Next i
End Sub

Private Sub szunet_Timer()
    If jatszando(mutato) <> 51 Then
            Kovetkezo
        Else
            mutato = mutato - 1
            szunet_Timer
    End If
End Sub
Public Sub Szamot_hangga(szam As Double)
On Error GoTo hiba:
    Form_Load
    Olvaso.Cls
    Olvaso.Print (szam)

    Dim egesz_resz As String, tort_resz As String
    Dim tort_str As String, i As Integer
    
    TorolJatszando
    tort_str = CStr(szam)
    egesz_resz = tort_str
    
    For i = 1 To Len(tort_str)
        If Mid(tort_str, i, 1) = "." Or Mid(tort_str, i, 1) = "," Then
            egesz_resz = CLng(Mid(tort_str, 1, i - 1))
            tort_resz = Mid(tort_str, i + 1, Len(tort_str) - i)
            GoTo kilepes
        End If
    Next i
kilepes:
    If Len(tort_resz) > 0 And Len(tort_resz) <= 10 Then
        UjJatszando (26 + Len(tort_resz))
        Egesz (tort_resz)
        UjJatszando (26)
    End If
    Egesz (egesz_resz)
    If szam < 0 Then UjJatszando (25)
    
    szunet.Enabled = True
    mutato = jatszando(0)
    Exit Sub
hiba:
    MsgBox "Túl nagy számot adott meg.", vbInformation, "Longint Túlcsordulás"
    Unload Me
End Sub
Private Sub UjJatszando(Index As Byte)
    jatszando(0) = jatszando(0) + 1
    jatszando(jatszando(0)) = Index
End Sub
Private Sub TorolJatszando()
    For i = 0 To 100
        jatszando(i) = 0
    Next i
End Sub
Private Function UtolsoJatszando() As Byte
    UtolsoJatszando = jatszando(jatszando(0))
End Function
Private Sub UjEgyes(szam)
    If szam <> 0 Then
            UjJatszando (szam)
        Else
            UjJatszando 51
    End If
End Sub
Private Sub Egesz(szam As Long)
    On Error GoTo hiba
    Dim i As Integer
    Dim szam_str As String ', szoveg As String, kj As String
    Dim jegyek(1 To 10) As Byte ', atalakitott(1 To 10) As String, minus As String
    
    szam_str = CStr(Abs(szam))
    'TorolJatszando
    
    'a számjegyek sorrendjének felcserélése
    For i = 1 To 10
        jegyek(i) = 0
    Next i
    
    For i = 1 To Len(szam_str)
        jegyek(i) = Mid(szam_str, Len(szam_str) - i + 1, 1)
    Next i
    
    'Helyiértékes vizsgálat
    For i = 1 To Len(szam_str)
        Select Case i
            Case 1
                If jegyek(i) = 0 And Len(szam_str) = 1 Then
                       UjJatszando (0)
                    Else
                        UjEgyes (jegyek(i))
                End If
            
            Case 2, 5, 8
                If UtolsoJatszando = 51 And (jegyek(i) = 1 Or jegyek(i) = 2) Then
                        If jegyek(i) = 1 Then
                                UjJatszando (10)
                            Else
                                UjJatszando (12)
                        End If
                    Else
                        Select Case jegyek(i)
                            Case 0
                                UjJatszando (51)
                            Case 1
                                UjJatszando (11)
                            Case 2
                                UjJatszando (13)
                            Case Else
                                UjJatszando (11 + jegyek(i))
                        End Select
                End If
            Case 3, 6, 9
                If jegyek(i) = 1 Then
                        UjJatszando (21)
                    Else
                        If jegyek(i) <> 0 Then
                            UjJatszando (21)
                            UjEgyes jegyek(i)
                        End If
                End If
            Case 4
                If jegyek(i) <> 0 Or jegyek(i + 1) <> 0 Or jegyek(i + 2) <> 0 Then
                    UjJatszando (22)
                End If
                
                If (jegyek(i) = 1 And Len(szam_str) > 4) Or jegyek(i) <> 1 Then  'jegyek(i + 1) = 0 And jegyek(i + 2) = 0 Then
                       UjEgyes jegyek(i)
                End If
            Case 7
                If jegyek(i) <> 0 Or jegyek(i + 1) <> 0 Or jegyek(i + 2) <> 0 Then
                    UjJatszando (23)
                End If
                UjEgyes jegyek(i)
            Case 10
                UjJatszando (24)
                UjEgyes jegyek(i)
        End Select
kov:
    Next i
    
    'A szám negatív
    'If szam < 0 Then UjJatszando (25)
    
    'szunet.Enabled = True
    'mutato = jatszando(0)
    Exit Sub
hiba:
    MsgBox Err.Description, vbInformation, "A rendszer hibaüzenete:"
End Sub
