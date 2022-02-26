VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 2  'vbComplexBound
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "olvasasa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Sub Class_Initialize()
    Debug.Print "Elindultam..."
End Sub

Public Sub Szamot_hangga(szam As Double)
On Error GoTo hiba
    Olvaso.Szamot_hangga (szam)
Exit Sub
hiba:
    MsgBox "Hibás értéket adott meg!", vbCritical, "Értelmezési hiba!"
End Sub

Public Function nevjegy() As String
    nevjegy = "Magyar számok olvasása v" & App.Major & "." & App.Minor & "." & App.Revision & ". Freeware ActiveX - Muráti Ákos 2003."
End Function
Public Property Get Kesleltet() As Integer
    Kesleltet = Olvaso.szunet.Index
End Property
Public Property Let Kesleltet(uj As Integer)
    If uj = 0 Then Exit Property
    Olvaso.szunet.Interval = uj
End Property

