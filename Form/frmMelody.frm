VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMelody 
   Caption         =   "멜로디 편집기"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  '화면 가운데
   Begin VB.Timer tmrPlay 
      Left            =   0
      Top             =   0
   End
   Begin MSComctlLib.ListView lstMelody 
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   2143
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmMelody"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
lstMelody.Width = Me.ScaleWidth
lstMelody.Height = Me.ScaleHeight
End Sub

Private Sub lstMelody_Click()
Play
End Sub

Private Sub Play()
On Error Resume Next
Dim i As Long
Dim getnote() As String
For i = 1 To lstMelody.ListItems.Count

    Timer Tempo
    
    getnote = Split(lstMelody.ListItems(i).ListSubItems(1).Text, ",")
    temp = getnote(0)
    If temp < 0 Then
        'key-up
        PlayNote Abs(temp)
    Else
        'key-down
        StopNote temp
    End If
    
    playinc = playinc + 1
    getnote = Split(lstMelody.ListItems(i).ListSubItems(1).Text, ",")
    
    temp = getnote(1) * 50
    
    If temp = 0 Then 'a 0 means another event happens at the exact same time, so do it now!
        'tmrPlayBack_Timer
        Exit Sub
    End If
    
Next i
End Sub
