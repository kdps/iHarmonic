VERSION 5.00
Begin VB.Form frmMidi 
   BorderStyle     =   1  '단일 고정
   Caption         =   "미디"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   9015
   Icon            =   "frmMidi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   9015
   StartUpPosition =   2  '화면 가운데
   Begin VB.CheckBox chkDrum 
      Caption         =   "드럼"
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   2445
      Value           =   1  '확인
      Width           =   855
   End
   Begin VB.CheckBox chkBass 
      Caption         =   "베이스"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2445
      Value           =   1  '확인
      Width           =   855
   End
   Begin VB.ComboBox cbInstrument 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6480
      TabIndex        =   0
      Text            =   "Default"
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label laste 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Height          =   180
      Left            =   1800
      TabIndex        =   4
      Top             =   2445
      Width           =   3540
   End
   Begin VB.Label labInstrument 
      Alignment       =   1  '오른쪽 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "악기 :"
      Height          =   180
      Left            =   5940
      TabIndex        =   1
      Top             =   2445
      Width           =   480
   End
End
Attribute VB_Name = "frmMidi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbInstrument_Click()

InitializInstrument cbInstrument.ListIndex

End Sub

Private Sub Form_Load()

On Error Resume Next

If Style = 1 Then
    laste = "피아노 재즈"
ElseIf Style = 2 Then
    laste = "4/4 보사노바 리듬"
ElseIf Style = 3 Then
    laste = "4/4 재즈"
ElseIf Style = 4 Then
    laste = "동요"
End If

SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE

For a = 0 To 128 'load up instrument names
    cbInstrument.AddItem Format(a, "000") & "   " & LoadResString(a)
Next a

End Sub

Private Sub sldNote_change()
noteLong = sldNote.Value
End Sub

Private Sub sldPitch_change()
Pitch = sldPitch.Value
End Sub

Private Sub sldTempo_change()
Tempo = sldTempo.Value
End Sub

Private Sub sldVelocity_change()
Velocity = sldVelocity.Value
End Sub
