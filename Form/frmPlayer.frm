VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmPlayer 
   BorderStyle     =   1  '단일 고정
   Caption         =   "플레이어"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6975
   BeginProperty Font 
      Name            =   "맑은 고딕"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPlayer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   6975
   StartUpPosition =   2  '화면 가운데
   Begin VB.ComboBox cbStyle 
      Height          =   315
      Left            =   120
      TabIndex        =   17
      Text            =   "Combo1"
      Top             =   120
      Width           =   6735
   End
   Begin VB.CheckBox chkOpen 
      Caption         =   "오픈 보이싱"
      Height          =   255
      Left            =   3000
      TabIndex        =   16
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CheckBox chkVoice 
      Caption         =   "자동 보이싱"
      Height          =   255
      Left            =   1680
      TabIndex        =   14
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CheckBox chkDrum 
      Caption         =   "드럼"
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   2760
      Value           =   1  '확인
      Width           =   855
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "재생(&P)"
      Height          =   735
      Left            =   5040
      TabIndex        =   4
      Top             =   2280
      Width           =   1815
   End
   Begin ComctlLib.Slider sldPlay 
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   2280
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   661
      _Version        =   327682
      TickStyle       =   3
   End
   Begin VB.Frame Frame1 
      Caption         =   "설정"
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   6735
      Begin VB.Timer tmrPlay 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   0
         Top             =   0
      End
      Begin ComctlLib.Slider sldOctave 
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   240
         Width           =   5800
         _ExtentX        =   10239
         _ExtentY        =   450
         _Version        =   327682
         Min             =   1
         Max             =   6
         SelStart        =   1
         TickStyle       =   3
         Value           =   1
      End
      Begin ComctlLib.Slider sldPitch 
         Height          =   255
         Left            =   840
         TabIndex        =   7
         Top             =   600
         Width           =   5800
         _ExtentX        =   10239
         _ExtentY        =   450
         _Version        =   327682
         Max             =   200
         SelStart        =   23
         TickStyle       =   3
         Value           =   23
      End
      Begin ComctlLib.Slider sldVelocity 
         Height          =   255
         Left            =   840
         TabIndex        =   8
         Top             =   960
         Width           =   5800
         _ExtentX        =   10239
         _ExtentY        =   450
         _Version        =   327682
         Max             =   160
         SelStart        =   127
         TickStyle       =   3
         Value           =   127
      End
      Begin ComctlLib.Slider sldTempo 
         Height          =   255
         Left            =   840
         TabIndex        =   9
         Top             =   1320
         Width           =   5800
         _ExtentX        =   10239
         _ExtentY        =   450
         _Version        =   327682
         Min             =   170
         Max             =   350
         SelStart        =   250
         TickStyle       =   3
         Value           =   250
      End
      Begin VB.Label labOctave 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "옥타브 :"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   585
      End
      Begin VB.Label labPitch 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "음정 :"
         Height          =   195
         Left            =   300
         TabIndex        =   12
         Top             =   600
         Width           =   420
      End
      Begin VB.Label labVelocity 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "세기 :"
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   960
         Width           =   420
      End
      Begin VB.Label labSpeed 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "속도 :"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   420
      End
   End
   Begin VB.ComboBox cbInstrument 
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   600
      TabIndex        =   2
      Text            =   "Default"
      Top             =   3360
      Width           =   6255
   End
   Begin VB.CheckBox chkBass 
      Caption         =   "베이스"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Value           =   1  '확인
      Width           =   855
   End
   Begin VB.Label labInstrument 
      Alignment       =   1  '오른쪽 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "악기 :"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   420
   End
End
Attribute VB_Name = "frmPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Style As Integer
Dim NoteBass, NoteBass2, TopBass, strStyle As Long
Private ListItem() As String
Dim Backup() As String
Dim strPlay, lStyle As Long
Dim midiMsg As Long
Dim cntPlay, locPlay
Private Sub cbInstrument_Click()
InitializInstrument cbInstrument.ListIndex
End Sub

Private Sub cbStyle_Click()
modPlaySong.Style = cbStyle.ListIndex
End Sub

Private Sub cmdPlay_Click()
On Error Resume Next
'If modPlaySong.Active = False Then
'    frmPlayer.sldPlay.value = 0
'    frmPlayer.sldPlay.Max = mdi_frmMain.ActiveForm.lstChord.ListItems.Count
'    modPlaySong.t = sldPlay.value + 1
'    modPlaySong.Style = cbStyle.ListIndex + 1
'    modPlaySong.PlaySong
'Else
'    modPlaySong.z = mdi_frmMain.ActiveForm.lstChord.ListItems.Count - 1
'    modPlaySong.Style = cbStyle.ListIndex + 1
'    modPlaySong.PlaySong
'    modPlaySong.Active = False
'End If
If tmrPlay.Enabled = False Then
    cmdPlay.Caption = "정지(&S)"
    cntPlay = 1
    locPlay = 1
    tmrPlay.Enabled = True
Else
    cmdPlay.Caption = "재생(&P)"
   tmrPlay.Enabled = False
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim A As Long

cbStyle.AddItem "Piano - Basic"
cbStyle.AddItem "Jazz - BossaNova"
cbStyle.AddItem "Pop - Game"
cbStyle.AddItem "Piano - Jazz Piano 1"

cbStyle.ListIndex = modPlaySong.Style

For A = 0 To 128 'load up instrument names
    cbInstrument.AddItem Format(A, "000") & "   " & LoadResString(A)
Next A

cbInstrument.ListIndex = 0

End Sub

Private Sub sldPitch_change()
Pitch = sldPitch.value
End Sub

Private Sub sldPlay_Change()
On Error Resume Next
For i = 1 To mdi_frmMain.ActiveForm.lstChord.ListItems.Count - 1
    mdi_frmMain.ActiveForm.lstChord.ListItems(i).Bold = False
Next i

mdi_frmMain.ActiveForm.lstChord.ListItems(sldPlay.value + 1).Bold = True
End Sub

Private Sub sldTempo_change()
modPlay.Tempo = sldTempo.value
End Sub

Private Sub sldVelocity_change()
Velocity = sldVelocity.value
End Sub

Private Sub tmrPlay_Timer()
'On Error Resume Next
cntPlay = cntPlay + 1

If locPlay = 0 Then
    locPlay = 1
    cntPlay = 1
End If

If locPlay = mdi_frmMain.ActiveForm.lstChord.ListItems.Count Then
    tmrPlay.Enabled = False
End If

If mdi_frmMain.ActiveForm.lstChord.ListItems(locPlay).ListSubItems(1).Text <> "" Then
    ListItem() = Split(mdi_frmMain.ActiveForm.lstChord.ListItems(locPlay).ListSubItems(1).Text, "/") 'Split Note
Else
    
End If

If locPlay > 2 Then
    mdi_frmMain.ActiveForm.lstChord.ListItems(locPlay - 2).Bold = False
    mdi_frmMain.ActiveForm.lstChord.ListItems(locPlay - 1).Bold = True
    frmPlayer.sldPlay.value = locPlay
End If

Select Case ListItem(0)
    Case "Style"
        Style = ListItem(1)
        GoTo Pass
    Case "Comment"
        Select Case ListItem(3)
            Case "[:"
                s = locPlay
                GoTo Pass
            Case ":]"
                If Repeat = False Then
                    e = locPlay - 1
                    locPlay = s - 1
                    Repeat = True
                End If
            Case "┌"
                If Repeat = True Then
                    locPlay = e
                End If
            Case "A"
                lStyle = 0
                strStyle = 0
            Case "B"
                lStyle = 1
                strStyle = 0
            Case "C"
                lStyle = 2
                strStyle = 0
            Case "D"
                lStyle = 3
                strStyle = 0
            Case Else
                GoTo Pass
        End Select
        GoTo Pass
    Case "Time"
        p = Str(ListItem(1)) - 1
        GoTo Pass
End Select

mdi_frmMain.ActiveForm.labChord = mdi_frmMain.ActiveForm.lstChord.ListItems(locPlay).Text '& chdClassic

'If (ListItem(0) - Key) > 7 Then
'    ListItem(0) = ListItem(0) - 12
'ElseIf (ListItem(0) - Key) < 0 Then
'    ListItem(0) = ListItem(0) + 12
'End If

Note() = Split(CalcNote((12 * (frmPlayer.sldOctave.value + 1)) + ListItem(0), ListItem(1), ListItem(2)), ",") 'Split Note

i = UBound(Note)

If UBound(ListItem) > 2 Then
    'NoteBass = 12 * (frmPlayer.sldOctave.value + 1) + ListItem(3)
    'NoteBass2 = 12 * (frmPlayer.sldOctave.value - 1) + Note(2)
    'TopBass = 12 * (frmPlayer.sldOctave.value + 1) + ListItem(0)

    'If ((ListItem(0) - Key) < 1 And ListItem(3) = 1) Or ((ListItem(0) - Key) > 7 And ListItem(3) = 1) Then
    '    ListItem(0) = ListItem(0) - 12
    'ElseIf ((ListItem(0) - Key)) > 7 And ListItem(3) = 2 Then
    '    ListItem(0) = ListItem(0) - 12
    'End If

    If (ListItem(3) - ListItem(0)) < 0 Then
        Inv = ListItem(0) - ListItem(3)
    Else
        Inv = ListItem(3) - ListItem(0)
    End If
    
End If

If UBound(ListItem) > 3 And frmPlayer.chkVoice.value = 1 Then

    If Not Inv = 0 And Inv < 5 And Len(Inv) Then   'Invention 1 C/E 1~4
        If i = 2 Then
            If frmPlayer.chkOpen.value = 1 Then
                ReDim Backup(3)
                Note(0) = Note(0) + 12 'C
                'E
                Note(2) = Note(2) 'G
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Note(0) = Backup(2)
                Note(1) = Backup(3) + 12
                Note(2) = Backup(1)
            Else
                ReDim Backup(3)
                Note(0) = Note(0) + 12 'C
                'E
                Note(2) = Note(2) 'G
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Note(0) = Backup(2)
                Note(1) = Backup(3)
                Note(2) = Backup(1)
            End If
        ElseIf i = 3 Then
            If frmPlayer.chkOpen.value = 1 Then
                ReDim Backup(4)
                Note(0) = Note(0) + 12 'C
                'E
                Note(2) = Note(2) 'G
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Backup(4) = Note(3)
                Note(0) = Backup(2)
                Note(1) = Backup(3) + 12
                Note(2) = Backup(1)
                Note(3) = Backup(4) + 12
            Else
                ReDim Backup(4)
                Note(0) = Note(0) + 12 'C
                'E
                Note(2) = Note(2) 'G
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Backup(4) = Note(3)
                Note(0) = Backup(2)
                Note(1) = Backup(3)
                Note(2) = Backup(1)
                Note(3) = Backup(4)
            End If
        End If
    End If
    
    If Not Inv = 4 And Inv < 8 And Inv > 4 Then  'Invention 2 C/G 5~7
        If i = 2 Then
            If frmPlayer.chkOpen.value = 1 Then
                ReDim Backup(3)
                Note(0) = Note(0) + 12 'C
                Note(1) = Note(1) + 12 'E
                'G
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Note(0) = Backup(3) - 12
                Note(1) = Backup(1)
                Note(2) = Backup(2) - 12
            Else
                ReDim Backup(3)
                Note(0) = Note(0) + 12 'C
                Note(1) = Note(1) + 12 'E
                'G
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Note(0) = Backup(3)
                Note(1) = Backup(1)
                Note(2) = Backup(2)
            End If
        ElseIf i = 3 Then
            If frmPlayer.chkOpen.value = 1 Then
                ReDim Backup(4)
                Note(0) = Note(0) + 12 'C
                Note(1) = Note(1) + 12 'E
                'G
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Backup(4) = Note(3)
                Note(0) = Backup(3) - 12
                Note(1) = Backup(1)
                Note(2) = Backup(2) - 12
                Note(3) = Backup(4)
            Else
                ReDim Backup(4)
                Note(0) = Note(0) + 12 'C
                Note(1) = Note(1) + 12 'E
                'G
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Backup(4) = Note(3)
                Note(0) = Backup(3)
                Note(1) = Backup(1)
                Note(2) = Backup(2)
                Note(3) = Backup(4)
            End If
        End If
    End If
    
    If Not Inv = 7 And Inv < 12 And Inv > 7 Then     'Invention 3 CM7/B 8~11
        If i > 2 Then
            If frmPlayer.chkOpen.value = 1 Then
                ReDim Backup(4)
                Note(0) = Note(0) 'C
                Note(1) = Note(1) 'E
                Note(2) = Note(2) 'G
                Note(3) = Note(3) - 12 'B
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Backup(4) = Note(3)
                Note(0) = Backup(4) - 12
                Note(1) = Backup(1)
                Note(2) = Backup(2) - 12
                Note(3) = Backup(3)
            Else
                ReDim Backup(4)
                Note(0) = Note(0) 'C
                Note(1) = Note(1) 'E
                Note(2) = Note(2) 'G
                Note(3) = Note(3) - 12 'B
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Backup(4) = Note(3)
                Note(0) = Backup(4)
                Note(1) = Backup(1)
                Note(2) = Backup(2)
                Note(3) = Backup(3)
            End If
        End If
    End If
    
End If

If UBound(ListItem) > 3 And frmPlayer.chkVoice.value = 0 Then

    If ListItem(4) = "1" Then 'Invention 1 C/E
        If i = 2 Then
            If frmPlayer.chkOpen.value = 1 Then
                ReDim Backup(3)
                Note(0) = Note(0) + 12 'C
                Note(2) = Note(2) 'G
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Note(0) = Backup(2)
                Note(1) = Backup(3) + 12
                Note(2) = Backup(1)
            Else
                ReDim Backup(3)
                Note(0) = Note(0) + 12 'C
                Note(2) = Note(2) 'G
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Note(0) = Backup(2)
                Note(1) = Backup(3)
                Note(2) = Backup(1)
            End If
        ElseIf i = 3 Then
            If frmPlayer.chkOpen.value = 1 Then
                ReDim Backup(4)
                Note(0) = Note(0) + 12 'C
                Note(2) = Note(2) 'G
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Backup(4) = Note(3)
                Note(0) = Backup(2)
                Note(1) = Backup(3) + 12
                Note(2) = Backup(1)
                Note(3) = Backup(4) + 12
            Else
                ReDim Backup(4)
                Note(0) = Note(0) + 12 'C
                Note(2) = Note(2) 'G
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Backup(4) = Note(3)
                Note(0) = Backup(2)
                Note(1) = Backup(3)
                Note(2) = Backup(1)
                Note(3) = Backup(4)
            End If
        End If
    End If
    
    If ListItem(4) = "2" Then  'Invention 2 C/G
        If i = 2 Then
            If frmPlayer.chkOpen.value = 1 Then
                ReDim Backup(3)
                Note(0) = Note(0) + 12 'C
                Note(1) = Note(1) + 12 'E
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Note(0) = Backup(3) - 12
                Note(1) = Backup(1)
                Note(2) = Backup(2) - 12
            Else
                ReDim Backup(3)
                Note(0) = Note(0) + 12 'C
                Note(1) = Note(1) + 12 'E
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Note(0) = Backup(3) - 12
                Note(1) = Backup(1) - 12
                Note(2) = Backup(2) - 12
            End If
        ElseIf i = 3 Then
            If frmPlayer.chkOpen.value = 1 Then
                ReDim Backup(4)
                Note(0) = Note(0) + 12 'C
                Note(1) = Note(1) + 12 'E
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Backup(4) = Note(3)
                Note(0) = Backup(3) - 12
                Note(1) = Backup(1)
                Note(2) = Backup(2) - 12
                Note(3) = Backup(4)
            Else
                ReDim Backup(4)
                Note(0) = Note(0) + 12 'C
                Note(1) = Note(1) + 12 'E
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Backup(4) = Note(3)
                Note(0) = Backup(3) - 12
                Note(1) = Backup(1) - 12
                Note(2) = Backup(2) - 12
                Note(3) = Backup(4) - 12
            End If
        End If
    End If
    
    If ListItem(4) = "3" Then     'Invention 3 CM7/B
        If i > 2 Then
            If frmPlayer.chkOpen.value = 1 Then
                ReDim Backup(4)
                Note(0) = Note(0) 'C
                Note(1) = Note(1) 'E
                Note(2) = Note(2) 'G
                Note(3) = Note(3) - 12 'B
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Backup(4) = Note(3)
                Note(0) = Backup(4) - 12
                Note(1) = Backup(1)
                Note(2) = Backup(2) - 12
                Note(3) = Backup(3)
            Else
                ReDim Backup(4)
                Note(0) = Note(0) 'C
                Note(1) = Note(1) 'E
                Note(2) = Note(2) 'G
                Note(3) = Note(3) - 12 'B
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Backup(4) = Note(3)
                Note(0) = Backup(4) - 12
                Note(1) = Backup(1) - 12
                Note(2) = Backup(2) - 12
                Note(3) = Backup(3) - 12
            End If
        End If
    End If
End If

If UBound(ListItem) = "3" Then 'Invention 1 C/C
    If frmPlayer.chkOpen.value = 1 Then
        If i = 2 Then
            Note(1) = Note(1) + 12
        ElseIf i = 3 Then
            Note(1) = Note(1) + 12
            Note(2) = Note(2) + 12
        End If
    End If
End If

Select Case cntPlay
    Case 1
        Select Case UBound(Note)
        Case 2
            PlayNote Note(1)
            PlayNote Note(2)
        Case 3
            PlayNote Note(1)
            PlayNote Note(2)
            PlayNote Note(3)
        End Select
    Case 2
        Select Case UBound(Note)
        Case 2
            StopNote Note(1)
            StopNote Note(2)
        Case 3
            StopNote Note(1)
            StopNote Note(2)
            StopNote Note(3)
        End Select
    Case 3
        PlayNote Note(0)
    Case 4
        StopNote Note(0)
        GoTo NextPlay
End Select

Exit Sub
Pass:
Exit Sub
NextPlay:
locPlay = locPlay + 1
cntPlay = 0
End Sub
