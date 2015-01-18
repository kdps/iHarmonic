VERSION 5.00
Begin VB.Form frmKey 
   BorderStyle     =   1  '단일 고정
   Caption         =   "키 / 스케일"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5415
   Icon            =   "frmKey.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5415
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   60
      Left            =   5160
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   43
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   1440
      Width           =   135
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   56
      Left            =   4560
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   1440
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   58
      Left            =   4800
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   1440
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   49
      Left            =   3615
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   1440
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   51
      Left            =   3855
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   1440
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   54
      Left            =   4335
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1440
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   46
      Left            =   3135
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1440
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   44
      Left            =   2895
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1440
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   42
      Left            =   2655
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1440
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   39
      Left            =   2175
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1440
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   37
      Left            =   1935
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1440
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   34
      Left            =   1455
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1440
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   32
      Left            =   1215
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1440
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   30
      Left            =   975
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1440
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   27
      Left            =   510
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1440
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   25
      Left            =   255
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1440
      Width           =   195
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "적용(&A)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdListen 
      Caption         =   "미리듣기(&P)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   5175
      Begin VB.ComboBox cbType 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmKey.frx":1272
         Left            =   2520
         List            =   "frmKey.frx":12A6
         TabIndex        =   2
         Top             =   240
         Width           =   2535
      End
      Begin VB.ComboBox cbRoot 
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmKey.frx":1354
         Left            =   720
         List            =   "frmKey.frx":137C
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label labKind 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "종류 :"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1920
         TabIndex        =   6
         Top             =   315
         Width           =   420
      End
      Begin VB.Label labRoot 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "근음 :"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   315
         Width           =   420
      End
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   59
      Left            =   4905
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   41
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   57
      Left            =   4665
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   42
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   55
      Left            =   4440
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   34
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   53
      Left            =   4200
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   35
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   52
      Left            =   3960
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   36
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   50
      Left            =   3720
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   37
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   48
      Left            =   3480
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   38
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   47
      Left            =   3240
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   24
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   45
      Left            =   3000
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   25
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   43
      Left            =   2760
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   26
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   41
      Left            =   2520
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   27
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   40
      Left            =   2280
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   28
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   38
      Left            =   2040
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   29
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   36
      Left            =   1800
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   30
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   35
      Left            =   1560
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   12
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   33
      Left            =   1320
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   13
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   31
      Left            =   1080
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   14
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   29
      Left            =   840
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   15
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   28
      Left            =   600
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   16
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   26
      Left            =   360
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   17
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   24
      Left            =   120
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   18
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   1440
      Width           =   255
   End
End
Attribute VB_Name = "frmKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbRoot_Click()
cbType.Enabled = True
End Sub

Public Sub cbType_Click()
cmdListen.Enabled = True
cmdOK.Enabled = True

Dim Note() As String
Dim i As Integer

Note() = Split(CalcKey((24) + cbRoot.ListIndex + 1, cbType.Text), ",") 'Split Note

For i = 24 To 60
    If pKey(i).Tag = "1" Then
        pKey(i).BackColor = vbWhite
    Else
        pKey(i).BackColor = vbBlack
    End If
Next i

For i = 0 To UBound(Note)
    pKey(Note(i) - 1).BackColor = vbRed
Next i

End Sub

Private Sub cmdListen_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

Dim Note() As String
Dim i As Integer

Note() = Split(CalcKey((24) + cbRoot.ListIndex + 1, cbType.Text), ",") 'Split Note

For i = 0 To UBound(Note)
    PlayNote Note(i)
    Timer 300
    StopNote Note(i)
Next i

End Sub

Private Sub cmdOK_Click()
On Error Resume Next
Dim Note() As String
Dim i As Integer

Note() = Split(CalcKey(cbRoot.ListIndex + 1, cbType.Text), ",")  'Split Note
Key = cbRoot.ListIndex + 1
frmCustom.cbRoot.Clear
mdi_frmMain.ActiveForm.labkey = NotetoText(Key) & "Key"
mMScale = cbType.Text

For i = 0 To UBound(Note)
    If Note(i) = 1 Then
        frmCustom.cbRoot.AddItem "C", i
    ElseIf Note(i) = 2 Then
        frmCustom.cbRoot.AddItem "C#", i
    ElseIf Note(i) = 3 Then
        frmCustom.cbRoot.AddItem "D", i
    ElseIf Note(i) = 4 Then
        frmCustom.cbRoot.AddItem "D#", i
    ElseIf Note(i) = 5 Then
        frmCustom.cbRoot.AddItem "E", i
    ElseIf Note(i) = 6 Then
        frmCustom.cbRoot.AddItem "F", i
    ElseIf Note(i) = 7 Then
        frmCustom.cbRoot.AddItem "F#", i
    ElseIf Note(i) = 8 Then
        frmCustom.cbRoot.AddItem "G", i
    ElseIf Note(i) = 9 Then
        frmCustom.cbRoot.AddItem "G#", i
    ElseIf Note(i) = 10 Then
        frmCustom.cbRoot.AddItem "A", i
    ElseIf Note(i) = 11 Then
        frmCustom.cbRoot.AddItem "A#", i
    ElseIf Note(i) = 12 Then
        frmCustom.cbRoot.AddItem "B", i
    ElseIf Note(i) = 13 Then
        frmCustom.cbRoot.AddItem "C", i
    ElseIf Note(i) = 14 Then
        frmCustom.cbRoot.AddItem "C#", i
    ElseIf Note(i) = 15 Then
        frmCustom.cbRoot.AddItem "D", i
    ElseIf Note(i) = 16 Then
        frmCustom.cbRoot.AddItem "D#", i
    ElseIf Note(i) = 17 Then
        frmCustom.cbRoot.AddItem "E", i
    ElseIf Note(i) = 18 Then
        frmCustom.cbRoot.AddItem "F", i
    ElseIf Note(i) = 19 Then
        frmCustom.cbRoot.AddItem "F#", i
    ElseIf Note(i) = 20 Then
        frmCustom.cbRoot.AddItem "G", i
    ElseIf Note(i) = 21 Then
        frmCustom.cbRoot.AddItem "G#", i
    ElseIf Note(i) = 22 Then
        frmCustom.cbRoot.AddItem "A", i
    ElseIf Note(i) = 23 Then
        frmCustom.cbRoot.AddItem "A#", i
    ElseIf Note(i) = 24 Then
        frmCustom.cbRoot.AddItem "B", i
    End If
Next i

    If cbType.ListIndex = 0 Then
        mdi_frmMain.mnuModal.Enabled = True
        mdi_frmMain.mnuPassDimList.Enabled = True
    Else
        mdi_frmMain.mnuModal.Enabled = False
        mdi_frmMain.mnuPassDimList.Enabled = False
    End If
    
End Sub

Private Sub Form_Load()

If mMScale <> "" Then
    cbType.Enabled = True
    cbType.Text = mMScale
    cbRoot.ListIndex = Key - 1
    cbType_Click
End If

End Sub

