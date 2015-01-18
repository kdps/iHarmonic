VERSION 5.00
Begin VB.Form frmEditor 
   BorderStyle     =   1  '단일 고정
   Caption         =   "코드 편집"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   Icon            =   "frmEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   4695
   StartUpPosition =   2  '화면 가운데
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   120
      ScaleHeight     =   2265
      ScaleWidth      =   4425
      TabIndex        =   12
      Top             =   1200
      Width           =   4455
      Begin VB.Label Label3 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "0/0"
         Height          =   180
         Left            =   0
         TabIndex        =   15
         Top             =   2040
         Width           =   4455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   ">"
         Height          =   180
         Left            =   4320
         TabIndex        =   14
         Top             =   960
         Width           =   120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "<"
         Height          =   180
         Left            =   0
         TabIndex        =   13
         Top             =   960
         Width           =   120
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtInvention 
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2640
         TabIndex        =   7
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtKind 
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1080
         TabIndex        =   6
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtRoot 
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   855
      End
      Begin VB.ComboBox txtInvention2 
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
         ItemData        =   "frmEditor.frx":014A
         Left            =   3480
         List            =   "frmEditor.frx":015A
         TabIndex        =   4
         Text            =   "0"
         Top             =   480
         Width           =   855
      End
      Begin VB.Label labInv 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "베이스 :"
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
         Left            =   2640
         TabIndex        =   11
         Top             =   240
         Width           =   585
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
         Left            =   1080
         TabIndex        =   10
         Top             =   240
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
         TabIndex        =   9
         Top             =   240
         Width           =   420
      End
      Begin VB.Label labInv2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "전위 :"
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
         Left            =   3480
         TabIndex        =   8
         Top             =   240
         Width           =   420
      End
   End
   Begin VB.CommandButton cmdRevision 
      Caption         =   "교정(&R)"
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
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "취소(&C)"
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
      TabIndex        =   1
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "확인(&O)"
      Default         =   -1  'True
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
      Left            =   3600
      TabIndex        =   0
      Top             =   3600
      Width           =   975
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public EditVal

Public Sub AnalysChord(Chord As String)
'On Error Resume Next
Dim Chd As String
Dim Note() As String
Dim strInput As String

Chord = Replace(Chord, " ", "")
Note() = Split(Chord, "/")

If UBound(Note) = 2 Then
    If Mid$(Note(0), 2, 1) = "#" Or Mid$(Note(0), 2, 1) = "b" Then
        txtRoot = Mid$(Note(0), 1, 2)
        txtKind = Mid$(Note(0), 3, Len(Note(0)) - 2)
        txtInvention = Note(1)
    Else
        txtRoot = Mid$(Note(0), 1, 1)
        txtKind = Mid$(Note(0), 2, Len(Note(0)) - 1)
        txtInvention = Note(1)
    End If
ElseIf UBound(Note) = 1 Then
    If Mid$(Note(0), 2, 1) = "#" Or Mid$(Note(0), 2, 1) = "b" Then
        txtRoot = Mid$(Note(0), 1, 2)
        txtKind = Mid$(Note(0), 3, Len(Note(0)) - 2)
        txtInvention = Note(1)
    Else
        txtRoot = Mid$(Note(0), 1, 1)
        txtKind = Mid$(Note(0), 2, Len(Note(0)) - 1)
        txtInvention = Note(1)
    End If
End If
End Sub

Private Sub cmdCancel_Click()
mdi_frmMain.ActiveForm.Enabled = True
Unload Me
End Sub

Private Sub cmdOK_Click()
On Error Resume Next
Dim strInput As String
Dim n, l As Long

strInput = txtKind

If Not InStr(strInput, "(") = 0 Then
    n = InStr(strInput, "(") + 1
    l = Len(Mid$(strInput, n)) + 1
    strTention = Mid$(strInput, n, Len(strInput) - (n))
    mdi_frmMain.ActiveForm.lstChord.ListItems(EditVal).Text = txtRoot & txtKind & "/" & txtInvention
    mdi_frmMain.ActiveForm.lstChord.ListItems(EditVal).ListSubItems(1).Text = TexttoNote(txtRoot) & "/" & Mid$(txtKind, 1, Len(txtKind) - (l)) & "/" & strTention & "/" & TexttoNote(txtInvention) & "/" & txtInvention2
Else
    mdi_frmMain.ActiveForm.lstChord.ListItems(EditVal).Text = txtRoot & txtKind & "/" & txtInvention
    mdi_frmMain.ActiveForm.lstChord.ListItems(EditVal).ListSubItems(1).Text = TexttoNote(txtRoot) & "/" & txtKind & "/" & "0" & "/" & TexttoNote(txtInvention) & "/" & txtInvention2
End If

Unload Me
End Sub

Private Sub cmdRevision_Click()
On Error Resume Next
Dim i As Long
If Not mdi_frmMain.ActiveForm Is Nothing Then
    For i = 1 To mdi_frmMain.ActiveForm.lstChord.ListItems.Count
        If Not Mid$(mdi_frmMain.ActiveForm.lstChord.ListItems(i).ListSubItems(1).Text, 1, 4) = "Time" Or Mid$(mdi_frmMain.ActiveForm.lstChord.ListItems(i).ListSubItems(1).Text, 1, 7) = "Comment" Then
            RevisionChord mdi_frmMain.ActiveForm.lstChord.ListItems(i).Text, i
        End If
    Next i
    MsgBox "완료", vbExclamation, "교정 도구"
    Unload Me
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
mdi_frmMain.ActiveForm.Enabled = True
End Sub
