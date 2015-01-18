VERSION 5.00
Begin VB.Form frmCP2 
   BorderStyle     =   1  '단일 고정
   Caption         =   "코드 패턴 - I I7 IV IVm"
   ClientHeight    =   990
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6150
   Icon            =   "frmCP2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   6150
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton Command1 
      Caption         =   "듣기"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   5895
   End
   Begin VB.ComboBox cbPattern 
      Height          =   300
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox cbPattern 
      Height          =   300
      Index           =   2
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox cbPattern 
      Height          =   300
      Index           =   3
      Left            =   5040
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox cbPattern 
      Height          =   300
      Index           =   1
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lbLine 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "-"
      Height          =   180
      Index           =   0
      Left            =   1200
      TabIndex        =   6
      Top             =   150
      Width           =   90
   End
   Begin VB.Label lbLine 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "-"
      Height          =   180
      Index           =   1
      Left            =   3480
      TabIndex        =   5
      Top             =   150
      Width           =   90
   End
   Begin VB.Label lbLine 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "-"
      Height          =   180
      Index           =   2
      Left            =   4800
      TabIndex        =   4
      Top             =   150
      Width           =   90
   End
End
Attribute VB_Name = "frmCP2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i, z As Long
For z = 0 To 3
Play() = Split(lstDiatonic.ListItems(lstDiatonic.SelectedItem.Index).ListSubItems(1).Text, "/") 'Split Note
Note() = Split(CalcNote(Play(0), Play(1), ""), ",") 'Split Note

For i = 0 To UBound(Note)
Note(i) = Note(i) + (24)
Next i

If cbInv.ListIndex = 0 Then
ElseIf cbInv.ListIndex = 1 Then
    Note(0) = Note(0) + 12
    Note(1) = Note(1)
    Note(2) = Note(2)
ElseIf cbInv.ListIndex = 2 Then
    Note(0) = Note(0) + 12
    Note(1) = Note(1) + 12
    Note(2) = Note(2)
ElseIf cbInv.ListIndex = 3 Then
    Note(0) = Note(0) + 12
    Note(1) = Note(1) + 12
    Note(2) = Note(2) + 12
    Note(3) = Note(3)
End If

bBass = True
PlayScale UBound(Note), True, 70, False
Next z

End Sub

Private Sub Form_Load()
chkShape = False
cbPattern(0).Text = NotetoText(Key + 0)
cbPattern(1).AddItem NotetoText(Key + 0) & "7"
cbPattern(1).AddItem NotetoText(Key + 6) & "7"
cbPattern(1).AddItem NotetoText(Key + 7) & "m7 - " & NotetoText(Key + 0) & "7"
cbPattern(1).AddItem NotetoText(Key + 7) & "m7 - " & NotetoText(Key + 6) & "7"
cbPattern(1).AddItem NotetoText(Key) & "aug"
cbPattern(1).AddItem NotetoText(Key + 4) & "aug"
cbPattern(1).AddItem NotetoText(Key + 7) & "aug"
cbPattern(1).ListIndex = 0
cbPattern(2).Text = NotetoText(Key + 5)
cbPattern(3).AddItem NotetoText(Key + 5) & "m"
cbPattern(3).AddItem NotetoText(Key + 10) & "7"
cbPattern(3).AddItem NotetoText(Key + 5) & "#dim7"
cbPattern(3).ListIndex = 0
End Sub

