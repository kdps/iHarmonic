VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmDiatonic 
   BorderStyle     =   1  '단일 고정
   Caption         =   "다이아토닉 코드"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   6630
   Icon            =   "frmDiatonic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6630
   StartUpPosition =   2  '화면 가운데
   Begin MSComctlLib.ImageList ImageList 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiatonic.frx":1272
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiatonic.frx":180C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiatonic.frx":1DA6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   5895
      Begin MSComctlLib.ListView lstDiatonic 
         Height          =   2535
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   4471
         View            =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList"
         SmallIcons      =   "ImageList"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.ComboBox cbInv 
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
         ItemData        =   "frmDiatonic.frx":2340
         Left            =   120
         List            =   "frmDiatonic.frx":2350
         TabIndex        =   4
         Top             =   2880
         Width           =   5655
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   4095
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   7223
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "다이아토닉 코드"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "맑은 고딕"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "입력(&I)"
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
      Left            =   5280
      TabIndex        =   1
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "미리듣기(&P)"
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
      TabIndex        =   0
      Top             =   4440
      Width           =   1215
   End
End
Attribute VB_Name = "frmDiatonic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdInsert_Click()

On Error Resume Next
If InStr(lstDiatonic.SelectedItem.Text, "#") Or InStr(lstDiatonic.SelectedItem.Text, "b") Then
    mdi_frmMain.ActiveForm.lstChord.ListItems.Add , , Mid$(lstDiatonic.SelectedItem.Text, 1, 2) & Mid$(lstDiatonic.SelectedItem.Text, 3, (Len(lstDiatonic.SelectedItem.Text)) - 2) & "/" & Mid$(lstDiatonic.SelectedItem.Text, 1, 2)
    mdi_frmMain.ActiveForm.lstChord.ListItems(mdi_frmMain.ActiveForm.lstChord.ListItems.Count).ListSubItems.Add.Text = TexttoNote(Mid$(lstDiatonic.SelectedItem.Text, 1, 2)) & "/" & (Mid$(lstDiatonic.SelectedItem.Text, 3, (Len(lstDiatonic.SelectedItem.Text)) - 2)) & "/" & "0" & "/" & TexttoNote(Mid$(lstDiatonic.SelectedItem.Text, 1, 2))
Else
    mdi_frmMain.ActiveForm.lstChord.ListItems.Add , , Mid$(lstDiatonic.SelectedItem.Text, 1, 1) & Mid$(lstDiatonic.SelectedItem.Text, 2, (Len(lstDiatonic.SelectedItem.Text)) - 1) & "/" & Mid$(lstDiatonic.SelectedItem.Text, 1, 1)
    mdi_frmMain.ActiveForm.lstChord.ListItems(mdi_frmMain.ActiveForm.lstChord.ListItems.Count).ListSubItems.Add.Text = TexttoNote(Mid$(lstDiatonic.SelectedItem.Text, 1, 1)) & "/" & (Mid$(lstDiatonic.SelectedItem.Text, 2, Len(lstDiatonic.SelectedItem.Text) - 1)) & "/" & "0" & "/" & TexttoNote(Mid$(lstDiatonic.SelectedItem.Text, 1, 1))
End If

End Sub

Private Sub Form_Load()

Dim Note() As String
Dim i As Integer

If mMScale = "Major" Then
    
    Note() = Split(CalcKey(Key, "Major"), ",") 'Split Note
    
    lstDiatonic.ListItems.Clear

    lstDiatonic.ListItems.Add , , NotetoText(Note(0)), , 1
    lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(0) & "/" & "/" & "0"
    
    lstDiatonic.ListItems.Add , , NotetoText(Note(0)) & "M7", , 1
    lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(0) & "/" & "M7" & "/" & "0"
    
    lstDiatonic.ListItems.Add , , NotetoText(Note(0)) & "6", , 1
    lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(0) & "/" & "6" & "/" & "0"
    
    lstDiatonic.ListItems.Add , , NotetoText(Note(2)) & "m", , 1
    lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(2) & "/" & "m" & "/" & "0"
    
    lstDiatonic.ListItems.Add , , NotetoText(Note(2)) & "m7", , 1
    lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(2) & "/" & "m7" & "/" & "0"
    
    lstDiatonic.ListItems.Add , , NotetoText(Note(5)) & "m", , 1
    lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(5) & "/" & "m" & "/" & "0"
    
    lstDiatonic.ListItems.Add , , NotetoText(Note(5)) & "m7", , 1
    lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(5) & "/" & "m7" & "/" & "0"
    
    lstDiatonic.ListItems.Add , , NotetoText(7) & "dim", , 1
    lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , "6" & "/" & "dim" & "/" & "0"
    
    lstDiatonic.ListItems.Add , , NotetoText(7) & "m7b5", , 1
    lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , "6" & "/" & "m7b5" & "/" & "0"
    
    lstDiatonic.ListItems.Add , , NotetoText(Note(1)) & "m", , 2
    lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(1) & "/" & "m" & "/" & "0"
    
    lstDiatonic.ListItems.Add , , NotetoText(Note(1)) & "m7", , 2
    lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(1) & "/" & "m7" & "/" & "0"
    
    lstDiatonic.ListItems.Add , , NotetoText(Note(3)), , 2
    lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(3) & "/" & "/" & "0"
    
    lstDiatonic.ListItems.Add , , NotetoText(Note(3)) & "M7", , 2
    lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(3) & "/" & "M7" & "/" & "0"
    
    lstDiatonic.ListItems.Add , , NotetoText(Note(4)), , 3
    lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(4) & "/" & "/" & "0"
    
    lstDiatonic.ListItems.Add , , NotetoText(Note(4)) & "7", , 3
    lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(4) & "/" & "7" & "/" & "0"
    
    lstDiatonic.ListItems.Add , , NotetoText(Note(6)) & "dim", , 3
    lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(6) & "/" & "dim" & "/" & "0"
    
    lstDiatonic.ListItems.Add , , NotetoText(Note(6)) & "m7b5", , 3
    lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(6) & "/" & "m7b5" & "/" & "0"
    
ElseIf mMScale = "Minor" Then

    chkShape = False
    
    Note() = Split(CalcKey(Key, "Minor"), ",") 'Split Note
    
    lstDiatonic.ListItems.Clear

    lstDiatonic.ListItems.Add , , NotetoText(Note(0)) & "m", , 1
    lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(0) & "/" & "m" & "/" & "0"
    lstDiatonic.ListItems.Add , , NotetoText(Note(0)) & "m7", , 1
    lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(0) & "/" & "m7" & "/" & "0"
    lstDiatonic.ListItems.Add , , NotetoText(Note(0)) & "m6", , 1
    lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(0) & "/" & "m6" & "/" & "0"
    lstDiatonic.ListItems.Add , , NotetoText(Note(5) + 1) & "m7b5", , 1
    lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(5) + 1 & "/" & "m7b5" & "/" & "0"
    lstDiatonic.ListItems.Add , , NotetoText(Note(5)) & "M7", , 1
    lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(5) & "/" & "M7" & "/" & "0"
    lstDiatonic.ListItems.Add , , NotetoText(Note(2)) & "M7", , 1
    lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(2) & "/" & "M7" & "/" & "0"
    
    lstDiatonic.ListItems.Add , , NotetoText(Note(1)) & "m7b5", , 2
    lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(1) & "/" & "m7b5" & "/" & "0"
    lstDiatonic.ListItems.Add , , NotetoText(Note(1)) & "m7", , 2
    lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(1) & "/" & "m7" & "/" & "0"
    lstDiatonic.ListItems.Add , , NotetoText(Note(3)) & "m6", , 2
    lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(3) & "/" & "m6" & "/" & "0"
    lstDiatonic.ListItems.Add , , NotetoText(Note(3)) & "m7", , 2
    lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(3) & "/" & "m7" & "/" & "0"
    lstDiatonic.ListItems.Add , , NotetoText(Note(5)) & "M7", , 2
    lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(5) & "/" & "M7" & "/" & "0"
    
    lstDiatonic.ListItems.Add , , NotetoText(Note(4)) & "7", , 3
    lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(4) & "/" & "7" & "/" & "0"
    
    lstDiatonic.ListItems.Add , , NotetoText(Note(6) + 1) & "dim", , 3
    lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(6) + 1 & "/" & "dim" & "/" & "0"
    
    lstDiatonic.ListItems.Add , , NotetoText(Note(6) + 1) & "m7b5", , 3
    lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(6) + 1 & "/" & "m7b5" & "/" & "0"
    
    lstDiatonic.ListItems.Add , , NotetoText(Note(6)) & "7", , 3
    lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(6) & "/" & "7" & "/" & "0"
    
    lstDiatonic.ListItems.Add , , NotetoText(Note(4)) & "m7", , 3
    lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(4) & "/" & "m7" & "/" & "0"
End If
End Sub

Private Sub cmdCancel_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim i As Integer

Play() = Split(lstDiatonic.ListItems(lstDiatonic.SelectedItem.Index).ListSubItems(1).Text, "/") 'Split Note
Note() = Split(CalcNote(Play(0), Play(1), ""), ",") 'Split Note

For i = 0 To UBound(Note)
    Note(i) = Note(i) + (24)
Next i

Select Case cbInv.ListIndex
Case 1
    Note(0) = Note(0) + 12
    Note(1) = Note(1)
    Note(2) = Note(2)
Case 2
    Note(0) = Note(0) + 12
    Note(1) = Note(1) + 12
    Note(2) = Note(2)
Case 3
    Note(0) = Note(0) + 12
    Note(1) = Note(1) + 12
    Note(2) = Note(2) + 12
    Note(3) = Note(3)
End Select

bBass = True
PlayScale UBound(Note), True, 70, False
End Sub

Private Sub cmdCancel_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim i As Integer

Play() = Split(lstDiatonic.ListItems(lstDiatonic.SelectedItem.Index).ListSubItems(1).Text, "/") 'Split Note
Note() = Split(CalcNote(Play(0), Play(1), ""), ",") 'Split Note

For i = 0 To UBound(Note)
    Note(i) = Note(i) + ((frmPlayer.sldOctave.value + 1) * 12)
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

PlayScale UBound(Note), False, 70, False

End Sub
