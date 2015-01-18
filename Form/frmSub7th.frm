VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSub7th 
   BorderStyle     =   1  '단일 고정
   Caption         =   "섭스티튜드 도미넌트 7th"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   Icon            =   "frmsub7th.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6630
   StartUpPosition =   2  '화면 가운데
   Begin MSComctlLib.ImageList ImageList 
      Left            =   0
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsub7th.frx":1272
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsub7th.frx":180C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsub7th.frx":1DA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsub7th.frx":2340
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
         ItemData        =   "frmsub7th.frx":28DA
         Left            =   120
         List            =   "frmsub7th.frx":28EA
         TabIndex        =   4
         Top             =   2880
         Width           =   5655
      End
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
            Caption         =   "섭스티튜드 도미넌트 7th"
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
      TabIndex        =   1
      Top             =   4440
      Width           =   1215
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
      TabIndex        =   0
      Top             =   4440
      Width           =   1215
   End
End
Attribute VB_Name = "frmsub7th"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim i As Long

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
Dim i As Long

Play() = Split(lstDiatonic.ListItems(lstDiatonic.SelectedItem.Index).ListSubItems(1).Text, "/") 'Split Note
Note() = Split(CalcNote(Play(0), Play(1), ""), ",") 'Split Note

For i = 0 To UBound(Note)
Note(i) = Note(i) + ((frmPlayer.sldOctave.value + 1) * 12)
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

PlayScale UBound(Note), False, 70, False

End Sub

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

If mMScale = "Major" Then

    chkShape = False
    
    lstDiatonic.ListItems.Add , , NotetoText(Key + 1) & "7", , 1
    lstDiatonic.ListItems(1).ListSubItems.Add 1, , Key + 1 & "/" & "7" & "/" & "0"
    lstDiatonic.ListItems.Add , , NotetoText(Key) & "M7", , 2
    lstDiatonic.ListItems(2).ListSubItems.Add 1, , Key & "/" & "M7" & "/" & "0"
    
    lstDiatonic.ListItems.Add , , NotetoText(Key + 3) & "7", , 1
    lstDiatonic.ListItems(3).ListSubItems.Add 1, , Key + 3 & "/" & "7" & "/" & "0"
    lstDiatonic.ListItems.Add , , NotetoText(Key + 2) & "m7", , 3
    lstDiatonic.ListItems(4).ListSubItems.Add 1, , Key + 2 & "/" & "m7" & "/" & "0"
    
    lstDiatonic.ListItems.Add , , NotetoText(Key + 5) & "7", , 1
    lstDiatonic.ListItems(3).ListSubItems.Add 1, , Key + 5 & "/" & "7" & "/" & "0"
    lstDiatonic.ListItems.Add , , NotetoText(Key + 4) & "m7", , 3
    lstDiatonic.ListItems(4).ListSubItems.Add 1, , Key + 4 & "/" & "m7" & "/" & "0"
    
    lstDiatonic.ListItems.Add , , NotetoText(Key + 6) & "7", , 1
    lstDiatonic.ListItems(5).ListSubItems.Add 1, , Key + 6 & "/" & "7" & "/" & "0"
    lstDiatonic.ListItems.Add , , NotetoText(Key + 5) & "M7", , 3
    lstDiatonic.ListItems(6).ListSubItems.Add 1, , Key + 5 & "/" & "M7" & "/" & "0"
    
    lstDiatonic.ListItems.Add , , NotetoText(Key + 8) & "7", , 1
    lstDiatonic.ListItems(7).ListSubItems.Add 1, , Key + 8 & "/" & "7" & "/" & "0"
    lstDiatonic.ListItems.Add , , NotetoText(Key + 7) & "7", , 4
    lstDiatonic.ListItems(8).ListSubItems.Add 1, , Key + 7 & "/" & "7" & "/" & "0"
    
    lstDiatonic.ListItems.Add , , NotetoText(Key + 10) & "7", , 1
    lstDiatonic.ListItems(9).ListSubItems.Add 1, , Key + 10 & "/" & "7" & "/" & "0"
    lstDiatonic.ListItems.Add , , NotetoText(Key + 9) & "m7", , 2
    lstDiatonic.ListItems(10).ListSubItems.Add 1, , Key + 9 & "/" & "m7" & "/" & "0"
    
ElseIf mMScale = "Minor" Then

    chkShape = False
    
    lstDiatonic.ListItems.Add , , NotetoText(Key + 1) & "7", , 1
    lstDiatonic.ListItems(1).ListSubItems.Add 1, , Key + 1 & "/" & "7" & "/" & "0"
    lstDiatonic.ListItems.Add , , NotetoText(Key) & "m", , 2
    lstDiatonic.ListItems(2).ListSubItems.Add 1, , Key & "/" & "m" & "/" & "0"
    
    lstDiatonic.ListItems.Add , , NotetoText(Key + 6) & "7", , 1
    lstDiatonic.ListItems(1).ListSubItems.Add 1, , Key + 6 & "/" & "7" & "/" & "0"
    lstDiatonic.ListItems.Add , , NotetoText(Key + 5) & "m", , 3
    lstDiatonic.ListItems(2).ListSubItems.Add 1, , Key + 5 & "/" & "m" & "/" & "0"

    lstDiatonic.ListItems.Add , , NotetoText(Key + 8) & "7", , 1
    lstDiatonic.ListItems(1).ListSubItems.Add 1, , Key + 8 & "/" & "7" & "/" & "0"
    lstDiatonic.ListItems.Add , , NotetoText(Key + 7) & "7", , 4
    lstDiatonic.ListItems(2).ListSubItems.Add 1, , Key + 7 & "/" & "7" & "/" & "0"

    lstDiatonic.ListItems.Add , , NotetoText(Key + 9) & "7", , 1
    lstDiatonic.ListItems(1).ListSubItems.Add 1, , Key + 9 & "/" & "7" & "/" & "0"
    lstDiatonic.ListItems.Add , , NotetoText(Key + 8) & "M7", , 2
    lstDiatonic.ListItems(2).ListSubItems.Add 1, , Key + 8 & "/" & "M7" & "/" & "0"

End If
End Sub
