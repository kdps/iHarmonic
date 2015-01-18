VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmDecendDim 
   Caption         =   "디센딩 디미니쉬드"
   ClientHeight    =   4950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6615
   Icon            =   "frmDecendDim.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   6615
   StartUpPosition =   2  '화면 가운데
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
      Left            =   5160
      TabIndex        =   5
      Top             =   4440
      Width           =   1335
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
      Left            =   3720
      TabIndex        =   4
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   360
      TabIndex        =   0
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
         ItemData        =   "frmDecendDim.frx":1272
         Left            =   120
         List            =   "frmDecendDim.frx":1282
         TabIndex        =   1
         Top             =   2880
         Width           =   5655
      End
      Begin MSComctlLib.ListView lstDiatonic 
         Height          =   2535
         Left            =   120
         TabIndex        =   2
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
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDecendDim.frx":1292
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDecendDim.frx":182C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDecendDim.frx":1DC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDecendDim.frx":2360
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   4095
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   7223
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "디센딩 디미니쉬드"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   ""
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
End
Attribute VB_Name = "frmDecendDim"
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

Private Sub Form_Load()

Dim Note() As String
Dim i As Integer

If mMScale = "Major" Then
    
    Note() = Split(CalcKey(Key, "Major"), ",") 'Split Note
    
    chkShape = False
    
    lstDiatonic.ListItems.Clear

    lstDiatonic.ListItems.Add , , NotetoText(Note(4) + 1) & "dim7"
    lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(4) + 1 & "/" & "dim7" & "/" & "0"
    lstDiatonic.ListItems.Add , , NotetoText(Note(4)) & "7"
    lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(4) & "/" & "7" & "/" & "0"

    lstDiatonic.ListItems.Add , , NotetoText(Note(1) + 1) & "dim7"
    lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(1) + 1 & "/" & "dim7" & "/" & "0"
    lstDiatonic.ListItems.Add , , NotetoText(Note(1)) & "m7"
    lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(1) & "/" & "m7" & "/" & "0"

    lstDiatonic.ListItems.Add , , NotetoText(Note(0) + 1) & "dim7"
    lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(1) + 1 & "/" & "dim7" & "/" & "0"
    lstDiatonic.ListItems.Add , , NotetoText(Note(0)) & "M7"
    lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(1) & "/" & "M7" & "/" & "0"

End If
End Sub

