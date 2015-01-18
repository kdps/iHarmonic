VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmModal 
   BorderStyle     =   1  '단일 고정
   Caption         =   "모달 인터체인지"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6630
   Icon            =   "frmScaleChord.frx":0000
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
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScaleChord.frx":1272
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   700
      Left            =   360
      TabIndex        =   4
      Top             =   600
      Width           =   5895
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
         ItemData        =   "frmScaleChord.frx":180C
         Left            =   720
         List            =   "frmScaleChord.frx":1834
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cbType 
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
         ItemData        =   "frmScaleChord.frx":1861
         Left            =   3720
         List            =   "frmScaleChord.frx":1877
         TabIndex        =   5
         Top             =   240
         Width           =   1935
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
         TabIndex        =   8
         Top             =   315
         Width           =   420
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
         Left            =   3120
         TabIndex        =   7
         Top             =   315
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   5895
      Begin MSComctlLib.ListView lstDiatonic 
         Height          =   2295
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   4048
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
            Caption         =   "모달 인터체인지"
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
   Begin VB.CommandButton cmdPlay 
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
Attribute VB_Name = "frmModal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cbRoot_Click()
cbType.Enabled = True
End Sub

Private Sub cbType_Click()
If cbRoot.Text <> "" Then
    Dim Note() As String
    Dim i As Long
    
    Note() = Split(CalcKey(TexttoNote(cbRoot.Text), cbType.Text), ",") 'Split Note
    
    lstDiatonic.ListItems.Clear
    
    Select Case cbType.ListIndex
    Case 0
        For i = 0 To UBound(Note)
            If i = 0 Or i = 1 Or i = 4 Then
                lstDiatonic.ListItems.Add , , NotetoText(Note(i)) & "m7", , 1
                lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(i) & "/" & "m7" & "/" & "0"
            ElseIf i = 2 Or i = 6 Then
                lstDiatonic.ListItems.Add , , NotetoText(Note(i)) & "M7", , 1
                lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(i) & "/" & "M7" & "/" & "0"
            ElseIf i = 3 Then
                lstDiatonic.ListItems.Add , , NotetoText(Note(i)) & "7", , 1
                lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(i) & "/" & "7" & "/" & "0"
            ElseIf i = 5 Then
                lstDiatonic.ListItems.Add , , NotetoText(Note(i)) & "m7b5", , 1
                lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(i) & "/" & "m7b5" & "/" & "0"
            End If
        Next i
    Case 1
        For i = 0 To UBound(Note)
            If i = 0 Or i = 3 Then
                lstDiatonic.ListItems.Add , , NotetoText(Note(i)) & "m7", , 1
                lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(i) & "/" & "m7" & "/" & "0"
            ElseIf i = 1 Or i = 5 Then
                lstDiatonic.ListItems.Add , , NotetoText(Note(i)) & "M7", , 1
                lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(i) & "/" & "M7" & "/" & "0"
            ElseIf i = 2 Then
                lstDiatonic.ListItems.Add , , NotetoText(Note(i)) & "7", , 1
                lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(i) & "/" & "7" & "/" & "0"
            ElseIf i = 4 Then
                lstDiatonic.ListItems.Add , , NotetoText(Note(i)) & "m7b5", , 1
                lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(i) & "/" & "m7b5" & "/" & "0"
            ElseIf i = 6 Then
                lstDiatonic.ListItems.Add , , NotetoText(Note(i)) & "m7", , 1
                lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(i) & "/" & "m7" & "/" & "0"
            End If
        Next i
    Case 2
        For i = 0 To UBound(Note)
            If i = 0 Or i = 4 Then
                lstDiatonic.ListItems.Add , , NotetoText(Note(i)) & "M7", , 1
                lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(i) & "/" & "M7" & "/" & "0"
            ElseIf i = 1 Then
                lstDiatonic.ListItems.Add , , NotetoText(Note(i)) & "7", , 1
                lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(i) & "/" & "7" & "/" & "0"
            ElseIf i = 2 Or i = 5 Or i = 6 Then
                lstDiatonic.ListItems.Add , , NotetoText(Note(i)) & "m7", , 1
                lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(i) & "/" & "m7" & "/" & "0"
            ElseIf i = 3 Then
                lstDiatonic.ListItems.Add , , NotetoText(Note(i)) & "m7b5", , 1
                lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(i) & "/" & "m7b5" & "/" & "0"
            End If
        Next i
    Case 3
        For i = 0 To UBound(Note)
            If i = 0 Then
                lstDiatonic.ListItems.Add , , NotetoText(Note(i)) & "7", , 1
                lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(i) & "/" & "7" & "/" & "0"
            ElseIf i = 1 Or i = 4 Or i = 5 Then
                lstDiatonic.ListItems.Add , , NotetoText(Note(i)) & "m7", , 1
                lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(i) & "/" & "m7" & "/" & "0"
            ElseIf i = 2 Then
                lstDiatonic.ListItems.Add , , NotetoText(Note(i)) & "m7b5", , 1
                lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(i) & "/" & "m7b5" & "/" & "0"
            ElseIf i = 3 Or i = 6 Then
                lstDiatonic.ListItems.Add , , NotetoText(Note(i)) & "M7", , 1
                lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(i) & "/" & "M7" & "/" & "0"
            End If
        Next i
    Case 4
        For i = 0 To UBound(Note)
            If i = 0 Or i = 3 Or i = 4 Then
                lstDiatonic.ListItems.Add , , NotetoText(Note(i)) & "m7", , 1
                lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(i) & "/" & "m7" & "/" & "0"
            ElseIf i = 1 Then
                lstDiatonic.ListItems.Add , , NotetoText(Note(i)) & "m7b5", , 1
                lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(i) & "/" & "m7b5" & "/" & "0"
            ElseIf i = 2 Or i = 5 Then
                lstDiatonic.ListItems.Add , , NotetoText(Note(i)) & "M7", , 1
                lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(i) & "/" & "M7" & "/" & "0"
            ElseIf i = 6 Then
                lstDiatonic.ListItems.Add , , NotetoText(Note(i)) & "7", , 1
                lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(i) & "/" & "7" & "/" & "0"
            End If
        Next i
    Case 5
        For i = 0 To UBound(Note)
            If i = 0 Then
                lstDiatonic.ListItems.Add , , NotetoText(Note(i)) & "m7b5", , 1
                lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(i) & "/" & "m7b5" & "/" & "0"
            ElseIf i = 2 Or i = 3 Or i = 6 Then
                lstDiatonic.ListItems.Add , , NotetoText(Note(i)) & "m7", , 1
                lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(i) & "/" & "m7" & "/" & "0"
            ElseIf i = 1 Or i = 4 Then
                lstDiatonic.ListItems.Add , , NotetoText(Note(i)) & "M7", , 1
                lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(i) & "/" & "M7" & "/" & "0"
            ElseIf i = 5 Then
                lstDiatonic.ListItems.Add , , NotetoText(Note(i)) & "7", , 1
                lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add 1, , Note(i) & "/" & "7" & "/" & "0"
            End If
        Next i
    End Select
End If
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

Private Sub cmdPlay_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim i As Long

Play() = Split(lstDiatonic.ListItems(lstDiatonic.SelectedItem.Index).ListSubItems(1).Text, "/") 'Split Note
Note() = Split(CalcNote(Play(0), Play(1), ""), ",") 'Split Note

For i = 0 To UBound(Note)
Note(i) = Note(i) + (24)
Next i

bBass = True
PlayScale UBound(Note), True, 70, False

End Sub

Private Sub cmdPlay_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim i As Long

Play() = Split(lstDiatonic.ListItems(lstDiatonic.SelectedItem.Index).ListSubItems(1).Text, "/") 'Split Note
Note() = Split(CalcNote(Play(0), Play(1), ""), ",") 'Split Note

For i = 0 To UBound(Note)
Note(i) = Note(i) + ((frmPlayer.sldOctave.value + 1) * 12)
Next i

PlayScale UBound(Note), False, 70, False

End Sub
