VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCustomChord 
   BorderStyle     =   1  '단일 고정
   Caption         =   "사용자 코드 목록"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   855
   ClientWidth     =   7095
   Icon            =   "frmCustomChord.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   7095
   StartUpPosition =   2  '화면 가운데
   Begin VB.Frame Frame1 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   6615
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
         ItemData        =   "frmCustomChord.frx":014A
         Left            =   120
         List            =   "frmCustomChord.frx":015A
         TabIndex        =   5
         Top             =   3600
         Width           =   6375
      End
      Begin MSComDlg.CommonDialog cdFile 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ListView lstDiatonic 
         Height          =   3255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   5741
         View            =   2
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
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
   Begin VB.TextBox txtChord 
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
      Left            =   240
      TabIndex        =   3
      Top             =   4680
      Width           =   6615
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   4935
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   8705
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "사용자 코드 목록"
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
      Left            =   5760
      TabIndex        =   1
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdPreview 
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
      Left            =   4440
      TabIndex        =   0
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Menu mnuChord 
      Caption         =   "코드(&C)"
      Begin VB.Menu mnuNew 
         Caption         =   "새 코드 파일(&N)"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "코드 파일 열기(&O)"
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "코드 파일 저장(&S)"
      End
      Begin VB.Menu mnuSaveas 
         Caption         =   "다른 이름으로 코드파일 저장(&A).."
      End
   End
End
Attribute VB_Name = "frmCustomChord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPreview_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim i As Long

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

End Sub

Private Sub cmdPreview_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim i As Long

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

Private Sub cmdInsert_Click()
On Error Resume Next
If lstDiatonic.SelectedItem.Text <> "" Then
    If Mid$(lstDiatonic.SelectedItem.Text, 2, 1) = "#" Or Mid$(lstDiatonic.SelectedItem.Text, 2, 1) = "b" Then
        mdi_frmMain.ActiveForm.lstChord.ListItems.Add , , Mid$(lstDiatonic.SelectedItem.Text, 1, 2) & Mid$(lstDiatonic.SelectedItem.Text, 3, (Len(lstDiatonic.SelectedItem.Text)) - 2) & "/" & Mid$(lstDiatonic.SelectedItem.Text, 1, 2)
        mdi_frmMain.ActiveForm.lstChord.ListItems(mdi_frmMain.ActiveForm.lstChord.ListItems.Count).ListSubItems.Add.Text = TexttoNote(Mid$(lstDiatonic.SelectedItem.Text, 1, 2)) & "/" & (Mid$(lstDiatonic.SelectedItem.Text, 3, (Len(lstDiatonic.SelectedItem.Text)) - 2)) & "/" & "0" & "/" & cbInv.Text
    Else
        mdi_frmMain.ActiveForm.lstChord.ListItems.Add , , Mid$(lstDiatonic.SelectedItem.Text, 1, 1) & Mid$(lstDiatonic.SelectedItem.Text, 2, (Len(lstDiatonic.SelectedItem.Text)) - 1) & "/" & Mid$(lstDiatonic.SelectedItem.Text, 1, 1)
        mdi_frmMain.ActiveForm.lstChord.ListItems(mdi_frmMain.ActiveForm.lstChord.ListItems.Count).ListSubItems.Add.Text = TexttoNote(Mid$(lstDiatonic.SelectedItem.Text, 1, 1)) & "/" & (Mid$(lstDiatonic.SelectedItem.Text, 2, Len(lstDiatonic.SelectedItem.Text) - 1)) & "/" & "0" & "/" & cbInv.Text
    End If
End If
End Sub

Private Sub lstDiatonic_DblClick()
lstDiatonic.ListItems.Remove (lstDiatonic.SelectedItem.Index)
End Sub

Private Sub mnuNew_Click()

Dim i As Long

If Not lstDiatonic.ListItems.Count = 0 Then
    i = MsgBox("파일이 변경되었습니다" & vbCrLf & vbCrLf & "새로 작성하시겠습니까?", vbYesNoCancel)
End If

If i = 7 Then
    lstDiatonic.ListItems.Clear
End If

End Sub

Private Sub mnuOpen_Click()

On Error GoTo Pass

Dim strtmp

cdFile.DialogTitle = "코드파일 열기"
cdFile.Filter = "코드 파일(*.chd)|*.chd"
cdFile.ShowOpen

If cdFile.CancelError = True Then Exit Sub

If Dir(cdFile.FileName) <> "" Then

    lstDiatonic.ListItems.Clear
    Me.Caption = "사용자 코드 목록 - " & cdFile.FileTitle
    
    Open cdFile.FileName For Input As #1
    Do While Not EOF(1)
    Line Input #1, strtmp
    
    If Left$(strtmp, 1) = "0" Then
        lstDiatonic.ListItems.Add , , Mid$(strtmp, 2, Len(strtmp) - 1)
    ElseIf Left$(strtmp, 1) = "1" Then
        lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add , , Mid$(strtmp, 2, Len(strtmp) - 1)
    End If
    
    Loop
    Close #1
    
End If

Exit Sub

Pass:

End Sub

Private Sub mnuSave_Click()
On Error GoTo Save
Dim i, o As Long

If ChkSave = True Then

    If Dir(cdFile.FileName) <> "" Then
    
        o = MsgBox("파일이 이미 존재합니다, 덮어쓰겠습니까?", vbYesNo, "파일이 이미 존재함")
        
        If o = 7 Then
            Exit Sub
        End If
        
        Call SetAttr(cdFile.FileName, vbNormal)
        Kill cdFile.FileName
        GoTo Save
    End If
Else
    mnuSaveas_Click
End If

Exit Sub

Save:

    Open cdFile.FileName For Append Access Write As #2
    Print #2, "2" & NotetoText(Key)
    Close #2
    
For i = 1 To lstDiatonic.ListItems.Count
    Open cdFile.FileName For Append Access Write As #2
    Print #2, "0" & lstDiatonic.ListItems(i).Text
    Close #2
    Open cdFile.FileName For Append Access Write As #2
    Print #2, "1" & lstDiatonic.ListItems(i).ListSubItems.item(1).Text
    Close #2
Next i

End Sub

Private Sub mnuSaveas_Click()

On Error Resume Next

Dim i, o As Long

cdFile.Filter = "코드 파일(*.chd)|*.chd"
cdFile.DialogTitle = "다른 이름으로 코드파일 저장"
cdFile.ShowSave

If cdFile.CancelError = True Then Exit Sub

If Dir(cdFile.FileName) <> "" Then
o = MsgBox("파일이 이미 존재합니다, 덮어쓰겠습니까??", vbYesNo, "파일이 이미 존재함")
    If o = 7 Then
        Exit Sub
    End If
Call SetAttr(cdFile.FileName, vbNormal)
Kill cdFile.FileName
End If

If Not Dir(cdFile.FileName) <> "" Then
    Me.Caption = "Chord Progressive - " & cdFile.FileTitle
    
    Open cdFile.FileName For Append Access Write As #2
    Print #2, "2" & NotetoText(Key)
    Close #2
        
    For i = 1 To lstDiatonic.ListItems.Count
    
        Open cdFile.FileName For Append Access Write As #2
        Print #2, "0" & lstDiatonic.ListItems(i).Text
        Close #2
        
        Open cdFile.FileName For Append Access Write As #2
        Print #2, "1" & lstDiatonic.ListItems(i).ListSubItems.item(1).Text
        Close #2
    Next i
    
    ChkSave = True
    
End If
End Sub

Private Sub txtChord_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Dim Chd As String
If KeyCode = 13 Then
    If Mid$(txtChord.Text, 2, 1) = "#" Or Mid$(txtChord.Text, 2, 1) = "b" Then
        Chd = AllNotePos(Mid$(txtChord, 1, 2), Mid$(txtChord, 3, Len(txtChord) - 2))
        lstDiatonic.ListItems.Add , , Mid$(txtChord, 1, 2) & Mid$(txtChord, 3, Len(txtChord) - 2) & "/" & Mid$(txtChord, 1, 2)
        lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add , , TexttoNote(Mid$(txtChord, 1, 2)) & "/" & Mid$(txtChord, 3, Len(txtChord) - 2) & "/" & "0" & "/" & "0"
    Else
        Chd = AllNotePos(Mid$(txtChord, 1, 1), Mid$(txtChord, 2, Len(txtChord) - 1))
        lstDiatonic.ListItems.Add , , Mid$(txtChord, 1, 1) & Mid$(txtChord, 2, Len(txtChord) - 1) & "/" & Mid$(txtChord, 1, 1)
        lstDiatonic.ListItems(lstDiatonic.ListItems.Count).ListSubItems.Add , , TexttoNote(Mid$(txtChord, 1, 1)) & "/" & Mid$(txtChord, 2, Len(txtChord) - 1) & "/" & "0" & "/" & "0"
    End If
    txtChord.Text = ""
End If

End Sub
