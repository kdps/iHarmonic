VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmStyle 
   BorderStyle     =   1  '단일 고정
   Caption         =   "사용자 코드 진행"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   5265
   Icon            =   "frmStyle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   5265
   StartUpPosition =   2  '화면 가운데
   Begin TabDlg.SSTab SSTab 
      Height          =   3950
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   6959
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   4
      TabHeight       =   520
      TabCaption(0)   =   "3 코드"
      TabPicture(0)   =   "frmStyle.frx":57E2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lstNote(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cdFile"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "4코드"
      TabPicture(1)   =   "frmStyle.frx":57FE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstNote(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "5코드"
      TabPicture(2)   =   "frmStyle.frx":581A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lstNote(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "6코드"
      TabPicture(3)   =   "frmStyle.frx":5836
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lstNote(3)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin MSComDlg.CommonDialog cdFile 
         Left            =   1920
         Top             =   1920
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ListView lstNote 
         Height          =   3075
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   5424
         View            =   2
         Arrange         =   2
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FlatScrollBar   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         PictureAlignment=   1
         _Version        =   393217
         SmallIcons      =   "Imglist"
         ForeColor       =   -2147483641
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "코드"
            Object.Width           =   15877
         EndProperty
      End
      Begin MSComctlLib.ListView lstNote 
         Height          =   3075
         Index           =   1
         Left            =   -74880
         TabIndex        =   3
         Top             =   120
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   5424
         View            =   2
         Arrange         =   2
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FlatScrollBar   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         PictureAlignment=   1
         _Version        =   393217
         SmallIcons      =   "Imglist"
         ForeColor       =   -2147483641
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "코드"
            Object.Width           =   15877
         EndProperty
      End
      Begin MSComctlLib.ListView lstNote 
         Height          =   3075
         Index           =   2
         Left            =   -74880
         TabIndex        =   4
         Top             =   120
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   5424
         View            =   2
         Arrange         =   2
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FlatScrollBar   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         PictureAlignment=   1
         _Version        =   393217
         SmallIcons      =   "Imglist"
         ForeColor       =   -2147483641
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "코드"
            Object.Width           =   15877
         EndProperty
      End
      Begin MSComctlLib.ListView lstNote 
         Height          =   3075
         Index           =   3
         Left            =   -74880
         TabIndex        =   5
         Top             =   120
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   5424
         View            =   2
         Arrange         =   2
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FlatScrollBar   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         PictureAlignment=   1
         _Version        =   393217
         SmallIcons      =   "Imglist"
         ForeColor       =   -2147483641
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "코드"
            Object.Width           =   15877
         EndProperty
      End
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "노트 삽입(&I)"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4200
      Width           =   5055
   End
   Begin VB.Menu mnuStyle 
      Caption         =   "스타일(&S)"
      Begin VB.Menu mnuOpen 
         Caption         =   "스타일 파일 열기(&O)"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "스타일 파일 저장(&S)"
      End
   End
End
Attribute VB_Name = "frmStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdInsert_Click()

frmInsert.Show
frmInsert.lstNote.Clear

If SSTab.Tab = 0 Then
    frmInsert.lstNote.AddItem "1st"
    frmInsert.lstNote.AddItem "3st"
    frmInsert.lstNote.AddItem "5st"
ElseIf SSTab.Tab = 1 Then
    frmInsert.lstNote.AddItem "1st"
    frmInsert.lstNote.AddItem "3st"
    frmInsert.lstNote.AddItem "5st"
    frmInsert.lstNote.AddItem "7st"
ElseIf SSTab.Tab = 2 Then
    frmInsert.lstNote.AddItem "1st"
    frmInsert.lstNote.AddItem "3st"
    frmInsert.lstNote.AddItem "5st"
    frmInsert.lstNote.AddItem "7st"
    frmInsert.lstNote.AddItem "8st"
ElseIf SSTab.Tab = 3 Then
    frmInsert.lstNote.AddItem "1st"
    frmInsert.lstNote.AddItem "3st"
    frmInsert.lstNote.AddItem "5st"
    frmInsert.lstNote.AddItem "7st"
    frmInsert.lstNote.AddItem "8st"
    frmInsert.lstNote.AddItem "9st"
End If

frmInsert.lstNote.AddItem "s"

End Sub

Private Sub lstNote_DblClick(Index As Integer)

lstNote(Index).ListItems.Remove (lstNote(Index).SelectedItem.Index)

End Sub

Private Sub lstNote_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

On Error Resume Next

Dim Backup As String

If KeyCode = 38 And lstNote(Index).SelectedItem.Index > 1 Then
    Backup = lstNote(Index).ListItems(lstNote(Index).SelectedItem.Index).Text
    lstNote(Index).ListItems(lstNote(Index).SelectedItem.Index).Text = lstNote(Index).ListItems(lstNote(Index).SelectedItem.Index - 1).Text
    lstNote(Index).ListItems(lstNote(Index).SelectedItem.Index - 1).Text = Backup
    Backup = lstNote(Index).ListItems(lstNote(Index).SelectedItem.Index).ListSubItems(1).Text
    lstNote(Index).ListItems(lstNote(Index).SelectedItem.Index).ListSubItems(1).Text = lstNote(Index).ListItems(lstNote(Index).SelectedItem.Index - 1).ListSubItems(1).Text
    lstNote(Index).ListItems(lstNote(Index).SelectedItem.Index - 1).ListSubItems(1).Text = Backup
ElseIf KeyCode = 40 Then
    Backup = lstNote(Index).ListItems(lstNote(Index).SelectedItem.Index + 1).Text
    lstNote(Index).ListItems(lstNote(Index).SelectedItem.Index + 1).Text = lstNote(Index).ListItems(lstNote(Index).SelectedItem.Index).Text
    lstNote(Index).ListItems(lstNote(Index).SelectedItem.Index).Text = Backup
    Backup = lstNote(Index).ListItems(lstNote(Index).SelectedItem.Index + 1).ListSubItems(1).Text
    lstNote(Index).ListItems(lstNote(Index).SelectedItem.Index + 1).ListSubItems(1).Text = lstNote(Index).ListItems(lstNote(Index).SelectedItem.Index).ListSubItems(1).Text
    lstNote(Index).ListItems(lstNote(Index).SelectedItem.Index).ListSubItems(1).Text = Backup
ElseIf KeyCode = 46 Then
    lstNote(Index).ListItems.Remove (lstNote(Index).SelectedItem.Index)
End If

End Sub


Private Sub mnuOpen_Click()

On Error Resume Next

Dim strtmp

cdFile.Filter = "스타일 파일(*.sty)|*.sty"
cdFile.flags = cdlOFNExplorer + cdlOFNAllowMultiselect
cdFile.DialogTitle = "스타일 파일 열기"
cdFile.ShowOpen

If cdFile.FileName <> "" Then

    lstChord.ListItems.Clear
    Me.Caption = "Chord Progressive - " & cdFile.FileTitle
    
    Open cdFile.FileName For Input As #1
    Do While Not EOF(1)
    Line Input #1, strtmp
        lstNote(SSTab.Tab).ListItems.Add , , strtmp
    Loop
    Close #1
    
    Saved = True
    
End If

End Sub

Private Sub mnuSave_Click()

On Error GoTo Save

Dim i, o As Long

cdFile.Filter = "스타일 파일(*.sty)|*.sty"
cdFile.flags = cdlOFNExplorer + cdlOFNAllowMultiselect
cdFile.DialogTitle = "스타일 파일 저장"
cdFile.ShowSave

If FileLen(cdFile.FileName) > 0 Then
o = MsgBox("파일이 존재합니다, 덮어쓰시겠습니까?", vbYesNo, "쓰기 오류")
    If o = 7 Then
        Exit Sub
    End If
Call SetAttr(cdFile.FileName, vbNormal)
Kill cdFile.FileName
GoTo Save
End If
Exit Sub

Save:
If cdFile.FileName <> "" Then
    For i = 1 To lstNote(SSTab.Tab).ListItems.Count
        Open cdFile.FileName For Append Access Write As #2
        Print #2, lstNote(SSTab.Tab).ListItems(i).Text
        Close #2
    Next i
    
    Saved = True
    
End If
End Sub
