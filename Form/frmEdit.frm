VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmEdit 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8160
   Icon            =   "frmEdit.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4785
   ScaleWidth      =   8160
   Begin MSComctlLib.ListView lstChord 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3201
      View            =   1
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   4210752
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label labKey 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Key"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   165
      Width           =   330
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   285
      Picture         =   "frmEdit.frx":014A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1390
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   240
      Picture         =   "frmEdit.frx":01EC
      Top             =   120
      Width           =   45
   End
   Begin VB.Image Image9 
      Height          =   360
      Left            =   0
      Picture         =   "frmEdit.frx":034E
      Top             =   120
      Width           =   315
   End
   Begin VB.Label labchord 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1880
      TabIndex        =   2
      Top             =   165
      Width           =   75
   End
   Begin VB.Image Image7 
      Height          =   135
      Left            =   0
      Picture         =   "frmEdit.frx":0990
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5535
   End
   Begin VB.Image Image6 
      Height          =   360
      Left            =   1800
      Picture         =   "frmEdit.frx":09D6
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3375
   End
   Begin VB.Image Image5 
      Height          =   330
      Left            =   5160
      Picture         =   "frmEdit.frx":0AD8
      Top             =   135
      Width           =   165
   End
   Begin VB.Image Image3 
      Height          =   360
      Left            =   1680
      Picture         =   "frmEdit.frx":0E32
      Top             =   120
      Width           =   105
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      DrawMode        =   1  '검정
      X1              =   0
      X2              =   8160
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Image Image8 
      Height          =   615
      Left            =   0
      Picture         =   "frmEdit.frx":10B4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8055
   End
End
Attribute VB_Name = "frmedit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public RootKey As String
Public ScoreNum As Integer

Private Sub Form_Load()

Select Case INIRead("Font", "txtStyle", App.Path & "\language.ini")
Case "보통"
    lstChord.Font.Bold = False
    lstChord.Font.Italic = False
Case "굵게"
    lstChord.Font.Bold = True
    lstChord.Font.Italic = False
Case "굵은 기울임꼴"
    lstChord.Font.Bold = True
    lstChord.Font.Italic = True
Case "기울임꼴"
    lstChord.Font.Bold = False
    lstChord.Font.Italic = True
End Select

If INIRead("Font", "Size", App.Path & "\language.ini") <> "" Then
    lstChord.Font.Size = INIRead("Font", "Size", App.Path & "\language.ini")
End If

If INIRead("Font", "Font", App.Path & "\language.ini") <> "" Then
    lstChord.Font = INIRead("Font", "Font", App.Path & "\language.ini")
End If

If INIRead("Font", "Color", App.Path & "\language.ini") <> "" Then
    lstChord.ForeColor = INIRead("Font", "Color", App.Path & "\language.ini")
    labkey.ForeColor = lstChord.ForeColor
End If

If INIRead("Editor", "Color", App.Path & "\language.ini") <> "" Then
    lstChord.BackColor = INIRead("Editor", "Color", App.Path & "\language.ini")
End If

End Sub

Private Sub Form_Resize()
On Error Resume Next
If Not Me.WindowState = 1 Then
    lstChord.Top = 600
    lstChord.Width = Me.ScaleWidth
    lstChord.Height = Me.ScaleHeight - 600
    Line1.X2 = Me.ScaleWidth
    Image8.Width = Me.ScaleWidth
End If
End Sub

Private Sub labKey_Click()
Call frmKey.Show
End Sub

Private Sub lstChord_KeyDown(KeyCode As Integer, Shift As Integer)

On Error Resume Next

Dim Backup, Backup2 As String

If KeyCode = 37 And lstChord.SelectedItem.index > 1 Then
    Backup = lstChord.ListItems(lstChord.SelectedItem.index).Text
    Backup2 = lstChord.ListItems(lstChord.SelectedItem.index).ForeColor
    lstChord.ListItems(lstChord.SelectedItem.index).Text = lstChord.ListItems(lstChord.SelectedItem.index - 1).Text
    lstChord.ListItems(lstChord.SelectedItem.index).ForeColor = lstChord.ListItems(lstChord.SelectedItem.index - 1).ForeColor
    lstChord.ListItems(lstChord.SelectedItem.index - 1).Text = Backup
    lstChord.ListItems(lstChord.SelectedItem.index - 1).ForeColor = Backup2
    
    Backup = lstChord.ListItems(lstChord.SelectedItem.index).ListSubItems(1).Text
    lstChord.ListItems(lstChord.SelectedItem.index).ListSubItems(1).Text = lstChord.ListItems(lstChord.SelectedItem.index - 1).ListSubItems(1).Text
    lstChord.ListItems(lstChord.SelectedItem.index - 1).ListSubItems(1).Text = Backup
    
    Backup = lstChord.ListItems(lstChord.SelectedItem.index).SmallIcon
    lstChord.ListItems(lstChord.SelectedItem.index).SmallIcon = lstChord.ListItems(lstChord.SelectedItem.index - 1).SmallIcon
    lstChord.ListItems(lstChord.SelectedItem.index - 1).SmallIcon = Backup
ElseIf KeyCode = 39 And lstChord.SelectedItem.index < lstChord.ListItems.Count - 1 Then
    Backup = lstChord.ListItems(lstChord.SelectedItem.index + 1).Text
    Backup2 = lstChord.ListItems(lstChord.SelectedItem.index + 1).ForeColor
    lstChord.ListItems(lstChord.SelectedItem.index + 1).Text = lstChord.ListItems(lstChord.SelectedItem.index).Text
    lstChord.ListItems(lstChord.SelectedItem.index + 1).ForeColor = lstChord.ListItems(lstChord.SelectedItem.index).ForeColor
    lstChord.ListItems(lstChord.SelectedItem.index).Text = Backup
    lstChord.ListItems(lstChord.SelectedItem.index).ForeColor = Backup2
    
    Backup = lstChord.ListItems(lstChord.SelectedItem.index + 1).ListSubItems(1).Text
    lstChord.ListItems(lstChord.SelectedItem.index + 1).ListSubItems(1).Text = lstChord.ListItems(lstChord.SelectedItem.index).ListSubItems(1).Text
    lstChord.ListItems(lstChord.SelectedItem.index).ListSubItems(1).Text = Backup
    
    Backup = lstChord.ListItems(lstChord.SelectedItem.index + 1).SmallIcon
    lstChord.ListItems(lstChord.SelectedItem.index + 1).SmallIcon = lstChord.ListItems(lstChord.SelectedItem.index).SmallIcon
    lstChord.ListItems(lstChord.SelectedItem.index).SmallIcon = Backup
ElseIf KeyCode = 46 Then
    lstChord.ListItems.Remove (lstChord.SelectedItem.index)
End If

End Sub

Private Sub lstChord_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 2 Then
    PopupMenu mdi_frmMain.mnuEdit
End If
End Sub
