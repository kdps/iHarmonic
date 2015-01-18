VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmPKC 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "파인더 : 나란한조 코드"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin MSComctlLib.ListView lstOrin 
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1508
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
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
      ItemData        =   "frmPKC.frx":0000
      Left            =   600
      List            =   "frmPKC.frx":0028
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin MSComctlLib.ListView lstPKC 
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1508
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label labRoot 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "근음:"
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
      TabIndex        =   1
      Top             =   150
      Width           =   360
   End
End
Attribute VB_Name = "frmPKC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbRoot_Click()
Dim Note() As String

Note() = Split(CalcKey(cbRoot.ListIndex + 1, "Minor"), ",") 'Split Note
lstPKC.ListItems.Clear
lstPKC.ListItems.Add , , NotetoText(Note(0)) & "m7"
lstPKC.ListItems.Add , , NotetoText(Note(1)) & "m7b5"
lstPKC.ListItems.Add , , NotetoText(Note(2)) & "M7"
lstPKC.ListItems.Add , , NotetoText(Note(3)) & "m7"
lstPKC.ListItems.Add , , NotetoText(Note(4)) & "7"
lstPKC.ListItems.Add , , NotetoText(Note(5)) & "M7"
lstPKC.ListItems.Add , , NotetoText(Note(6)) & "dim7"
Note() = Split(CalcKey(cbRoot.ListIndex + 1, "Major"), ",") 'Split Note
lstOrin.ListItems.Clear
lstOrin.ListItems.Add , , NotetoText(Note(0)) & "M7"
lstOrin.ListItems.Add , , NotetoText(Note(1)) & "m7"
lstOrin.ListItems.Add , , NotetoText(Note(2)) & "m7"
lstOrin.ListItems.Add , , NotetoText(Note(3)) & "M7"
lstOrin.ListItems.Add , , NotetoText(Note(4)) & "7"
lstOrin.ListItems.Add , , NotetoText(Note(5)) & "m7"
lstOrin.ListItems.Add , , NotetoText(Note(6)) & "m7b5"
End Sub
