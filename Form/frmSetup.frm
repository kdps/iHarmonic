VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOption 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Option"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5295
   Icon            =   "frmSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   5295
   Begin VB.PictureBox tabs 
      BorderStyle     =   0  '없음
      Height          =   3735
      Index           =   1
      Left            =   360
      ScaleHeight     =   3735
      ScaleWidth      =   4575
      TabIndex        =   20
      Top             =   600
      Visible         =   0   'False
      Width           =   4575
      Begin MSComDlg.CommonDialog cdColor 
         Left            =   4320
         Top             =   3600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog cdBackground 
         Left            =   4320
         Top             =   3480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         Caption         =   "프로그램"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3735
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   4575
         Begin VB.VScrollBar scrollTrans 
            Height          =   255
            Left            =   1330
            Max             =   130
            Min             =   10
            TabIndex        =   25
            Top             =   720
            Value           =   10
            Width           =   255
         End
         Begin VB.CommandButton cmdConnect 
            Caption         =   "악보 파일 연결하기"
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
            Left            =   2280
            Style           =   1  '그래픽
            TabIndex        =   24
            Top             =   3120
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox txtTrans 
            Enabled         =   0   'False
            Height          =   270
            Left            =   240
            TabIndex        =   23
            Top             =   720
            Width           =   1095
         End
         Begin VB.CommandButton cmdBackground 
            Caption         =   "배경화면 변경"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   370
            Left            =   240
            TabIndex        =   22
            Top             =   3120
            Width           =   1890
         End
         Begin ComctlLib.ImageList Imglst 
            Left            =   1920
            Top             =   1800
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   327682
            BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
               NumListImages   =   2
               BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "frmSetup.frx":014A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "frmSetup.frx":0324
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin VB.Label labTrans 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "투명도 :"
            Height          =   180
            Left            =   240
            TabIndex        =   26
            Top             =   360
            Width           =   660
         End
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      Caption         =   "편집기"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   4575
      Begin VB.PictureBox tabs 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   3255
         Index           =   0
         Left            =   120
         ScaleHeight     =   3255
         ScaleWidth      =   4335
         TabIndex        =   4
         Top             =   240
         Width           =   4335
         Begin VB.ListBox lstFont 
            Height          =   1140
            Left            =   120
            TabIndex        =   27
            Top             =   720
            Width           =   1575
         End
         Begin VB.Frame Frame2 
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            Caption         =   "색상"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1095
            Left            =   120
            TabIndex        =   13
            Top             =   2160
            Width           =   4095
            Begin VB.PictureBox txtPicture 
               Height          =   255
               Left            =   240
               ScaleHeight     =   195
               ScaleWidth      =   1635
               TabIndex        =   15
               Top             =   600
               Width           =   1695
               Begin VB.Label txtColor 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  '투명
                  Caption         =   "TEXT COLOR"
                  Height          =   180
                  Left            =   0
                  TabIndex        =   16
                  Top             =   0
                  Width           =   1170
               End
            End
            Begin VB.PictureBox editPicture 
               BackColor       =   &H00000000&
               Height          =   255
               Left            =   2160
               ScaleHeight     =   195
               ScaleWidth      =   1635
               TabIndex        =   14
               Top             =   600
               Width           =   1695
            End
            Begin VB.Label labEditor 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "편집기 :"
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
               Left            =   2160
               TabIndex        =   18
               Top             =   360
               Width           =   585
            End
            Begin VB.Label labFontColor 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "글꼴 :"
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
               Left            =   240
               TabIndex        =   17
               Top             =   360
               Width           =   420
            End
         End
         Begin VB.ListBox lstStyle 
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1230
            ItemData        =   "frmSetup.frx":04FE
            Left            =   1800
            List            =   "frmSetup.frx":050E
            TabIndex        =   9
            Top             =   720
            Width           =   1455
         End
         Begin VB.ListBox lstSize 
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1230
            ItemData        =   "frmSetup.frx":0534
            Left            =   3360
            List            =   "frmSetup.frx":0536
            TabIndex        =   8
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtSize 
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
            Left            =   3360
            TabIndex        =   7
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtStyle 
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
            Left            =   1800
            TabIndex        =   6
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox txtFont 
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
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label labSize 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "크기 :"
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
            Left            =   3360
            TabIndex        =   12
            Top             =   120
            Width           =   420
         End
         Begin VB.Label labStyle 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "글꼴 스타일 :"
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
            Left            =   1800
            TabIndex        =   11
            Top             =   120
            Width           =   975
         End
         Begin VB.Label labFont 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "글꼴 :"
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
            TabIndex        =   10
            Top             =   120
            Width           =   420
         End
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   4455
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   7858
      ImageList       =   "Imglst"
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "편집기"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "프로그램"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
            ImageIndex      =   2
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
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
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
      Left            =   2880
      TabIndex        =   2
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
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
      Left            =   1680
      TabIndex        =   1
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
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
      Left            =   4080
      TabIndex        =   0
      Top             =   4680
      Width           =   1095
   End
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub cmdApply_Click()

On Error Resume Next

If txtFont <> "" Then
    mdi_frmMain.ActiveForm.lstChord.Font = txtFont
End If

If txtTrans <> "" Then
    Trans mdi_frmMain.hWnd, 1, Val(txtTrans)
End If

If txtSize <> "" Then
    mdi_frmMain.ActiveForm.lstChord.Font.Size = txtSize
End If

Select Case txtStyle
Case "Nomal"
    mdi_frmMain.ActiveForm.lstChord.Font.Bold = False
    mdi_frmMain.ActiveForm.lstChord.Font.Italic = False
Case "Bold"
    mdi_frmMain.ActiveForm.lstChord.Font.Bold = True
    mdi_frmMain.ActiveForm.lstChord.Font.Italic = False
Case "Bold Italic"
    mdi_frmMain.ActiveForm.lstChord.Font.Bold = True
    mdi_frmMain.ActiveForm.lstChord.Font.Italic = True
Case "Italic"
    mdi_frmMain.ActiveForm.lstChord.Font.Bold = False
    mdi_frmMain.ActiveForm.lstChord.Font.Italic = True
End Select

mdi_frmMain.ActiveForm.lstChord.BackColor = editPicture.BackColor
mdi_frmMain.ActiveForm.lstChord.ForeColor = txtColor.ForeColor

MsgBox "적용 완료", vbExclamation, "설정"

End Sub

Private Sub cmdBackground_Click()
On Error Resume Next

cdBackground.DialogTitle = "Open Background"
cdBackground.Filter = "JPEG File(*.jpg)|*.jpg|Bitmap File(*.bmp)|*.bmp"
cdBackground.ShowOpen

If Dir(cdBackground.FileName) <> "" Then
    mdi_frmMain.Picture = LoadPicture(cdBackground.FileName)
    INIWrite "Form", "Background", cdBackground.FileName, App.Path & "\language.ini"
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdConnect_Click()

On Error Resume Next

SetDefExt "GED", "iHarmonic File", ".ged", App.Path & "\" & App.EXEName & ".exe"

End Sub

Private Sub cmdOK_Click()

On Error GoTo ErrOK

cmdApply_Click

If txtSize <> "" Then
    INIWrite "Font", "Size", txtSize, App.Path & "\language.ini"
End If

If txtFont <> "" Then
    INIWrite "Font", "Font", txtFont, App.Path & "\language.ini"
End If

If txtStyle <> "" Then
    INIWrite "Font", "Style", txtStyle, App.Path & "\language.ini"
End If

If txtTrans <> "" Then
    INIWrite "Font", "Trans", txtTrans, App.Path & "\language.ini"
End If

INIWrite "Font", "Color", txtColor.ForeColor, App.Path & "\language.ini"
INIWrite "Editor", "Color", editPicture.BackColor, App.Path & "\language.ini"
Unload Me
Exit Sub
ErrOK:
MsgBox "Description : " & Err.Description & vbCrLf & "Number: " & Err.Number, vbCritical, "Error"
End Sub

Private Sub Form_Load()

On Error Resume Next

For i = 8 To 72 Step 2
    lstSize.AddItem i
Next i
    
For i = 0 To Screen.FontCount - 1
    lstFont.AddItem Screen.Fonts(i)
Next i

Select Case INIRead("Font", "Style", App.Path & "\language.ini")
Case "Nomal"
    txtStyle.Text = "Nomal"
    lstStyle.Text = "Nomal"
Case "Bold"
    txtStyle.Text = "Bold"
    lstStyle.Text = "Bold"
Case "Bold Italic"
    txtStyle.Text = "Bold Italic"
    lstStyle.Text = "Bold Italic"
Case "Italic"
    txtStyle.Text = "Italic"
    lstStyle.Text = "Italic"
End Select

If INIRead("Font", "Size", App.Path & "\language.ini") <> "" Then
    txtSize.Text = INIRead("Font", "Size", App.Path & "\language.ini")
    lstSize.Text = INIRead("Font", "Size", App.Path & "\language.ini")
End If

If INIRead("Font", "Font", App.Path & "\language.ini") <> "" Then
    txtFont.Text = INIRead("Font", "Font", App.Path & "\language.ini")
    lstFont.Text = INIRead("Font", "Font", App.Path & "\language.ini")
End If

If INIRead("Font", "Color", App.Path & "\language.ini") <> "" Then
    txtColor.BackColor = INIRead("Font", "Color", App.Path & "\language.ini")
End If

If INIRead("Editor", "Color", App.Path & "\language.ini") <> "" Then
    editPicture.BackColor = INIRead("Editor", "Color", App.Path & "\language.ini")
    txtPicture.BackColor = INIRead("Editor", "Color", App.Path & "\language.ini")
End If

If INIRead("Font", "Trans", App.Path & "\language.ini") <> "" Then
    txtTrans = INIRead("Font", "Trans", App.Path & "\language.ini")
    scrollTrans.Value = INIRead("Font", "Trans", App.Path & "\language.ini")
End If

End Sub

Private Sub lstFont_Click()
txtFont.Text = lstFont.Text
End Sub

Private Sub lstSize_Click()
txtSize.Text = lstSize.Text
End Sub

Private Sub lstStyle_Click()
txtStyle.Text = lstStyle.Text
End Sub

Private Sub editPicture_Click()
cdColor.ShowColor
editPicture.BackColor = cdColor.Color
txtPicture.BackColor = cdColor.Color
End Sub

Private Sub scrollTrans_Change()
txtTrans.Text = scrollTrans.Value
End Sub

Private Sub TabStrip1_Click()
If TabStrip1.SelectedItem.Index = 1 Then
    tabs(0).Visible = True
    tabs(1).Visible = False
ElseIf TabStrip1.SelectedItem.Index = 2 Then
    tabs(0).Visible = False
    tabs(1).Visible = True
End If
End Sub

Private Sub txtColor_Click()
cdColor.ShowColor
txtColor.ForeColor = cdColor.Color
End Sub
