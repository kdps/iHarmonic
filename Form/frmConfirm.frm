VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmConfirm 
   BorderStyle     =   1  '단일 고정
   Caption         =   "인증 도구"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   Icon            =   "frmConfirm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   6375
   StartUpPosition =   2  '화면 가운데
   Begin VB.TextBox txtCDKEY 
      Height          =   270
      Left            =   1440
      TabIndex        =   8
      Top             =   1800
      Width           =   4455
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   5655
      Begin VB.TextBox txtGen 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   4455
      End
      Begin VB.TextBox txtKey 
         Height          =   270
         Left            =   1080
         TabIndex        =   3
         Top             =   720
         Width           =   4455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "시디키 :"
         Height          =   180
         Left            =   280
         TabIndex        =   7
         Top             =   1240
         Width           =   660
      End
      Begin VB.Label labProductNum 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "이름 :"
         Height          =   180
         Left            =   480
         TabIndex        =   6
         Top             =   285
         Width           =   480
      End
      Begin VB.Label labCdKey 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "제품번호 :"
         Height          =   180
         Left            =   120
         TabIndex        =   5
         Top             =   765
         Width           =   840
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   4048
      ImageList       =   "Imglst"
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "인증"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   ""
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "인증(&C)"
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   2520
      Width           =   1455
   End
   Begin ComctlLib.ImageList Imglst 
      Left            =   120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmConfirm.frx":014A
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConfirm_Click()
Dim turnName As String
Dim i As Integer

'A-Z = 65-90 , 0-9 = 48-57

If txtCDKEY.Text = KeyGen(ComputerName, txtKey, 3) Then
    MsgBox "제품인증이 완료되었습니다"
    INIWrite "frmConfirm", "Key", txtKey, App.Path & "\language.ini"
    INIWrite "frmConfirm", "CDKEY", txtCDKEY, App.Path & "\language.ini"
    frmSplash.Show
    Unload Me
Else
    MsgBox "틀린 시디키입니다"
End If

End Sub

Private Sub Form_Load()
txtGen.Text = ComputerName
End Sub
