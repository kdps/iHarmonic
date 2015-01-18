VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0080FFFF&
   BorderStyle     =   0  '없음
   ClientHeight    =   5280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Timer tmrTimer 
      Left            =   0
      Top             =   0
   End
   Begin VB.Shape Shape1 
      Height          =   5280
      Left            =   0
      Top             =   0
      Width           =   8760
   End
   Begin VB.Label labProcess 
      Alignment       =   2  '가운데 맞춤
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
      Height          =   855
      Left            =   0
      TabIndex        =   5
      Top             =   3000
      Width           =   8775
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "iHarmonic"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   26.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   4440
      Width           =   8730
   End
   Begin VB.Label txtPath 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
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
      Left            =   480
      TabIndex        =   4
      Top             =   1200
      Width           =   60
   End
   Begin VB.Label txtProduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
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
      Left            =   480
      TabIndex        =   3
      Top             =   960
      Width           =   60
   End
   Begin VB.Label txtCopyright 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Width           =   60
   End
   Begin VB.Label txtVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
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
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   60
   End
   Begin VB.Image imgSplash 
      Height          =   5280
      Left            =   0
      Picture         =   "frmSplash.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8760
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
On Error Resume Next
labProcess.Caption = "Wave Initializing..."

Timer2 200
Unload Me
mdi_frmMain.Show
End Sub

Private Sub Form_Load()
Me.Width = imgSplash.Width
Me.Height = imgSplash.Height
txtVersion.Caption = "Version : " & App.Major & "." & App.Minor & "." & App.Revision
txtCopyright.Caption = "Copyright : " & App.LegalCopyright$
txtProduct.Caption = "Product Name : " & App.ProductName
End Sub

Private Sub Image1_DblClick()
Unload Me
mdi_frmMain.Show
Timer1.Enabled = False
End Sub

Private Sub tmrTimer_Timer()
tmrTimer.Enabled = False
End Sub
