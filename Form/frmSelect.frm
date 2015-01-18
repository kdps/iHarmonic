VERSION 5.00
Begin VB.Form frmSelect 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  '없음
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton cmdCheck 
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
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public nSel As Integer
Dim nBtn As Integer

Public Function AddBtn(strName As String)
Dim i As Integer
Dim nHeight As String

nHeight = 0
nBtn = nBtn + 1
Load cmdCheck(nBtn)
cmdCheck(nBtn).Visible = True
cmdCheck(nBtn).Caption = strName
cmdCheck(nBtn).Top = ((cmdCheck(0).Height * nBtn) - cmdCheck(0).Height) + (150 * nBtn)

For i = 0 To cmdCheck.Count - 1
    nHeight = nHeight + (cmdCheck(i).Height + 100)
Next i

Me.Height = nHeight
End Function

Private Sub cmdCheck_Click(Index As Integer)
nSel = Index
Unload Me
End Sub

Private Sub Form_Load()
nBtn = 0
nSel = -1
End Sub

