VERSION 5.00
Begin VB.Form frm 
   BorderStyle     =   1  '단일 고정
   Caption         =   "인증"
   ClientHeight    =   855
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4575
   Icon            =   "frmCertification.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   4575
   StartUpPosition =   2  '화면 가운데
   Begin VB.TextBox txtText 
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "다음 내용을 따라쓰세요 ""Cert"""
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3960
   End
End
Attribute VB_Name = "frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub txtText_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtText = "Cert" Then
        mdi_frmMain.Show
        Unload Me
    Else
        MsgBox "Err"
        End
    End If
End If
End Sub
