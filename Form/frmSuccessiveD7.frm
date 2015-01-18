VERSION 5.00
Begin VB.Form frmSD7 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "파인더 : 지속된 도미넌트 7th"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton cmdFind 
      Caption         =   "찾기(&F)"
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox txtFind 
      Height          =   1095
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  '양방향
      TabIndex        =   4
      Top             =   480
      Width           =   4575
   End
   Begin VB.TextBox txtChord 
      Height          =   270
      Left            =   3600
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtD7 
      Height          =   270
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label labChord 
      AutoSize        =   -1  'True
      Caption         =   "찾는 코드 :"
      Height          =   180
      Left            =   2640
      TabIndex        =   2
      Top             =   150
      Width           =   900
   End
   Begin VB.Label labD7 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "도미넌트 7th : "
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   150
      Width           =   1200
   End
End
Attribute VB_Name = "frmSD7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFind_Click()
Dim i As Long
Dim z As Integer

txtFind = ""
z = 0

If Mid(txtChord, 2, 1) = "b" Then
    chkShape = False
Else
    chkShape = True
End If

If (Mid(txtD7, 2, 1) = "7" Or Mid(txtD7, 3, 1) = "7") And (Mid(txtChord, 2, 1) = "7" Or Mid(txtChord, 3, 1) = "7") Then
    If Mid(txtD7, 2, 1) = "#" Or Mid(txtD7, 2, 1) = "b" Then
        
        For i = 0 To 500
            z = z + 5
            txtFind = txtFind & NotetoText(TexttoNote(Mid(txtD7, 1, 2)) + z) & "7 "
            If txtChord = (NotetoText(TexttoNote(Mid(txtD7, 1, 2)) + z) & "7") Then
                Exit Sub
            End If
        Next
        
    Else
        For i = 0 To 500
            z = z + 5
            txtFind = txtFind & NotetoText(TexttoNote(Mid(txtD7, 1, 1)) + z) & "7 "
            If txtChord = (NotetoText(TexttoNote(Mid(txtD7, 1, 1)) + z) & "7") Then
                Exit Sub
            End If
        Next
    End If
End If

End Sub

