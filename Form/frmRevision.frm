VERSION 5.00
Begin VB.Form frmRevision 
   BorderStyle     =   1  '단일 고정
   Caption         =   "보정도구"
   ClientHeight    =   615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3375
   Icon            =   "frmRevision.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   615
   ScaleWidth      =   3375
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton cmdrevision 
      Caption         =   "보정(R)"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmRevision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub frmRevisionChord(Chord As String, i As Long)
On Error Resume Next
Dim Chd, txtRoot, txtKind, txtInvention As String
Dim Note() As String
Dim n, l As Long

Chord = Replace(Chord, " ", "")

ReDim Note(1)

Note() = Split(Chord, "/")

If Mid(Note(0), 2, 1) = "#" Or Mid(Note(0), 2, 1) = "b" Then
    txtRoot = Mid(Note(0), 1, 2)
    txtKind = Mid(Note(0), 3, Len(Note(0)) - 2)
    txtInvention = Note(1)
Else
    txtRoot = Mid(Note(0), 1, 1)
    txtKind = Mid(Note(0), 2, Len(Note(0)) - 1)
    txtInvention = Note(1)
End If

If Mid(txtKind, 2, 1) = "#" Or Mid(txtKind, 2, 1) = "b" Then
    strinput = Mid(txtKind, 3, Len(txtKind) - 2)
Else
    strinput = Mid(txtKind, 2, Len(txtKind) - 1)
End If

If Not InStr(strinput, "(") = 0 Then
    n = InStr(strinput, "(") + 1
    l = Len(Mid(strinput, n)) + 1
    RTention = Mid(strinput, n, Len(strinput) - (n))
    mdi_frmMain.ActiveForm.lstChord.ListItems(i).Text = txtRoot & txtKind & "/" & txtInvention
    mdi_frmMain.ActiveForm.lstChord.ListItems(i).ListSubItems(1).Text = TexttoNote(txtRoot) & "/" & Mid(strinput, 3, Len(strinput) - (2 + l)) & "/" & RTention & "/" & TexttoNote(txtInvention)
Else
    mdi_frmMain.ActiveForm.lstChord.ListItems(i).Text = txtRoot & txtKind & "/" & txtInvention
    mdi_frmMain.ActiveForm.lstChord.ListItems(i).ListSubItems(1).Text = TexttoNote(txtRoot) & "/" & txtKind & "/" & "0" & "/" & TexttoNote(txtInvention)
End If

End Sub

Private Sub cmdRevision_Click()
Dim i As Long
If Not mdi_frmMain.ActiveForm Is Nothing Then
    For i = 1 To mdi_frmMain.ActiveForm.lstChord.ListItems.Count - 1
        frmRevisionChord mdi_frmMain.ActiveForm.lstChord.ListItems(i).Text, i
    Next i
    MsgBox "보정 완료", vbExclamation, "보정 도구"
    Unload Me
End If
End Sub
