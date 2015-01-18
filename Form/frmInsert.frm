VERSION 5.00
Begin VB.Form frmInsert 
   BorderStyle     =   1  '단일 고정
   Caption         =   "노트 삽입"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5535
   Icon            =   "frmInsert.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5535
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton cmdSplit 
      Caption         =   "나누기(&P)"
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "정지(&S)"
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   2880
      Width           =   1695
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      Max             =   24
      Min             =   -12
      TabIndex        =   4
      Top             =   2160
      Width           =   5295
   End
   Begin VB.HScrollBar HScroll 
      Height          =   255
      Left            =   120
      Max             =   15
      Min             =   -3
      TabIndex        =   2
      Top             =   1560
      Value           =   5
      Width           =   5295
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "재생(&P)"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   1695
   End
   Begin VB.ListBox lstNote 
      Height          =   1320
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label labPitch 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "0"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   5295
   End
   Begin VB.Label labTempo 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "4분음표"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   5295
   End
End
Attribute VB_Name = "frmInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()

frmStyle.Enabled = True

If Mid(lstNote.Text, 1, 1) = "s" Then
    frmStyle.lstNote(frmStyle.SSTab.Tab).ListItems.Add , , lstNote & "," & HScroll.Value
Else
    frmStyle.lstNote(frmStyle.SSTab.Tab).ListItems.Add , , lstNote & "," & HScroll1.Value & "," & "p"
End If

End Sub

Private Sub cmdSplit_Click()

    frmStyle.lstNote(frmStyle.SSTab.Tab).ListItems.Add , , "w" & "," & frmStyle.lstNote(frmStyle.SSTab.Tab).ListItems.Count

End Sub

Private Sub cmdStop_Click()

    frmStyle.lstNote(frmStyle.SSTab.Tab).ListItems.Add , , lstNote & "," & HScroll1.Value & "," & "s"

End Sub

Private Sub HScroll_Change()

If HScroll.Value = -3 Then
    labTempo = "64분음표"
ElseIf HScroll.Value = -2 Then
    labTempo = "점64분음표"
ElseIf HScroll.Value = -1 Then
    labTempo = "32분음표"
ElseIf HScroll.Value = 0 Then
    labTempo = "점32분음표"
ElseIf HScroll.Value = 1 Then
    labTempo = "16분음표"
ElseIf HScroll.Value = 2 Then
    labTempo = "점16분음표"
ElseIf HScroll.Value = 3 Then
    labTempo = "8분음표"
ElseIf HScroll.Value = 4 Then
    labTempo = "점8분음표"
ElseIf HScroll.Value = 5 Then
    labTempo = "4분음표"
ElseIf HScroll.Value = 6 Then
    labTempo = "점4분음표"
ElseIf HScroll.Value = 7 Then
    labTempo = "2분음표"
ElseIf HScroll.Value = 8 Then
    labTempo = "점2분음표"
ElseIf HScroll.Value = 9 Then
    labTempo = "온음표"
End If

End Sub

Private Function NotetoText(txtNote) As String 'Convert Velocity to Note Name
On Error Resume Next

Do
DoEvents
txtNote = txtNote - 12
Loop Until txtNote < 12

Do
DoEvents
txtNote = txtNote + 12
Loop Until txtNote > 0

If txtNote > 12 Then
    Do
    DoEvents
    txtNote = txtNote - 12
    Loop Until txtNote < 12
End If

If txtNote = "1" Then
    NotetoText = "C"
ElseIf txtNote = "2" Then
    NotetoText = "C#"
ElseIf txtNote = "3" Then
    NotetoText = "D"
ElseIf txtNote = "4" Then
    NotetoText = "D#"
ElseIf txtNote = "5" Then
    NotetoText = "E"
ElseIf txtNote = "6" Then
    NotetoText = "F"
ElseIf txtNote = "7" Then
    NotetoText = "F#"
ElseIf txtNote = "8" Then
    NotetoText = "G"
ElseIf txtNote = "9" Then
    NotetoText = "G#"
ElseIf txtNote = "10" Then
    NotetoText = "A"
ElseIf txtNote = "11" Then
    NotetoText = "A#"
ElseIf txtNote = "12" Then
    NotetoText = "B"
End If

End Function

Private Sub HScroll1_Change()

If HScroll1.Value = 13 Then
    labPitch = "옥타브"
ElseIf HScroll1.Value < 0 Then
    labPitch = "다운 옥타브 + " & Key + HScroll1.Value - 1
ElseIf HScroll1.Value > 12 Then
    labPitch = "옥타브 + " & Key + HScroll1.Value - 14
Else
    labPitch = Key + HScroll1.Value
End If

End Sub

Private Sub HScroll2_Change()

    labSplit = HScroll2.Value & "씩 나누기"
    
End Sub
