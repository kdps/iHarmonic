VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  '평면
   BackColor       =   &H80000005&
   Caption         =   "Chord Progressive - Untitled.chd"
   ClientHeight    =   6510
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9585
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   9585
   StartUpPosition =   2  '화면 가운데
   Begin VB.Menu mnuFile 
      Caption         =   "파일(&F)"
      WindowList      =   -1  'True
      Begin VB.Menu mnuNew 
         Caption         =   "새 악보(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "악보 열기(&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "악보 저장(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "다른이름으로 악보 저장(&A)"
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "인쇄(&P)"
      End
      Begin VB.Menu mnuSetupPrint 
         Caption         =   "인쇄 설정(&U)"
      End
      Begin VB.Menu mnuLine6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "프로그램 종료(&Q)"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdits 
      Caption         =   "편집(&E)"
      Begin VB.Menu mnuCopys 
         Caption         =   "복사(&C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuDels 
         Caption         =   "삭제(&D)"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu mnuChord 
      Caption         =   "코드(&C)"
      Begin VB.Menu mnuCustom 
         Caption         =   "사용자 정의 코드(&T)"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuChart 
         Caption         =   "사용자 정의 코드표(&R)"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuLine5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDiatonic 
         Caption         =   "다이아토닉 코드(&D)"
      End
      Begin VB.Menu mnuSecondary 
         Caption         =   "세컨더리 도미넌트(&S)"
      End
      Begin VB.Menu mnuSub7th 
         Caption         =   "서브도미넌트 7th(&V)"
      End
      Begin VB.Menu mnuSubminor 
         Caption         =   "서브마이너(&M)"
      End
      Begin VB.Menu mnuPassDim 
         Caption         =   "패싱 디미니쉬드(&P)"
      End
      Begin VB.Menu mnuModal 
         Caption         =   "모달 인터체인지(&M)"
      End
   End
   Begin VB.Menu mnuDevice 
      Caption         =   "장치(&D)"
      Begin VB.Menu mnuOutput 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnuTool 
      Caption         =   "도구(&T)"
      Begin VB.Menu mnuPiano 
         Caption         =   "피아노(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuStyle 
         Caption         =   "스타일(&S)"
      End
      Begin VB.Menu mnuLine4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFunction 
         Caption         =   "기능 분석 도구(&F)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuSetup 
      Caption         =   "설정(&S)"
      Begin VB.Menu mnuOption 
         Caption         =   "옵션(&O)"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuMidi 
         Caption         =   "미디(&M)"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuKey 
         Caption         =   "키(&K)"
         Shortcut        =   ^K
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "도움말(&H)"
      Begin VB.Menu mnuAbout 
         Caption         =   "Chord Progressive에 대하여..."
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuPopup1 
      Caption         =   ""
      Begin VB.Menu mnuCopy 
         Caption         =   "복사(&C)"
      End
      Begin VB.Menu mnuDel 
         Caption         =   "제거(&D)"
      End
      Begin VB.Menu mnuUp 
         Caption         =   "위로(&U)"
      End
      Begin VB.Menu mnuDown 
         Caption         =   "아래로(&D)"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "편집(&E)"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Terminate()
On Error Resume Next
midiOutClose hMidi
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
midiOutClose hMidi
End
End Sub

Private Sub lstChord_DblClick()
On Error Resume Next
lstChord.ListItems.Remove (lstChord.SelectedItem.Index)
End Sub


Private Sub lstChord_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

End Sub

Private Sub mnuChart_Click()
frmCustomChord.Show
End Sub

Private Sub mnuCopy_Click()
On Error Resume Next
lstChord.ListItems.Add , , lstChord.SelectedItem.Text, , 1
lstChord.ListItems(lstChord.ListItems.Count).ListSubItems.Add , , lstChord.SelectedItem.ListSubItems.Item(0).Text
End Sub

Private Sub mnuCopys_Click()
End Sub

Private Sub mnuCustom_Click()
frmCustom.Show
End Sub

Private Sub mnuDel_Click()
On Error Resume Next
lstChord.ListItems.Remove (lstChord.SelectedItem.Index)
End Sub

Private Sub mnuDels_Click()

End Sub

Private Sub mnuDiatonic_Click()

frmDiatonic.Show

End Sub

Private Sub mnuDown_Click()
On Error Resume Next
Dim Backup As String
If lstChord.SelectedItem.Index > lstChord.ListItems.Count - 1 Then
    Backup = lstChord.ListItems(lstChord.SelectedItem.Index + 1).Text
    lstChord.ListItems(lstChord.SelectedItem.Index + 1).Text = lstChord.ListItems(lstChord.SelectedItem.Index).Text
    lstChord.ListItems(lstChord.SelectedItem.Index).Text = Backup
    Backup = lstChord.ListItems(lstChord.SelectedItem.Index + 1).ListSubItems(1).Text
    lstChord.ListItems(lstChord.SelectedItem.Index + 1).ListSubItems(1).Text = lstChord.ListItems(lstChord.SelectedItem.Index).ListSubItems(1).Text
    lstChord.ListItems(lstChord.SelectedItem.Index).ListSubItems(1).Text = Backup
    
    Backup = lstChord.ListItems(lstChord.SelectedItem.Index + 1).SmallIcon
    lstChord.ListItems(lstChord.SelectedItem.Index + 1).SmallIcon = lstChord.ListItems(lstChord.SelectedItem.Index).SmallIcon
    lstChord.ListItems(lstChord.SelectedItem.Index).SmallIcon = Backup
End If
End Sub

Private Sub mnuFunction_Click()
frmFunction.Show
End Sub

Private Sub mnuKey_Click()
frmKey.Show
End Sub

Private Sub mnuMidi_Click()
frmMidi.Show
End Sub

Private Sub mnuModal_Click()
frmModal.Show
End Sub

Private Sub mnuOption_Click()
frmSetup.Show
End Sub

Private Sub mnuPassDim_Click()
frmPassDim.Show
End Sub

Private Sub mnuPrint_Click()

Dim i As Long

For i = 1 To lstChord.ListItems.Count
    Printer.Print lstChord.ListItems.Item(i)
    Printer.EndDoc
Next i

End Sub

Private Sub mnuQuit_Click()
midiOutClose hMidi
End
End Sub

Public Sub mnuSave_Click()

End Sub

Public Sub mnuSaveas_Click()


End Sub

Private Sub mnuSecondary_Click()
frmSecondary.Show
End Sub

Private Sub mnuSetupPrint_Click()
cdFile.ShowPrinter
End Sub

Private Sub mnuStyle_Click()
Dim i As Long
On Error Resume Next
i = InputBox("", "스타일 삽입")
If Not i = 0 Then
    lstChord.ListItems.Add , , "Style : " & i, , 1
    lstChord.ListItems(lstChord.ListItems.Count).ListSubItems.Add , , "Style/" & i
End If
End Sub

Private Sub mnuSub7th_Click()
frmsub7th.Show
End Sub

Private Sub mnuSubminor_Click()
frmSubMinor.Show
End Sub

Private Sub mnuUp_Click()
On Error Resume Next
Dim Backup As String
If lstChord.SelectedItem.Index > 1 Then
    Backup = lstChord.ListItems(lstChord.SelectedItem.Index).Text
    lstChord.ListItems(lstChord.SelectedItem.Index).Text = lstChord.ListItems(lstChord.SelectedItem.Index - 1).Text
    lstChord.ListItems(lstChord.SelectedItem.Index - 1).Text = Backup
    
    Backup = lstChord.ListItems(lstChord.SelectedItem.Index).ListSubItems(1).Text
    lstChord.ListItems(lstChord.SelectedItem.Index).ListSubItems(1).Text = lstChord.ListItems(lstChord.SelectedItem.Index - 1).ListSubItems(1).Text
    lstChord.ListItems(lstChord.SelectedItem.Index - 1).ListSubItems(1).Text = Backup
    
    Backup = lstChord.ListItems(lstChord.SelectedItem.Index).SmallIcon
    lstChord.ListItems(lstChord.SelectedItem.Index).SmallIcon = lstChord.ListItems(lstChord.SelectedItem.Index - 1).SmallIcon
    lstChord.ListItems(lstChord.SelectedItem.Index - 1).SmallIcon = Backup
End If
End Sub



Private Sub Timer1_Timer()

End Sub

Private Sub Timer2_Timer()
frmPiano.Left = frmMain.Left
frmPiano.Top = frmMain.Top + frmMain.Height
End Sub
