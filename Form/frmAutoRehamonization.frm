VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmAutoReharmonization 
   BorderStyle     =   5  '크기 조정 가능 도구 창
   Caption         =   "Auto Reharmonization"
   ClientHeight    =   8070
   ClientLeft      =   120
   ClientTop       =   390
   ClientWidth     =   12255
   Icon            =   "frmAutoRehamonization.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   12255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton front 
      Caption         =   "▶"
      Height          =   360
      Left            =   360
      TabIndex        =   5
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton back 
      Caption         =   "◀"
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox txtHistory 
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Text            =   "/"
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdAuto 
      Caption         =   "Auto Select"
      Height          =   375
      Left            =   6120
      TabIndex        =   2
      Top             =   7680
      Width           =   6135
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   375
      Left            =   -240
      TabIndex        =   1
      Top             =   7680
      Width           =   6375
   End
   Begin MSComctlLib.ListView lstReharmonization 
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   12938
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Chord"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Function"
         Object.Width           =   9596
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Favorite"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmAutoReharmonization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strPos As Integer

Private Function LoadPreset(dirFile As String)
On Error GoTo Pass
Dim strtmp As String
Dim BeforeFile As String

If Dir(App.Path & "\Preset\Reharmonization\" & dirFile & ".txt", vbDirectory) <> "" Then
    Open App.Path & "\Preset\Reharmonization\" & dirFile & ".txt" For Input As #1
    Do While Not EOF(1)
    Line Input #1, strtmp
    
    Dim strSplit() As String
    Dim intSel As Integer

    strSplit() = Split(strtmp, "/")
    
    If UBound(strSplit) = 0 Then
        Close #1
        chkShape = True
        If UBound(strSplit) > 4 Then
            If strSplit(5) <> "" Then
                LoadPreset strSplit(0) & "/" & strSplit(5)
                Exit Function
            End If
            Else
            LoadPreset strSplit(0)
            Exit Function
        End If
    End If
    
    If strtmp <> "" Then
        If UBound(strSplit) >= 3 Then
            If UBound(strSplit) = 6 Then
                chkShape = False
            Else
                chkShape = True
            End If
            If UBound(strSplit) = 5 Then
                If strSplit(5) = "" Then
                    lstReharmonization.ListItems.Add , , NotetoText(strSplit(0) + Key) & strSplit(1) & "/" & NotetoText(strSplit(0) + Key)
                    lstReharmonization.ListItems(lstReharmonization.ListItems.Count).ListSubItems.Add.Text = strSplit(3)
                    lstReharmonization.ListItems(lstReharmonization.ListItems.Count).ToolTipText = (strSplit(0)) & "/" & strSplit(1)
                Else
                    lstReharmonization.ListItems.Add , , NotetoText(strSplit(0) + Key) & strSplit(1) & "/" & NotetoText(strSplit(5) + Key)
                    lstReharmonization.ListItems(lstReharmonization.ListItems.Count).ListSubItems.Add.Text = strSplit(3)
                    lstReharmonization.ListItems(lstReharmonization.ListItems.Count).ToolTipText = (strSplit(0)) & "/" & strSplit(1) & "/" & NotetoText(strSplit(5) + Key)
                End If
            ElseIf UBound(strSplit) < 5 Then
                lstReharmonization.ListItems.Add , , NotetoText(strSplit(0) + Key) & strSplit(1) & "/" & NotetoText(strSplit(0) + Key)
                lstReharmonization.ListItems(lstReharmonization.ListItems.Count).ListSubItems.Add.Text = strSplit(3)
                lstReharmonization.ListItems(lstReharmonization.ListItems.Count).ToolTipText = (strSplit(0)) & "/" & strSplit(1)
            End If
            
            If UBound(strSplit) > 3 And strSplit(4) <> "" Then
                lstReharmonization.ListItems(lstReharmonization.ListItems.Count).ListSubItems.Add , , strSplit(4)
            End If

            intSel = lstReharmonization.ListItems.Count
            Select Case strSplit(2)
            Case 0
                lstReharmonization.ListItems(lstReharmonization.ListItems.Count).ForeColor = vbBlack
            Case 1
                lstReharmonization.ListItems(lstReharmonization.ListItems.Count).ForeColor = RGB(189, 189, 0)
            Case 2
                lstReharmonization.ListItems(lstReharmonization.ListItems.Count).ForeColor = RGB(0, 0, 255)
            Case 3
                lstReharmonization.ListItems(lstReharmonization.ListItems.Count).ForeColor = RGB(128, 128, 255)
            Case 4
                lstReharmonization.ListItems(lstReharmonization.ListItems.Count).ForeColor = RGB(0, 128, 128)
            Case 5
                lstReharmonization.ListItems(lstReharmonization.ListItems.Count).ForeColor = RGB(255, 0, 128)
            Case 6
                lstReharmonization.ListItems(lstReharmonization.ListItems.Count).ForeColor = RGB(255, 0, 255)
            Case 7
                lstReharmonization.ListItems(lstReharmonization.ListItems.Count).ForeColor = RGB(255, 0, 0)
            Case 8
                lstReharmonization.ListItems(lstReharmonization.ListItems.Count).ForeColor = RGB(163, 29, 75)
            Case 9
                lstReharmonization.ListItems(lstReharmonization.ListItems.Count).ForeColor = RGB(36, 137, 137)
            End Select
        End If
    End If
    
    Loop
    Close #1
    
    BeforeFile = dirFile
Else
    MsgBox dirFile & " file is corrupt or invalid file!" & BeforeFile, vbCritical, "Error"
End If
Exit Function
Pass:
End Function

Private Sub back_Click()
On Error Resume Next
Dim strSplit() As String
strSplit() = Split(txtHistory, "/")
If UBound(strSplit) > 1 And strPos > 1 Then
    lstReharmonization.ListItems.Clear
    LoadPreset strSplit(strPos - 1)
    If strSplit(strPos - 1) <> "" And strPos > 1 Then Me.Caption = strSplit(strPos - 1)
    strPos = strPos - 1
End If
End Sub

Private Sub cmdAuto_Click()
Dim strSplit() As String
Dim Comment As Integer
Dim i As Long
Comment = InputBox("How much would you like to create?", "Comment")
For i = 0 To Comment
    Randomize
    strSplit() = Split(lstReharmonization.ListItems(Int(Rnd * lstReharmonization.ListItems.Count) + 1).ToolTipText, "/")
    AddChord NotetoText(strSplit(0) + Key) & strSplit(1)
    lstReharmonization.ListItems.Clear
    chkShape = True
    LoadPreset NotetoText(strSplit(0) + 1) & strSplit(1)
Next i
End Sub

Private Sub cmdPreview_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
Dim i As Long

Play() = Split(lstReharmonization.ListItems(lstReharmonization.SelectedItem.Index).ToolTipText, "/") 'Split Note
Note() = Split(CalcNote(Play(0), Play(1), ""), ",") 'Split Note

For i = 0 To UBound(Note)
Note(i) = Note(i) + ((frmPlayer.sldOctave.value + 1) * 12)
Next i

PlayScale UBound(Note), True, 70, False

End Sub

Private Sub cmdPreview_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
Dim i As Long

Play() = Split(lstReharmonization.ListItems(lstReharmonization.SelectedItem.Index).ToolTipText, "/") 'Split Note
Note() = Split(CalcNote(Play(0), Play(1), ""), ",") 'Split Note

For i = 0 To UBound(Note)
Note(i) = Note(i) + ((frmPlayer.sldOctave.value + 1) * 12)
Next i

PlayScale UBound(Note), False, 70, False

End Sub

Private Sub Form_Load()
Dim strSplit() As String

strPos = 1

If Dir(App.Path & "\Preset\Reharmonization", vbDirectory) = "" Then
    MsgBox "Can't Find the Auto Reharmonization Preset Folder!", vbCritical, "Error"
Else
    If mdi_frmMain.ActiveForm.lstChord.ListItems.Count = 0 Then
        txtHistory = txtHistory & "Default" & "/"
        LoadPreset "Default"
    Else
        strSplit() = Split(mdi_frmMain.ActiveForm.lstChord.ListItems(mdi_frmMain.ActiveForm.lstChord.ListItems.Count).Text, "/")
        If Dir(App.Path & "\Preset\Reharmonization\" & strSplit(0) & ".txt", vbDirectory) <> "" Then
            txtHistory = txtHistory & strSplit(0) & "/"
            LoadPreset strSplit(0)
        Else
            txtHistory = txtHistory & "Default" & "/"
            LoadPreset "Default"
        End If
    End If
End If

End Sub

Private Sub Form_Resize()
On Error Resume Next
lstReharmonization.Width = Me.ScaleWidth
lstReharmonization.Height = Me.ScaleHeight - (cmdPreview.Height + front.Height)
cmdPreview.Top = lstReharmonization.Height + front.Height
cmdPreview.Width = Me.ScaleWidth / 2
cmdAuto.Top = lstReharmonization.Height + front.Height
cmdAuto.Width = Me.ScaleWidth / 2
cmdAuto.Left = cmdPreview.Width
End Sub

Private Sub front_Click()
On Error Resume Next
Dim strSplit() As String
strSplit() = Split(txtHistory, "/")
If UBound(strSplit) > 1 And strPos > 0 And UBound(strSplit) - 1 > strPos Then
    lstReharmonization.ListItems.Clear
    LoadPreset strSplit(strPos + 1)
    Me.Caption = strSplit(strPos + 1)
    strPos = strPos + 1
End If
End Sub

Private Sub lstReharmonization_DblClick()
Dim strSplit() As String

chkShape = True

strSplit() = Split(lstReharmonization.SelectedItem.ToolTipText, "/")
txtHistory = txtHistory & NotetoText(strSplit(0) + 1) & strSplit(1) & "/"
strSplit() = Split(txtHistory, "/")
Me.Caption = strSplit(UBound(strSplit) - 1)
strPos = UBound(strSplit)
strSplit() = Split(lstReharmonization.SelectedItem.ToolTipText, "/")
AddChord NotetoText(strSplit(0) + Key) & strSplit(1)
lstReharmonization.ListItems.Clear

LoadPreset NotetoText(strSplit(0) + 1) & strSplit(1)
End Sub

