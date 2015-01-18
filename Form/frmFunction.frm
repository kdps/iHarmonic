VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmFunction 
   Caption         =   "기능분석기"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9480
   Icon            =   "frmFunction.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   9480
   StartUpPosition =   2  '화면 가운데
   Begin VB.CheckBox chkRelated 
      Caption         =   "Related II-V"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   4440
      Width           =   7935
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2990
      _Version        =   393217
      Style           =   7
      BorderStyle     =   1
      Appearance      =   0
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
   Begin VB.CommandButton cmdExpend 
      Caption         =   "펼치기(&E)"
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
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "시작(&S)"
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
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   2415
   End
End
Attribute VB_Name = "frmFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkRelated_Click()
If chkRelated.value = 0 Then
    Related = False
Else
    Related = True
End If
End Sub

Private Sub cmdExpend_Click()

On Error Resume Next

Dim z As Integer

For z = 0 To tv.Nodes.Count

    If tv.Nodes(z).Expanded = False Then
        If Not tv.Nodes(z).Child.Text = "알수없음" Then
            tv.Nodes(z).Expanded = True
            cmdExpend.Caption = "접기(&F)"
            GoTo ReStart
        End If
    ElseIf tv.Nodes(z).Expanded = True Then
        If Not tv.Nodes(z).Child.Text = "알수없음" Then
            tv.Nodes(z).Expanded = False
            cmdExpend.Caption = "펼치기(&F)"
            GoTo ReStart
        End If
    End If
ReStart:
Next z

End Sub

Private Sub cmdStart_Click()

On Error Resume Next

Dim ListItem() As String 'Array of Note Function
Dim Note() As String 'Array of Note
Dim z As Integer

tv.Nodes.Clear

For z = 1 To mdi_frmMain.ActiveForm.lstChord.ListItems.Count
    
    ListItem() = Split(mdi_frmMain.ActiveForm.lstChord.ListItems(z).ListSubItems(1).Text, "/")
    
    If Not ListItem(0) = "Style" Or Not ListItem(0) = "Comment" Then
        Note() = Split(CalcNote((12 * frmPlayer.sldOctave.value) + ListItem(0), ListItem(1), ListItem(2)), ",") 'Split Note
        
        If ListItem(0) - Key < 0 Then 'Overlab Delete
            ListItem(0) = ListItem(0) + 12
        End If
        
        CalcAll ListItem(0), ListItem(1)  'Calc Function

        tv.Nodes.Add , tvwChild, "General" & z, chdClassic
        tv.Nodes.Add "General" & z, tvwChild, , chdFunction
    End If
Next z

Dim i As Integer

For z = 1 To tv.Nodes.Count - 1

    If Not tv.Nodes(z).Child.Text = "알수없음" Then
        tv.Nodes(z).ForeColor = vbRed
        tv.Nodes(z).Child.Bold = True
        GoTo ReStart
    Else
        tv.Nodes(z).BackColor = vbGreen
    End If

ReStart:
Next z

MsgBox "완료!", vbExclamation, "기능 분석기"

End Sub

Private Sub Form_Load()
If Related = True Then
    chkRelated.value = 1
Else
    chkRelated.value = 0
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Not Me.WindowState = 1 Then
tv.Width = Me.ScaleWidth - 230
tv.Height = Me.ScaleHeight - (cmdExpend.Height + 300 + chkRelated.Height)
cmdExpend.Top = tv.Height + 200
cmdStart.Top = tv.Height + 200
cmdExpend.Width = (Me.ScaleWidth / 2) - 150
cmdStart.Width = (Me.ScaleWidth / 2) - 150
cmdStart.Left = cmdStart.Width + 150
chkRelated.Top = tv.Height + cmdExpend.Height + 250
End If
End Sub
