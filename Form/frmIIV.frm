VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmIIV 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "파인더 : IIm7 V7 I"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Frame Frame2 
      Caption         =   "II-V"
      Height          =   615
      Left            =   240
      TabIndex        =   35
      Top             =   1320
      Width           =   2295
      Begin VB.TextBox txtII 
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtV 
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   240
         Width           =   735
      End
      Begin VB.Label labLine 
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   180
         Left            =   1200
         TabIndex        =   38
         Top             =   260
         Width           =   90
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "I"
      Height          =   615
      Left            =   240
      TabIndex        =   33
      Top             =   600
      Width           =   2295
      Begin VB.TextBox txtChord 
         Height          =   270
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   2055
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   1935
      Left            =   120
      TabIndex        =   32
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3413
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "I"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "IIm"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "V7"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox chkFlat 
      Caption         =   "#"
      Height          =   255
      Left            =   2760
      TabIndex        =   31
      Top             =   120
      Width           =   495
   End
   Begin VB.CheckBox chkMinor 
      Caption         =   "IIm7b5 (단조)"
      Height          =   255
      Left            =   2760
      TabIndex        =   30
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "생성(&G)"
      Height          =   735
      Left            =   2760
      TabIndex        =   29
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   27
      Left            =   4110
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2160
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   25
      Left            =   3855
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2160
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   22
      Left            =   3375
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2160
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   20
      Left            =   3135
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2160
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   18
      Left            =   2895
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2160
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   15
      Left            =   2415
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2160
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   13
      Left            =   2175
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2160
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   10
      Left            =   1710
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2160
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   8
      Left            =   1455
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2160
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   6
      Left            =   1215
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2160
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   3
      Left            =   750
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2160
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   1
      Left            =   495
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2160
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   28
      Left            =   4200
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   26
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   26
      Left            =   3960
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   27
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   24
      Left            =   3720
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   28
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   23
      Left            =   3480
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   17
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   21
      Left            =   3240
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   18
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   19
      Left            =   3000
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   19
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   17
      Left            =   2760
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   20
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   16
      Left            =   2520
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   21
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   14
      Left            =   2280
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   22
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   12
      Left            =   2040
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   23
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   11
      Left            =   1800
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   5
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   9
      Left            =   1560
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   6
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   7
      Left            =   1320
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   7
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   5
      Left            =   1080
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   8
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   4
      Left            =   840
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   9
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   2
      Left            =   600
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   10
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   0
      Left            =   360
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   11
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   29
      Left            =   120
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   39
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   2160
      Width           =   255
   End
End
Attribute VB_Name = "frmIIV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private selitem As String

Private Sub pKey_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

PlayNote Index + 13

End Sub

Private Sub pKey_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

StopNote Index + 13

End Sub

Private Sub cmdGenerate_Click()
On Error Resume Next
Dim Note() As String
Dim i As Long

If chkFlat.value = 1 Then
    chkShape = True
Else
    chkShape = False
End If

For i = 0 To 28
    If pKey(i).Tag = "1" Then
        pKey(i).BackColor = vbWhite
    Else
        pKey(i).BackColor = vbBlack
    End If
Next i

If Mid(txtChord, 2, 1) = "#" Or Mid(txtChord, 2, 1) = "b" Then
    If Mid(txtChord, 3, 1) = "m" Then
        If selitem = 1 Then
            Note() = Split(CalcKey((24) + TexttoNote(Mid(txtChord, 1, 2)), "Minor"), ",")
            For i = 0 To UBound(Note)
                pKey(Note(i) - 25).BackColor = vbRed
            Next i
            If chkMinor.value = 0 Then
                txtII = NotetoText(Note(1)) & "m7"
            Else
                txtII = NotetoText(Note(1)) & "m7b5"
            End If
            txtV = NotetoText(Note(4)) & "7"
        End If
    End If
End If

If Mid(txtII, 2, 1) = "#" Or Mid(txtII, 2, 1) = "b" Then
    If Mid(txtII, 3, 1) = "m" Then
        If selitem = 2 Then
            Note() = Split(CalcKey((24) + TexttoNote(Mid(txtII, 1, 2)) - 2, "Minor"), ",")
            For i = 0 To UBound(Note)
                pKey(Note(i) - 25).BackColor = vbRed
            Next i
            If chkMinor.value = 0 Then
                txtV = NotetoText(Note(4)) & "7"
            Else
                txtV = NotetoText(Note(4)) & "7"
            End If
            txtChord = NotetoText(Note(0))
        End If
    End If
End If

If Mid(txtV, 2, 1) = "#" Or Mid(txtV, 2, 1) = "b" Then
    If Mid(txtV, 3, 1) = "m" Then
        If selitem = 3 Then
            Note() = Split(CalcKey((24) + TexttoNote(Mid(txtV, 1, 2)) - 7, "Minor"), ",")
            For i = 0 To UBound(Note)
                pKey(Note(i) - 25).BackColor = vbRed
            Next i
            If chkMinor.value = 0 Then
                txtII = NotetoText(Note(1)) & "m7"
            Else
                txtII = NotetoText(Note(1)) & "m7"
            End If
            txtChord = NotetoText(Note(0))
        End If
    End If
End If

If Mid(txtChord, 2, 1) = "#" Or Mid(txtChord, 2, 1) = "b" Then
    If Mid(txtChord, 3, 1) = "M" Or Len(txtChord) = 2 Then
        If selitem = 1 Then
            Note() = Split(CalcKey((24) + TexttoNote(Mid(txtChord, 1, 2)), "Major"), ",")
            For i = 0 To UBound(Note)
                pKey(Note(i) - 25).BackColor = vbRed
            Next i
            txtII = NotetoText(Note(1)) & "m7"
            txtV = NotetoText(Note(4)) & "7"
        End If
    End If
End If

If Mid(txtII, 2, 1) = "#" Or Mid(txtII, 2, 1) = "b" Then
    If Mid(txtII, 3, 1) = "M" Or Len(txtII) = 2 Then
        If selitem = 2 Then
            Note() = Split(CalcKey((24) + TexttoNote(Mid(txtII, 1, 2)) - 2, "Major"), ",")
            For i = 0 To UBound(Note)
                pKey(Note(i) - 25).BackColor = vbRed
            Next i
            txtV = NotetoText(Note(4)) & "7"
            txtChord = NotetoText(Note(0))
        End If
    End If
End If

If Mid(txtV, 2, 1) = "#" Or Mid(txtV, 2, 1) = "b" Then
    If Mid(txtV, 3, 1) = "M" Or Len(txtV) = 2 Then
        If selitem = 3 Then
            Note() = Split(CalcKey((24) + TexttoNote(Mid(txtV, 1, 2)) - 7, "Major"), ",")
            For i = 0 To UBound(Note)
                pKey(Note(i) - 25).BackColor = vbRed
            Next i
            txtII = NotetoText(Note(1)) & "m7"
            txtChord = NotetoText(Note(0))
        End If
    End If
End If

If Not Mid(txtChord, 2, 1) = "#" Or Mid(txtChord, 2, 1) = "b" Then
    If Mid(txtChord, 2, 1) = "m" Then
        If selitem = 1 Then
            Note() = Split(CalcKey((24) + TexttoNote(Mid(txtChord, 1, 1)), "Minor"), ",")
            For i = 0 To UBound(Note)
                pKey(Note(i) - 25).BackColor = vbRed
            Next i
            If chkMinor.value = 0 Then
                txtII = NotetoText(Note(1)) & "m7"
            Else
                txtII = NotetoText(Note(1)) & "m7b5"
            End If
            txtV = NotetoText(Note(4)) & "7"
        End If
    End If
End If

If Not Mid(txtII, 2, 1) = "#" Or Mid(txtII, 2, 1) = "b" Then
    If Mid(txtII, 2, 1) = "m" Then
        If selitem = 2 Then
            Note() = Split(CalcKey((24) + TexttoNote(Mid(txtII, 1, 1)) - 2, "Minor"), ",")
            For i = 0 To UBound(Note)
                pKey(Note(i) - 25).BackColor = vbRed
            Next i
            If chkMinor.value = 0 Then
                txtV = NotetoText(Note(4)) & "7"
            Else
                txtV = NotetoText(Note(4)) & "7"
            End If
            txtChord = NotetoText(Note(0))
        End If
    End If
End If

If Not Mid(txtV, 2, 1) = "#" Or Mid(txtV, 2, 1) = "b" Then
    If Mid(txtV, 2, 1) = "m" Then
        If selitem = 3 Then
            Note() = Split(CalcKey((24) + TexttoNote(Mid(txtV, 1, 1)) - 7, "Minor"), ",")
            For i = 0 To UBound(Note)
                pKey(Note(i) - 25).BackColor = vbRed
            Next i
            If chkMinor.value = 0 Then
                txtII = NotetoText(Note(1)) & "m7"
            Else
                txtII = NotetoText(Note(1)) & "m7"
            End If
            txtChord = NotetoText(Note(0))
        End If
    End If
End If

If Not Mid(txtChord, 2, 1) = "#" Or Mid(txtChord, 2, 1) = "b" Then
    If Mid(txtChord, 2, 1) = "M" Or Len(txtChord) = 1 Then
        If selitem = 1 Then
            Note() = Split(CalcKey((24) + TexttoNote(Mid(txtChord, 1, 1)), "Major"), ",")
            For i = 0 To UBound(Note)
                pKey(Note(i) - 25).BackColor = vbRed
            Next i
            txtII = NotetoText(Note(1)) & "m7"
            txtV = NotetoText(Note(4)) & "7"
        End If
    End If
End If

If Not Mid(txtII, 2, 1) = "#" Or Mid(txtII, 2, 1) = "b" Then
    If Mid(txtII, 2, 1) = "M" Or Len(txtII) = 1 Then
        If selitem = 2 Then
            Note() = Split(CalcKey((24) + TexttoNote(Mid(txtII, 1, 1)) - 2, "Major"), ",")
            For i = 0 To UBound(Note)
                pKey(Note(i) - 25).BackColor = vbRed
            Next i
            txtV = NotetoText(Note(4)) & "7"
            txtChord = NotetoText(Note(0))
        End If
    End If
End If

If Not Mid(txtV, 2, 1) = "#" Or Mid(txtV, 2, 1) = "b" Then
    If Mid(txtV, 2, 1) = "M" Or Len(txtV) = 1 Or Mid(txtV, 2, 1) = "7" Then
        If selitem = 3 Then
            Note() = Split(CalcKey((24) + TexttoNote(Mid(txtV, 1, 1)) - 7, "Major"), ",")
            For i = 0 To UBound(Note)
                pKey(Note(i) - 25).BackColor = vbRed
            Next i
            txtII = NotetoText(Note(1)) & "7"
            txtChord = NotetoText(Note(0))
        End If
    End If
End If

Exit Sub
Err:
MsgBox "잘못된 코드이름", vbCritical, "II-V 제네레이터"
End Sub

Private Sub Form_Load()
selitem = 1
End Sub

Private Sub TabStrip1_Click()
selitem = TabStrip1.SelectedItem.Index
Select Case TabStrip1.SelectedItem.Index
Case 1
    txtChord.Locked = False
    txtChord.BackColor = &H80000005
    txtII.Locked = True
    txtII.BackColor = &HE0E0E0
    txtV.Locked = True
    txtV.BackColor = &HE0E0E0
Case 2
    txtChord.Locked = True
    txtChord.BackColor = &HE0E0E0
    txtII.Locked = False
    txtII.BackColor = &H80000005
    txtV.Locked = True
    txtV.BackColor = &HE0E0E0
Case 3
    txtChord.Locked = True
    txtChord.BackColor = &HE0E0E0
    txtII.Locked = True
    txtII.BackColor = &HE0E0E0
    txtV.Locked = False
    txtV.BackColor = &H80000005
End Select
End Sub
