VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmRehamonization 
   BorderStyle     =   1  '단일 고정
   Caption         =   "리하모니 제이션"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   BeginProperty Font 
      Name            =   "맑은 고딕"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRehamonization.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   5655
   StartUpPosition =   2  '화면 가운데
   Begin VB.PictureBox pBox2 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   360
      ScaleHeight     =   4335
      ScaleWidth      =   5055
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   5055
      Begin VB.Frame Frame3 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         Caption         =   "패싱 디미니쉬드"
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   0
         TabIndex        =   19
         Top             =   3120
         Width           =   4935
         Begin VB.CheckBox chkDiminished 
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            Caption         =   "패싱 디미니쉬드"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   360
            Width           =   3375
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         Caption         =   "토닉"
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   0
         TabIndex        =   17
         Top             =   2160
         Width           =   4935
         Begin VB.CheckBox chkTonics 
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            Caption         =   "#IVm7b5"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "서브도미넌트 마이너"
         Height          =   855
         Left            =   0
         TabIndex        =   15
         Top             =   1200
         Width           =   4935
         Begin VB.CheckBox chkSM 
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            Caption         =   "서브도미넌트 마이너"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   4335
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         Caption         =   "논-코드 톤"
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   0
         TabIndex        =   13
         Top             =   120
         Width           =   4935
         Begin VB.CheckBox chkAugmented 
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            Caption         =   "어그먼티드"
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   3975
         End
      End
   End
   Begin VB.PictureBox pBox1 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   360
      ScaleHeight     =   4455
      ScaleWidth      =   4935
      TabIndex        =   3
      Top             =   480
      Width           =   4935
      Begin VB.Frame frSecondary 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         Caption         =   "세컨더리 도미넌트"
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   0
         TabIndex        =   7
         Top             =   1440
         Width           =   4935
         Begin VB.CheckBox chkSecondary2 
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            Caption         =   "세컨더리 도미넌트 II"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   720
            Width           =   2415
         End
         Begin VB.CheckBox chkSecondaryIIVMinor 
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            Caption         =   "II-V (단조)"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   1440
            Width           =   3615
         End
         Begin VB.CheckBox chkSubSecondary 
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            Caption         =   "섭스티튜드 도미넌트"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   1800
            Width           =   2775
         End
         Begin VB.CheckBox chkSecondary 
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            Caption         =   "세컨더리 도미넌트"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   2535
         End
         Begin VB.CheckBox chkSecondaryIIV 
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            Caption         =   "II-V (장조)"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   1080
            Width           =   3375
         End
      End
      Begin VB.Frame fmDiatonic 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         Caption         =   "다이아토닉 코드"
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   0
         TabIndex        =   4
         Top             =   120
         Width           =   4935
         Begin VB.CheckBox chkDiatonic 
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            Caption         =   "지속된 코드 변경"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   360
            Width           =   2895
         End
         Begin VB.CheckBox chkTwoFive 
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            Caption         =   "II-V"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   720
            Width           =   1095
         End
      End
   End
   Begin ComctlLib.TabStrip TabStrip2 
      Height          =   4935
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   8705
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "리하모니제이션 I"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "리하모니제이션 II"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
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
   Begin VB.CommandButton cmdRehamonization 
      Caption         =   "적용(&A)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label labKey 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   5280
      Width           =   1980
   End
End
Attribute VB_Name = "frmRehamonization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strTention, LTention, LLTention As String

Dim i As Long
Dim z As Long
Dim n As Long
Dim q As Long
Dim nMsg As Long
Dim Note() As String
Dim RRtxt, Rtxt, Ltxt, LLTxt As String

Private Sub chkAugmented_Click()
If chkAugmented.Value = 1 Then
    If chkSecondary.Value = 1 Or chkSecondaryIIV.Value = 1 Or chkSecondaryIIVMinor.Value = 1 Or chkSubSecondary.Value = 1 Then
        chkAugmented.Value = 0
    End If
End If
End Sub

Private Sub chkDiminished_Click()
If chkDiminished.Value = 1 Then
    If chkSecondary.Value = 1 Or chkSecondaryIIV.Value = 1 Or chkSecondaryIIVMinor.Value = 1 Or chkSubSecondary.Value = 1 Then
        chkDiminished.Value = 0
    End If
End If
End Sub

Private Sub chkSecondary_Click()
If chkSecondary.Value = 1 Then
    If chkAugmented.Value = 1 Then
        chkSecondary.Value = 0
    End If
End If
End Sub

Private Sub chkSecondaryIIV_Click()
If chkSecondaryIIV.Value = 1 Then
    If chkSecondary.Value = 1 Then
        chkSecondaryIIV.Value = 0
    ElseIf chkSecondaryIIVMinor.Value = 1 Then
        chkSecondaryIIV.Value = 0
    End If
    
    If chkAugmented.Value = 1 Then
        chkSecondaryIIV.Value = 0
    End If
End If
End Sub

Private Sub chkSecondaryIIVMinor_Click()
If chkSecondaryIIVMinor.Value = 1 Then
    If chkSecondary.Value = 1 Then
        chkSecondaryIIVMinor.Value = 0
    ElseIf chkSecondaryIIV.Value = 1 And chkSecondaryIIVMinor.Value = 1 Then
        chkSecondaryIIVMinor.Value = 0
    End If
    
    If chkAugmented.Value = 1 Then
        chkSecondaryIIVMinor.Value = 0
    End If
End If
End Sub

Private Sub chkSubSecondary_Click()
If chkSubSecondary.Value = 1 Then
    If chkSecondary.Value = 1 Then
        chkSubSecondary.Value = 0
    ElseIf chkSecondaryIIV.Value = 1 Or chkSecondaryIIVMinor.Value = 1 Then
        chkSubSecondary.Value = 0
    End If
    If chkAugmented.Value = 1 Then
        chkSubSecondary.Value = 0
    End If
End If
End Sub

Private Sub cmdRehamonization_Click()

'On Error GoTo ErrMSG

If mdi_frmMain.ActiveForm.lstChord.ListItems.Count < 2 Then
    MsgBox "필요한 코드의 양보다 적습니다," & vbCrLf & vbCrLf & "코드를 더 작성하십시오.", vbCritical, "리하모니제이션 도구"
    Exit Sub
End If

If mdi_frmMain.ActiveForm.lstChord.ListItems.Count > 3 Then

    For i = 2 To mdi_frmMain.ActiveForm.lstChord.ListItems.Count
        
        If i < mdi_frmMain.ActiveForm.lstChord.ListItems.Count - 1 Then
            'RRtxt========================================================
            
            If InStr(mdi_frmMain.ActiveForm.lstChord.ListItems(i + 1).Text, "(") Then
                n = InStr(mdi_frmMain.ActiveForm.lstChord.ListItems(i + 1).Text, "(")
                q = InStr(mdi_frmMain.ActiveForm.lstChord.ListItems(i + 1).Text, "/")
                Rtxt = Mid$(mdi_frmMain.ActiveForm.lstChord.ListItems(i + 1).Text, q) & "/" & Mid$(mdi_frmMain.ActiveForm.lstChord.ListItems(i).Text, 1, n - 1)
                strTention = Mid$(mdi_frmMain.ActiveForm.lstChord.ListItems(i + 1).Text, n, Len(mdi_frmMain.ActiveForm.lstChord.ListItems(i).Text) - (n + 1))
            Else
                Rtxt = mdi_frmMain.ActiveForm.lstChord.ListItems(i).Text
            End If
        End If
        
        Note() = Split(Rtxt, "/")
        Rtxt = Replace(Note(0), " ", "")
        
        'Rtxt========================================================
        
        If InStr(mdi_frmMain.ActiveForm.lstChord.ListItems(i).Text, "(") Then
            n = InStr(mdi_frmMain.ActiveForm.lstChord.ListItems(i).Text, "(")
            q = InStr(mdi_frmMain.ActiveForm.lstChord.ListItems(i).Text, "/")
            Rtxt = Mid$(mdi_frmMain.ActiveForm.lstChord.ListItems(i).Text, q) & "/" & Mid$(mdi_frmMain.ActiveForm.lstChord.ListItems(i).Text, 1, n - 1)
            strTention = Mid$(mdi_frmMain.ActiveForm.lstChord.ListItems(i).Text, n, Len(mdi_frmMain.ActiveForm.lstChord.ListItems(i).Text) - (n + 1))
        Else
            Rtxt = mdi_frmMain.ActiveForm.lstChord.ListItems(i).Text
        End If
        
        Note() = Split(Rtxt, "/")
        Rtxt = Replace(Note(0), " ", "")
        
        'Ltxt========================================================
        
        If InStr(mdi_frmMain.ActiveForm.lstChord.ListItems(i - 1).Text, "(") Then
            n = InStr(mdi_frmMain.ActiveForm.lstChord.ListItems(i - 1).Text, "(")
            q = InStr(mdi_frmMain.ActiveForm.lstChord.ListItems(i - 1).Text, "/")
            Ltxt = Mid$(mdi_frmMain.ActiveForm.lstChord.ListItems(i - 1).Text, q) & "/" & Mid$(mdi_frmMain.ActiveForm.lstChord.ListItems(i - 1).Text, 1, n - 1)
            LTention = Mid$(mdi_frmMain.ActiveForm.lstChord.ListItems(i - 1).Text, n, Len(mdi_frmMain.ActiveForm.lstChord.ListItems(i - 1).Text) - (n + 1))
        Else
            Ltxt = mdi_frmMain.ActiveForm.lstChord.ListItems(i - 1).Text
        End If
        
        Note() = Split(Ltxt, "/")
        Ltxt = Replace(Note(0), " ", "")
        
        'LLtxt========================================================
        
        If mdi_frmMain.ActiveForm.lstChord.ListItems.Count > 5 And i > 2 Then

            If InStr(mdi_frmMain.ActiveForm.lstChord.ListItems(i - 2).Text, "(") Then
                n = InStr(mdi_frmMain.ActiveForm.lstChord.ListItems(i - 2).Text, "(")
                q = InStr(mdi_frmMain.ActiveForm.lstChord.ListItems(i - 2).Text, "/")
                LLTxt = Mid$(mdi_frmMain.ActiveForm.lstChord.ListItems(i - 2).Text, q) & "/" & Mid$(mdi_frmMain.ActiveForm.lstChord.ListItems(i - 2).Text, 1, n - 2)
                LLTention = Mid$(mdi_frmMain.ActiveForm.lstChord.ListItems(i - 2).Text, n, Len(mdi_frmMain.ActiveForm.lstChord.ListItems(i - 2).Text) - (n + 2))
            Else
                LLTxt = mdi_frmMain.ActiveForm.lstChord.ListItems(i - 2).Text
            End If

            Note() = Split(LLTxt, "/")
            LLTxt = Replace(Note(0), " ", "")
        End If
        
        If mdi_frmMain.ActiveForm.lstChord.ListItems.Count > 1 Then
        chkShape = True
            If chkAugmented.Value = 1 And mdi_frmMain.ActiveForm.lstChord.ListItems.Count > 1 Then 'C F
                If Rtxt = ChkNote(5) Or Rtxt = ChkNote(5) & "Maj" Or Rtxt = ChkNote(5) & "M7" Then
                    If Ltxt = ChkNote(0) Or Ltxt = ChkNote(0) & "Maj" Or Ltxt = ChkNote(0) & "M7" Then
                        frmSelect.Show
                        frmSelect.AddBtn NotetoText(Key) & "+"
                        frmSelect.AddBtn NotetoText(Key + 4) & "+"
                        frmSelect.AddBtn NotetoText(Key + 7) & "+"
                        frmSelect.AddBtn NotetoText(Key + 4)
                        frmSelect.AddBtn "Cancel"
                        Do
                        DoEvents
                        Loop Until Not frmSelect.nSel = -1
                        Select Case frmSelect.nSel
                            Case 1
                                ResetPos (i - 2), NotetoText(Key) & "+" & "/" & NotetoText(Key), TexttoNote(NotetoText(Key)) & "/" & "+" & "/" & "0" & "/" & TexttoNote(NotetoText(Key))
                            Case 2
                                ResetPos (i - 2), NotetoText(Key + 4) & "+" & "/" & NotetoText(Key + 4), TexttoNote(NotetoText(Key + 4)) & "/" & "+" & "/" & "0" & "/" & TexttoNote(NotetoText(Key + 4))
                            Case 3
                                ResetPos (i - 2), NotetoText(Key + 7) & "+" & "/" & NotetoText(Key + 7), TexttoNote(NotetoText(Key + 7)) & "/" & "+" & "/" & "0" & "/" & TexttoNote(NotetoText(Key + 7))
                            Case 4
                                ResetPos (i - 2), NotetoText(Key + 4) & "7" & "/" & NotetoText(Key + 4), TexttoNote(NotetoText(Key + 4)) & "/" & "7" & "/" & "0" & "/" & TexttoNote(NotetoText(Key + 4))
                        End Select
                    End If
                End If
            End If
        End If

        'Diatonic
        If chkDiatonic.Value = 1 Then
            chkShape = True
            If Rtxt = ChkNote(0) & "Maj" Or Rtxt = ChkNote(0) & "M7" Or Rtxt = ChkNote(0) Then 'Em Am(C)
                If Rtxt = Ltxt Then
                    Reset i - 1, NotetoText(Key + 4), "m7"
                    Reset i, NotetoText(Key + 9), "m7"
                End If
            ElseIf Rtxt = ChkNote(4) & "min" Or Rtxt = ChkNote(4) & "m7" Or Rtxt = ChkNote(4) & "m" Then 'Em Am(E)
                If Rtxt = Ltxt Then
                    Reset i - 1, NotetoText(Key + 4), "m7"
                    Reset i, NotetoText(Key + 9), "m7"
                End If
            ElseIf Rtxt = ChkNote(2) & "min" Or Rtxt = ChkNote(2) & "m7" Or Rtxt = ChkNote(2) & "m" Then 'Dm Am(Dm)
                If Rtxt = Ltxt Then
                    Reset i - 1, NotetoText(Key + 2), "m7"
                    Reset i, NotetoText(Key + 9), "m7"
                End If
            ElseIf Rtxt = ChkNote(9) & "min" Or Rtxt = ChkNote(9) & "m7" Or Rtxt = ChkNote(9) & "m" Then 'Dm Am(Am)
                If Rtxt = Ltxt Then
                    Reset i, NotetoText(Key + 9), "m7"
                End If
            End If
        End If
        
        'II-V
        If chkTwoFive.Value = 1 Then
            chkShape = True
            If Rtxt = ChkNote(7) & "7" Or Rtxt = ChkNote(7) & "Maj" Then 'F-G -> Dm-G
                If Ltxt = ChkNote(5) & "M7" Or Ltxt = ChkNote(5) & "Maj" Or Ltxt = ChkNote(5) Then
                    Reset i - 1, NotetoText(Key + 2), "m7"
                    mdi_frmMain.ActiveForm.lstChord.ListItems(i).ForeColor = vbBlue
                End If
            End If
        End If
        
        'Secondary Dominant II
        If chkSecondary2.Value = 1 Then
            If Not Ltxt = ChkNote(0) & "7" And Not Ltxt = ChkNote(0) & "M" And Not Ltxt = ChkNote(0) Then 'C7/F
                If Not Rtxt = ChkNote(0) & "7" And Not Rtxt = ChkNote(0) & "M" And Not Rtxt = ChkNote(0) Then
                    If Rtxt = ChkNote(5) & "M7" Or Rtxt = ChkNote(4) & "M" Or Rtxt = ChkNote(5) Then
                        If Not RRtxt = ChkNote(0) & "7" And Not RRtxt = ChkNote(0) & "M" And Not RRtxt = ChkNote(0) Then
                            ResetPos i - 2, NotetoText(Key + 0) & "7", TexttoNote(NotetoText(Key + 0)) & "/" & "7" & "/" & "0" & "/" & TexttoNote(NotetoText(Key + 0))
                            i = i + 2
                        End If
                    End If
                End If
            End If
        
            If Not Ltxt = ChkNote(2) & "7" And Not Ltxt = ChkNote(2) & "M" And Not Ltxt = ChkNote(2) Then 'D7/G
                If Not Rtxt = ChkNote(2) & "7" And Not Rtxt = ChkNote(2) & "M" And Not Rtxt = ChkNote(2) Then
                    If Rtxt = ChkNote(7) & "7" Or Rtxt = ChkNote(7) & "M" Or Rtxt = ChkNote(7) Then
                        If Not RRtxt = ChkNote(2) & "7" And Not RRtxt = ChkNote(2) & "M" And Not RRtxt = ChkNote(2) Then
                            ResetPos i - 2, NotetoText(Key + 2) & "7", TexttoNote(NotetoText(Key + 2)) & "/" & "7" & "/" & "0" & "/" & TexttoNote(NotetoText(Key + 2))
                            i = i + 2
                        End If
                    End If
                End If
            End If
            
            If Not Ltxt = ChkNote(4) & "7" And Not Ltxt = ChkNote(4) & "M" And Not Ltxt = ChkNote(4) Then 'E7/Am
                If Not Rtxt = ChkNote(4) & "7" And Not Rtxt = ChkNote(4) & "M" And Not Rtxt = ChkNote(4) Then
                    If Rtxt = ChkNote(9) & "m7" Or Rtxt = ChkNote(9) & "m" Or Rtxt = ChkNote(9) & "min" Then
                        If Not RRtxt = ChkNote(4) & "7" And Not RRtxt = ChkNote(4) & "M" And Not RRtxt = ChkNote(4) Then
                            ResetPos i - 2, NotetoText(Key + 4) & "7", TexttoNote(NotetoText(Key + 4)) & "/" & "7" & "/" & "0" & "/" & TexttoNote(NotetoText(Key + 4))
                            i = i + 2
                        End If
                    End If
                End If
            End If
            
            If Not Ltxt = ChkNote(5) & "7" And Not Ltxt = ChkNote(5) & "M" And Not Ltxt = ChkNote(5) Then 'F7/Bm7b5
                If Not Rtxt = ChkNote(9) & "7" And Not Rtxt = ChkNote(9) & "M" And Not Rtxt = ChkNote(9) Then
                    If Rtxt = ChkNote(11) & "m7b5" Or Rtxt = ChkNote(4) & "Dim" Or Rtxt = ChkNote(4) & "dim" Then
                        If Not RRtxt = ChkNote(5) & "7" And Not RRtxt = ChkNote(5) & "M" And Not RRtxt = ChkNote(5) Then
                            ResetPos i - 2, NotetoText(Key + 5) & "7", TexttoNote(NotetoText(Key + 5)) & "/" & "7" & "/" & "0" & "/" & TexttoNote(NotetoText(Key + 5))
                            i = i + 2
                        End If
                    End If
                End If
            End If
            
            If Not Ltxt = ChkNote(9) & "7" And Not Ltxt = ChkNote(9) & "M" And Not Ltxt = ChkNote(9) Then 'A7/Dm7
                If Not Rtxt = ChkNote(9) & "7" And Not Rtxt = ChkNote(9) & "M" And Not Rtxt = ChkNote(9) Then
                    If Rtxt = ChkNote(2) & "m7" Or Rtxt = ChkNote(2) & "m" Or Rtxt = ChkNote(2) & "min" Then
                        If Not RRtxt = ChkNote(9) & "7" And Not RRtxt = ChkNote(9) & "M" And Not RRtxt = ChkNote(9) Then
                            ResetPos i - 2, NotetoText(Key + 9) & "7", TexttoNote(NotetoText(Key + 9)) & "/" & "7" & "/" & "0" & "/" & TexttoNote(NotetoText(Key + 9))
                            i = i + 2
                        End If
                    End If
                End If
            End If
            
            
            If Not Ltxt = ChkNote(11) & "7" And Not Ltxt = ChkNote(11) & "M" And Not Ltxt = ChkNote(11) Then 'B7/Em
                If Not Rtxt = ChkNote(11) & "7" And Not Rtxt = ChkNote(11) & "M" And Not Rtxt = ChkNote(11) Then
                    If Rtxt = ChkNote(4) & "m7" Or Rtxt = ChkNote(4) & "m" Or Rtxt = ChkNote(4) & "min" Then
                        If Not RRtxt = ChkNote(11) & "7" And Not RRtxt = ChkNote(11) & "M" And Not RRtxt = ChkNote(11) Then
                            ResetPos i - 2, NotetoText(Key + 11) & "7", TexttoNote(NotetoText(Key + 11)) & "/" & "7" & "/" & "0" & "/" & TexttoNote(NotetoText(Key + 11))
                            i = i + 2
                        End If
                    End If
                End If
            End If
        End If
        
        'Secondary Dominant
        If chkSecondary.Value = 1 Then
            chkShape = True
            If Rtxt = ChkNote(5) & "M7" Or Rtxt = ChkNote(5) & "Maj" Or Rtxt = ChkNote(5) Then  'C7/F
                If Ltxt = ChkNote(0) & "M7" Or Ltxt = ChkNote(0) & "Maj" Or Ltxt = ChkNote(0) Then
                    Reset i - 1, NotetoText(Key + 0), "7"
                    mdi_frmMain.ActiveForm.lstChord.ListItems(i).ForeColor = vbGreen
                End If
            End If
            
            If Rtxt = ChkNote(7) & "7" Or Rtxt = ChkNote(7) & "Maj" Or Rtxt = ChkNote(7) Then  'D7/G
                If Ltxt = ChkNote(2) & "m7" Or Ltxt = ChkNote(2) & "m" Or Ltxt = ChkNote(2) & "min" Then
                    Reset i - 1, NotetoText(Key + 2), "7"
                    mdi_frmMain.ActiveForm.lstChord.ListItems(i).ForeColor = vbGreen
                End If
            End If
            
            If Rtxt = ChkNote(9) & "m7" Or Rtxt = ChkNote(9) & "m" Or Rtxt = ChkNote(9) & "min" Then 'E7/Am
                If Ltxt = ChkNote(4) & "m7" Or Ltxt = ChkNote(4) & "m" Or Ltxt = ChkNote(4) & "min" Then
                    Reset i - 1, NotetoText(Key + 4), "7"
                    mdi_frmMain.ActiveForm.lstChord.ListItems(i).ForeColor = vbGreen
                End If
            End If
            
            If Rtxt = ChkNote(11) & "m7b5" Or Rtxt = ChkNote(11) & "Dim" Or Rtxt = ChkNote(11) & "dim" Then 'F7/Bm7b5
                If Ltxt = ChkNote(5) & "M7" Or Ltxt = ChkNote(5) & "Maj" Or Ltxt = ChkNote(5) Then
                    Reset i - 1, NotetoText(Key + 5), "7"
                    mdi_frmMain.ActiveForm.lstChord.ListItems(i).ForeColor = vbGreen
                End If
            End If
            
            If Rtxt = ChkNote(2) & "m7" Or Rtxt = ChkNote(2) & "min" Or Rtxt = ChkNote(2) & "m" Then  'A7/Dm
                If Ltxt = ChkNote(9) & "m7" Or Ltxt = ChkNote(9) & "m" Or Ltxt = ChkNote(9) & "min" Then
                    Reset i - 1, NotetoText(Key + 9), "7"
                    mdi_frmMain.ActiveForm.lstChord.ListItems(i).ForeColor = vbGreen
                End If
            End If
                
            If Rtxt = ChkNote(11) & "m7b5" Or Rtxt = ChkNote(11) & "Dim" Or Rtxt = ChkNote(11) & "dim" Then 'B7/Em
                If Ltxt = ChkNote(4) & "m7" Or Ltxt = ChkNote(2) & "m" Or Ltxt = ChkNote(2) & "min" Then
                    Reset i - 1, NotetoText(Key + 11), "7"
                    mdi_frmMain.ActiveForm.lstChord.ListItems(i).ForeColor = vbGreen
                End If
            End If
            
        End If
        
        'Sub Secondary Dominant
        If chkSubSecondary.Value = 1 Then
            chkShape = False
            If Rtxt = ChkNote(5) & "M7" Or Rtxt = ChkNote(5) & "Maj" Or Rtxt = ChkNote(5) Then 'C7/F
                If Ltxt = ChkNote(0) & "M7" Or Ltxt = ChkNote(0) & "Maj" Or Ltxt = ChkNote(0) Then
                    Reset i - 1, NotetoText(Key + 6), "7"
                    mdi_frmMain.ActiveForm.lstChord.ListItems(i).ForeColor = vbBlue
                End If
            ElseIf Rtxt = ChkNote(7) & "7" Or Rtxt = ChkNote(7) & "Maj" Or Rtxt = ChkNote(7) Then 'D7/G
                If Ltxt = ChkNote(2) & "7" Or Ltxt = ChkNote(2) & "Maj" Or Ltxt = ChkNote(2) Then
                    Reset i - 1, NotetoText(Key + 8), "7"
                    mdi_frmMain.ActiveForm.lstChord.ListItems(i).ForeColor = vbBlue
                End If
            ElseIf Rtxt = ChkNote(9) & "m7" Or Rtxt = ChkNote(9) & "m" Or Rtxt = ChkNote(9) & "min" Then 'E7/Am
                If Ltxt = ChkNote(4) & "7" Or Ltxt = ChkNote(4) & "Maj" Or Ltxt = ChkNote(4) Then
                    Reset i - 1, NotetoText(Key + 10), "7"
                    mdi_frmMain.ActiveForm.lstChord.ListItems(i).ForeColor = vbBlue
                End If
            ElseIf Rtxt = ChkNote(11) & "m7b5" Or Rtxt = ChkNote(11) & "Dim" Or Rtxt = ChkNote(11) & "dim" Then 'F7/Bm7b5
                If Ltxt = ChkNote(5) & "7" Or Ltxt = ChkNote(5) & "Maj" Or Ltxt = ChkNote(5) Then
                    Reset i - 1, NotetoText(Key + 12), "7"
                    mdi_frmMain.ActiveForm.lstChord.ListItems(i).ForeColor = vbBlue
                End If
            ElseIf Rtxt = ChkNote(2) & "m7" Or Rtxt = ChkNote(2) & "min" Or Rtxt = ChkNote(2) & "m" Then   'A7/Dm
                If Ltxt = ChkNote(9) & "7" Or Ltxt = ChkNote(9) & "Maj" Or Ltxt = ChkNote(9) Then
                    Reset i - 1, NotetoText(Key + 3), "7"
                    mdi_frmMain.ActiveForm.lstChord.ListItems(i).ForeColor = vbBlue
                End If
            End If
        End If
        
        '#IVm7b5
        If chkTonics.Value = 1 Then
            If Rtxt = ChkNote(5) & "Maj" Or Rtxt = ChkNote(5) & "M7" Or Rtxt = ChkNote(5) Then   'C7/F
                If Ltxt = ChkNote(0) & "7" Or Ltxt = ChkNote(0) & "Maj" Or Ltxt = ChkNote(0) Then
                    Reset i - 1, NotetoText(Key + 4), "Dim7"
                    mdi_frmMain.ActiveForm.lstChord.ListItems(i).ForeColor = vbRed
                End If
            End If
        End If
        
        'Passing Diminisehd(Up)
        If chkDiminished.Value = 1 Then
            chkShape = True
            If Rtxt = ChkNote(5) & "Maj" Or Rtxt = ChkNote(5) & "M7" Or Rtxt = ChkNote(5) Then   'C7/F
                If Ltxt = ChkNote(0) & "7" Or Ltxt = ChkNote(0) & "Maj" Or Ltxt = ChkNote(0) Then
                    Reset i - 1, NotetoText(Key + 4), "Dim7"
                    mdi_frmMain.ActiveForm.lstChord.ListItems(i).ForeColor = vbRed
                End If
            ElseIf Rtxt = ChkNote(7) & "7" Or Rtxt = ChkNote(7) & "Maj" Or Rtxt = ChkNote(7) Then   'D7/G
                If Ltxt = ChkNote(2) & "7" Or Ltxt = ChkNote(2) & "Maj" Or Ltxt = ChkNote(2) Then
                    Reset i - 1, NotetoText(Key + 6), "Dim7"
                    mdi_frmMain.ActiveForm.lstChord.ListItems(i).ForeColor = vbRed
                End If
            ElseIf Rtxt = ChkNote(9) & "m7" Or Rtxt = ChkNote(9) & "m" Or Rtxt = ChkNote(9) & "min" Then  'E7/Am
                If Ltxt = ChkNote(4) & "7" Or Ltxt = ChkNote(4) & "Maj" Or Ltxt = ChkNote(4) Then
                    Reset i - 1, NotetoText(Key + 8), "Dim7"
                    mdi_frmMain.ActiveForm.lstChord.ListItems(i).ForeColor = vbRed
                End If
            ElseIf Rtxt = ChkNote(11) & "m7b5" Or Rtxt = ChkNote(11) & "Dim" Or Rtxt = ChkNote(11) & "dim" Then 'F7/Bm7b5
                If Ltxt = ChkNote(5) & "7" Or Ltxt = ChkNote(5) & "Maj" Or Ltxt = ChkNote(5) Then
                    Reset i - 1, NotetoText(Key + 10), "Dim7"
                    mdi_frmMain.ActiveForm.lstChord.ListItems(i).ForeColor = vbRed
                End If
            ElseIf Rtxt = ChkNote(0) & "Maj" Or Rtxt = ChkNote(0) & "M7" Or Rtxt = ChkNote(0) Then 'G7/C
                If Ltxt = ChkNote(7) & "7" Or Ltxt = ChkNote(7) & "Maj" Or Ltxt = ChkNote(7) Then
                    Reset i - 1, NotetoText(Key + 11), "Dim7"
                    mdi_frmMain.ActiveForm.lstChord.ListItems(i).ForeColor = vbRed
                End If
            ElseIf Rtxt = ChkNote(2) & "m7" Or Rtxt = ChkNote(2) & "min" Or Rtxt = ChkNote(2) & "m" Then   'A7/Dm
                If Ltxt = ChkNote(9) & "7" Or Ltxt = ChkNote(9) & "Maj" Or Ltxt = ChkNote(9) Then
                    Reset i - 1, NotetoText(Key + 1), "Dim7"
                    mdi_frmMain.ActiveForm.lstChord.ListItems(i).ForeColor = vbRed
                End If
            ElseIf Rtxt = ChkNote(4) & "m7" Or Rtxt = ChkNote(4) & "m" Or Rtxt = ChkNote(4) & "min" Then  'B7/Em
                If Ltxt = ChkNote(11) & "7" Or Ltxt = ChkNote(11) & "Maj" Or Ltxt = ChkNote(11) Then
                    Reset i - 1, NotetoText(Key + 3), "Dim7"
                    mdi_frmMain.ActiveForm.lstChord.ListItems(i).ForeColor = vbRed
                End If
            End If
        End If
        
        'Secondary Dominant II-V
        If mdi_frmMain.ActiveForm.lstChord.ListItems.Count > 3 Then
            chkShape = True
            If chkSecondaryIIV.Value = 1 And mdi_frmMain.ActiveForm.lstChord.ListItems.Count > 5 Then
                If Rtxt = ChkNote(5) & "M7" Or Rtxt = ChkNote(5) & "Maj" Or Rtxt = ChkNote(5) Then  'C7/F
                    If Ltxt = ChkNote(0) & "7" Or Ltxt = ChkNote(0) & "Maj" Or Ltxt = ChkNote(0) Then
                        If Not LLTxt = ChkNote(7) & "m7" And Not LLTxt = ChkNote(7) & "m" And Not LLTxt = ChkNote(7) & "min" Then
                            If Not LLTxt = ChkNote(7) & "m7b5" And Not LLTxt = ChkNote(7) & "Dim" And Not LLTxt = ChkNote(7) & "dim" Then
                                ResetPos (i - 3), NotetoText(Key + 7) & "m7", TexttoNote(NotetoText(Key + 7)) & "/" & "m7" & "/" & "0" & "/" & TexttoNote(NotetoText(Key + 7))
                            End If
                        End If
                    End If
                ElseIf Rtxt = ChkNote(7) & "7" Or Rtxt = ChkNote(7) & "Maj" Or Rtxt = ChkNote(7) Then  'D7/G
                    If Ltxt = ChkNote(2) & "7" Or Ltxt = ChkNote(2) & "Maj" Or Ltxt = ChkNote(2) Then
                        If Not LLTxt = ChkNote(9) & "m7" And Not LLTxt = ChkNote(9) & "m" And Not LLTxt = ChkNote(9) & "min" Then
                            If Not LLTxt = ChkNote(9) & "m7b5" And Not LLTxt = ChkNote(9) & "Dim" And Not LLTxt = ChkNote(9) & "dim" Then
                                ResetPos (i - 3), NotetoText(Key + 9) & "m7", TexttoNote(NotetoText(Key + 9)) & "/" & "m7" & "/" & "0" & "/" & TexttoNote(NotetoText(Key + 9))
                            End If
                        End If
                    End If
                ElseIf Rtxt = ChkNote(9) & "m7" Or Rtxt = ChkNote(9) & "m" Or Rtxt = ChkNote(9) & "min" Then 'E7/Am
                    If Ltxt = ChkNote(4) & "7" Or Ltxt = ChkNote(4) & "Maj" Or Ltxt = ChkNote(4) Then
                        If Not LLTxt = ChkNote(11) & "m7" And Not LLTxt = ChkNote(11) & "m" And Not LLTxt = ChkNote(11) & "min" Then
                            If Not LLTxt = ChkNote(11) & "m7b5" And Not LLTxt = ChkNote(11) & "Dim" And Not LLTxt = ChkNote(11) & "dim" Then
                                ResetPos (i - 3), NotetoText(Key + 11) & "m7", TexttoNote(NotetoText(Key + 11)) & "/" & "m7" & "/" & "0" & "/" & TexttoNote(NotetoText(Key + 11))
                            End If
                        End If
                    End If
                ElseIf Rtxt = ChkNote(11) & "m7b5" Or Rtxt = ChkNote(11) & "Dim" Or Rtxt = ChkNote(11) & "dim" Then 'F7/Bm7b5
                    If Ltxt = ChkNote(5) & "7" Or Ltxt = ChkNote(5) & "Maj" Or Ltxt = ChkNote(5) Then
                        If Not LLTxt = ChkNote(0) & "m7" And Not LLTxt = ChkNote(0) & "m" And Not LLTxt = ChkNote(0) & "min" Then
                            If Not LLTxt = ChkNote(0) & "m7b5" And Not LLTxt = ChkNote(0) & "Dim" And Not LLTxt = ChkNote(0) & "dim" Then
                                ResetPos (i - 3), NotetoText(Key) & "m7", TexttoNote(NotetoText(Key)) & "/" & "m7" & "/" & "0" & "/" & TexttoNote(NotetoText(Key))
                            End If
                        End If
                    End If
                ElseIf Rtxt = ChkNote(2) & "m7" Or Rtxt = ChkNote(2) & "min" Or Rtxt = ChkNote(2) & "m" Then  'A7/Dm
                    If Ltxt = ChkNote(9) & "7" Or Ltxt = ChkNote(9) & "Maj" Or Ltxt = ChkNote(9) Then
                        If Not LLTxt = ChkNote(4) & "m7" And Not LLTxt = ChkNote(4) & "m" And Not LLTxt = ChkNote(4) & "min" Then
                            If Not LLTxt = ChkNote(4) & "m7b5" And Not LLTxt = ChkNote(4) & "Dim" And Not LLTxt = ChkNote(4) & "dim" Then
                                ResetPos (i - 3), NotetoText(Key + 4) & "m7", TexttoNote(NotetoText(Key + 4)) & "/" & "m7" & "/" & "0" & "/" & TexttoNote(NotetoText(Key + 4))
                            End If
                        End If
                    End If
                ElseIf Rtxt = ChkNote(4) & "m7" Or Rtxt = ChkNote(4) & "min" Or Rtxt = ChkNote(4) & "m" Then 'B7/Em
                    If Ltxt = ChkNote(11) & "7" Or Ltxt = ChkNote(11) & "Maj" Or Ltxt = ChkNote(11) Then
                        If Not LLTxt = ChkNote(6) & "m7" And Not LLTxt = ChkNote(6) & "m" And Not LLTxt = ChkNote(6) & "min" Then
                            If Not LLTxt = ChkNote(6) & "m7b5" And Not LLTxt = ChkNote(6) & "Dim" And Not LLTxt = ChkNote(6) & "dim" Then
                                ResetPos (i - 3), NotetoText(Key + 6) & "m7", TexttoNote(NotetoText(Key + 6)) & "/" & "m7" & "/" & "0" & "/" & TexttoNote(NotetoText(Key + 6))
                            End If
                        End If
                    End If
                End If
            End If
        End If

        'Sub Dominant Minor
        If mdi_frmMain.ActiveForm.lstChord.ListItems.Count > 1 Then
        chkShape = True
            If chkSM.Value = 1 And mdi_frmMain.ActiveForm.lstChord.ListItems.Count > 1 Then 'F
                If Ltxt = ChkNote(5) Or Rtxt = ChkNote(5) & "Maj" Or Rtxt = ChkNote(5) & "M7" Then
                    frmSelect.Show
                    frmSelect.AddBtn NotetoText(Key + 5) & "m6"
                    frmSelect.AddBtn NotetoText(Key + 5) & "m7"
                    frmSelect.AddBtn NotetoText(Key + 8) & "6"
                    frmSelect.AddBtn NotetoText(Key + 8) & "M7"
                    frmSelect.AddBtn NotetoText(Key + 1) & "M7"
                    frmSelect.AddBtn NotetoText(Key + 2) & "m7b5"
                    frmSelect.AddBtn NotetoText(Key + 10) & "7"
                    frmSelect.AddBtn "Cancel"
                    Do
                    DoEvents
                    Loop Until Not frmSelect.nSel = -1
                    Select Case frmSelect.nSel
                        Case 1
                            ResetPos (i - 2), NotetoText(Key + 5) & "m6" & "/" & NotetoText(Key + 5), TexttoNote(NotetoText(Key + 5)) & "/" & "m6" & "/" & "0" & "/" & TexttoNote(NotetoText(Key) + 5)
                        Case 2
                            ResetPos (i - 2), NotetoText(Key + 5) & "m7" & "/" & NotetoText(Key + 5), TexttoNote(NotetoText(Key + 5)) & "/" & "m7" & "/" & "0" & "/" & TexttoNote(NotetoText(Key) + 5)
                        Case 3
                            ResetPos (i - 2), NotetoText(Key + 8) & "6" & "/" & NotetoText(Key + 8), TexttoNote(NotetoText(Key + 8)) & "/" & "6" & "/" & "0" & "/" & TexttoNote(NotetoText(Key) + 8)
                        Case 4
                            ResetPos (i - 2), NotetoText(Key + 8) & "M7" & "/" & NotetoText(Key + 8), TexttoNote(NotetoText(Key + 8)) & "/" & "M7" & "/" & "0" & "/" & TexttoNote(NotetoText(Key) + 8)
                        Case 5
                            ResetPos (i - 2), NotetoText(Key + 1) & "M7" & "/" & NotetoText(Key + 1), TexttoNote(NotetoText(Key + 1)) & "/" & "M7" & "/" & "0" & "/" & TexttoNote(NotetoText(Key) + 1)
                        Case 6
                            ResetPos (i - 2), NotetoText(Key + 2) & "m7b5" & "/" & NotetoText(Key + 2), TexttoNote(NotetoText(Key + 2)) & "/" & "m7b5" & "/" & "0" & "/" & TexttoNote(NotetoText(Key) + 2)
                        Case 7
                            ResetPos (i - 2), NotetoText(Key + 10) & "7" & "/" & NotetoText(Key + 10), TexttoNote(NotetoText(Key + 10)) & "/" & "7" & "/" & "0" & "/" & TexttoNote(NotetoText(Key) + 10)
                    End Select
                End If
            End If
        End If
        
        'Secondary Dominant II-V Minor
        If mdi_frmMain.ActiveForm.lstChord.ListItems.Count > 3 Then
        chkShape = True
            If chkSecondaryIIVMinor.Value = 1 And mdi_frmMain.ActiveForm.lstChord.ListItems.Count > 5 Then
                If Rtxt = ChkNote(9) & "m7" Or Rtxt = ChkNote(9) & "m" Or Rtxt = ChkNote(9) & "min" Then 'E7/Am
                    If Ltxt = ChkNote(4) & "7" Or Ltxt = ChkNote(4) & "Maj" Or Ltxt = ChkNote(4) Then
                        If Not LLTxt = ChkNote(11) & "m7b5" And Not LLTxt = ChkNote(11) & "Dim" And Not LLTxt = ChkNote(11) & "dim" Then
                            If Not LLTxt = ChkNote(11) & "m7" And Not LLTxt = ChkNote(11) & "m" And Not LLTxt = ChkNote(11) & "min" Then
                                ResetPos (i - 3), NotetoText(Key + 11) & "m7b5" & "/" & NotetoText(Key + 11), TexttoNote(NotetoText(Key + 11)) & "/" & "m7b5" & "/" & "0" & "/" & TexttoNote(NotetoText(Key + 11))
                            End If
                        ElseIf LLTxt = ChkNote(11) & "m7" Or LLTxt = ChkNote(11) & "m" Or LLTxt = ChkNote(11) & "min" Then
                            Reset i - 3, NotetoText(Key + 11), "m7b5"
                        End If
                    End If
                ElseIf Rtxt = ChkNote(2) & "m7" Or Rtxt = ChkNote(2) & "min" Or Rtxt = ChkNote(2) & "m" Then  'A7/Dm
                    If Ltxt = ChkNote(9) & "7" Or Ltxt = ChkNote(9) & "Maj" Or Ltxt = ChkNote(9) Then
                        If Not LLTxt = ChkNote(4) & "m7b5" And Not LLTxt = ChkNote(4) & "Dim" And Not LLTxt = ChkNote(4) & "dim" Then
                            If Not LLTxt = ChkNote(4) & "m7" And Not LLTxt = ChkNote(4) & "m" And Not LLTxt = ChkNote(4) & "min" Then
                                ResetPos (i - 3), NotetoText(Key + 4) & "m7b5" & "/" & NotetoText(Key + 4), TexttoNote(NotetoText(Key + 4)) & "/" & "m7b5" & "/" & "0" & "/" & TexttoNote(NotetoText(Key + 4))
                            End If
                        ElseIf LLTxt = ChkNote(4) & "m7" Or LLTxt = ChkNote(4) & "m" Or LLTxt = ChkNote(4) & "min" Then
                            Reset i - 3, NotetoText(Key + 4), "m7b5"
                        End If
                    End If
                ElseIf Rtxt = ChkNote(4) & "m7" Or Rtxt = ChkNote(4) & "min" Or Rtxt = ChkNote(4) & "m" Then 'B7/Em
                    If Ltxt = ChkNote(11) & "7" Or Ltxt = ChkNote(11) & "Maj" Or Ltxt = ChkNote(11) Then
                        If Not LLTxt = ChkNote(6) & "m7b5" And Not LLTxt = ChkNote(6) & "Dim" And Not LLTxt = ChkNote(6) & "dim" Then
                            If Not LLTxt = ChkNote(6) & "m7" And Not LLTxt = ChkNote(6) & "m" And Not LLTxt = ChkNote(6) & "min" Then
                                ResetPos (i - 3), NotetoText(Key + 6) & "m7b5" & "/" & NotetoText(Key + 6), TexttoNote(NotetoText(Key + 6)) & "/" & "m7b5" & "/" & "0" & "/" & TexttoNote(NotetoText(Key + 6))
                            End If
                        ElseIf LLTxt = ChkNote(6) & "m7" Or LLTxt = ChkNote(6) & "m" Or LLTxt = ChkNote(6) & "min" Then
                            Reset i - 3, NotetoText(Key + 6), "m7b5"
                        End If
                    End If
                End If
            End If
        End If
    Next i
MsgBox "리하모니제이션이 완료되었습니다", vbExclamation, "리하모니제이션 도구"
End If
Exit Sub
ErrMSG:
MsgBox "Description : " & Err.Description & vbCrLf & "Number: " & Err.Number, vbCritical, "Error"
End Sub

Private Sub Form_Load()
labkey = "Key : " & NotetoText(Key)
End Sub

Public Sub Reset(Number As Long, ByVal Root As String, ByVal Kind As String)

nMsg = MsgBox("(" & i - 1 & ") " & Rtxt & "를 " & Root & Kind & "로 바꾸시겠습니까?", vbYesNo, "리하모니제이션 도구")

If nMsg = 6 Then
    mdi_frmMain.ActiveForm.lstChord.ListItems(Number).Text = Root & Kind & "/" & Root
    mdi_frmMain.ActiveForm.lstChord.ListItems(Number).ListSubItems(1).Text = TexttoNote(Root) & "/" & Kind & "/" & "0" & "/" & TexttoNote(Root)
    mdi_frmMain.ActiveForm.lstChord.ListItems(Number).ForeColor = vbRed
End If
End Sub

Public Function ChkNote(Number As String) As String
ChkNote = NotetoText(Key + Number)
End Function

Private Sub TabStrip2_Click()
If TabStrip2.SelectedItem.Index = 1 Then
    pBox1.Visible = True
    pBox2.Visible = False
ElseIf TabStrip2.SelectedItem.Index = 2 Then
    pBox1.Visible = False
    pBox2.Visible = True
End If
End Sub


Public Function ResetPos(x As Long, Listtext1 As String, Listtext2 As String)
On Error Resume Next
Dim i As Long

If mdi_frmMain.ActiveForm.lstChord.ListItems.Count = 0 Then
    mdi_frmMain.ActiveForm.lstChord.ListItems.Add , , Listtext1
    mdi_frmMain.ActiveForm.lstChord.ListItems(mdi_frmMain.ActiveForm.lstChord.ListItems.Count).ListSubItems.Add.Text = Listtext2
    mdi_frmMain.ActiveForm.lstChord.ListItems(mdi_frmMain.ActiveForm.lstChord.ListItems.Count).ForeColor = vbBlue
Else
    x = mdi_frmMain.ActiveForm.lstChord.ListItems.Count - x
    mdi_frmMain.ActiveForm.lstChord.ListItems.Add , , ""
    mdi_frmMain.ActiveForm.lstChord.ListItems(mdi_frmMain.ActiveForm.lstChord.ListItems.Count).Text = mdi_frmMain.ActiveForm.lstChord.ListItems(mdi_frmMain.ActiveForm.lstChord.ListItems.Count - 1).Text
    mdi_frmMain.ActiveForm.lstChord.ListItems(mdi_frmMain.ActiveForm.lstChord.ListItems.Count).ListSubItems.Add.Text = mdi_frmMain.ActiveForm.lstChord.ListItems(mdi_frmMain.ActiveForm.lstChord.ListItems.Count - 1).ListSubItems(1).Text
    For i = 2 To x - 1
        mdi_frmMain.ActiveForm.lstChord.ListItems(mdi_frmMain.ActiveForm.lstChord.ListItems.Count - (i - 1)).Text = mdi_frmMain.ActiveForm.lstChord.ListItems(mdi_frmMain.ActiveForm.lstChord.ListItems.Count - i).Text
        mdi_frmMain.ActiveForm.lstChord.ListItems(mdi_frmMain.ActiveForm.lstChord.ListItems.Count - (i - 1)).ListSubItems(1).Text = mdi_frmMain.ActiveForm.lstChord.ListItems(mdi_frmMain.ActiveForm.lstChord.ListItems.Count - i).ListSubItems(1).Text
    Next i
    mdi_frmMain.ActiveForm.lstChord.ListItems(mdi_frmMain.ActiveForm.lstChord.ListItems.Count - (x - 1)).Text = Listtext1
    mdi_frmMain.ActiveForm.lstChord.ListItems(mdi_frmMain.ActiveForm.lstChord.ListItems.Count - (x - 1)).ListSubItems(1).Text = Listtext2
    mdi_frmMain.ActiveForm.lstChord.ListItems(mdi_frmMain.ActiveForm.lstChord.ListItems.Count - (x - 1)).ForeColor = vbMagenta
End If

End Function

