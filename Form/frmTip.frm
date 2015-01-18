VERSION 5.00
Begin VB.Form frmTip 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "알고 계십니까?"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton cmdBack 
      Caption         =   "< 이전"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "다음 >"
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton mnuClose 
      Caption         =   "닫기(&C)"
      Height          =   375
      Left            =   7200
      TabIndex        =   1
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      Caption         =   "팁"
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      Begin VB.Label labTitle 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   960
         TabIndex        =   5
         Top             =   360
         Width           =   7335
      End
      Begin VB.Label labTip 
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   960
         TabIndex        =   4
         Top             =   1080
         Width           =   7335
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmTip.frx":0000
         Top             =   240
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmtip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Index As String
Public maxin As String

Public Sub Tip()
Select Case Index
Case 1
labTitle = "투파이브원"
labTip = "IIm7 - V7 - I은 코드의 이동이 2 - 5 - 1와 같이 진행해서 투파이브원이라고 부릅니다." & vbCrLf & _
         "스케일의 두번째 다섯번째의 음정에 m7과 7을 붙이면 투파이브를 쉽게 구할 수 있습니다." & vbCrLf & _
         "투파이브원은 완전4도 상행 완전4도 하행하므로 귀에 자연스럽게 들립니다."
Case 2
labTitle = "불협화음과 협화음"
labTip = "완전4도 완전5도 완전8도(완전1도)는 협화음이고 불협화음은 장3도 단3도 장6도 단6도," & vbCrLf & _
         "불협화음은 장2도 단2도 장7도 단7도 그리고 모든 증,감음정입니다."
Case 3
labTitle = "자리바꿈"
labTip = "원래 음정에서 9를 빼면 자리바꿈 음정의 도수가 나옵니다." & vbCrLf & vbCrLf & _
         "그 음정이 바뀔때의 특징은 완전음정은 완전음정, 단음정은 장음정으로 장음정은 단음정," & vbCrLf & _
         "감음정은 증음정으로 증음정은 감음정으로 바뀝니다."
Case 4
labTitle = "세컨더리 도미넌트"
labTip = "세컨더리 도미넌트는 다이아토닉 코드(Bm7b5,F#m7b5 및 VIIm7b5,#IVm7b5)를 제외하고" & vbCrLf & _
         "완전5도 상행 및 완전4도 하행하는 도미넌트라고 생각하면 됩니다." & vbCrLf & _
         "세컨더리의 입장에서는 완전4도 상행 및 완전5도 하행하여 다이아토닉으로 진행하는것이 됩니다."
Case 5
labTitle = "섭스티튜트 도미넌트 7th"
labTip = "세컨더리 도미넌트에서 감5도상행 및 증4도하행한 도미넌트입니다."
Case 6
labTitle = "서브도미넌트 마이너"
labTip = "서브도미넌트 마이너는 나란한조(C>Cm)에 있는 코드의 4번째 및 서브도미넌트를 빌려쓴 코드입니다."
Case 7
labTitle = "패싱 디미니쉬드"
labTip = "패싱 디미니쉬드는 세컨더리 도미넌트에 b9음을 붙여 근음을 제거한 코드입니다." & vbCrLf & _
        "세컨더리 도미넌트 자체의 기능을 가지고 있습니다"
Case 8
labTitle = "텐션"
labTip = "①9th와 b9th, 9th, #9th, b13th, 13th, b13th, 5th는 같이 사용하면 안됩니다." & vbCrLf & _
        "②#11th를 사용하면 5th를 생략하는게 좋습니다." & vbCrLf & _
        "③11th와 #11th를 사용할시 9th를 포합하는게 좋습니다." & vbCrLf & _
        "④#11th를 5th와 함께 하용할시에는 반음관계로 부딪히지 않도록 하는게 좋습니다." & vbCrLf & _
        "⑤4성 이상에서 13th를 사용시 9th도 넣는것이 소리가 풍부해집니다." & vbCrLf & _
        "⑥b9th, #9th, 13th, 5th는 사용할때 서로 때놓아야 됩니다" & vbCrLf & _
        "⑦#11th를 5th와 함께 사용할 경우 #11th가 5th 위치에 오는게 좋습니다."
Case 9
labTitle = "투파이브 제네레이터"
labTip = "투파이브 제네레이터는 자동으로 투파이브를 만들어줍니다."
Case 10
labTitle = "나란한조"
labTip = "나란한조는 같은 으뜸음을 가지는 장/단조입니다."
Case 11
labTitle = "거짓마침"
labTip = "거짓마침은 V7에서 I가 아닌 다른 코드들로 진행되는것을 말합니다."
Case 12
labTitle = "T, SD, D의 뜻"
labTip = "각각 토닉, 서브도미넌트, 도미넌트를 뜻합니다."
End Select
End Sub

Private Sub cmdBack_Click()
If Index = 2 Then
    Index = Index - 1
    Call Tip
    cmdBack.Enabled = False
    cmdNext.Enabled = True
Else
    cmdNext.Enabled = True
    Index = Index - 1
    Call Tip
End If
End Sub

Private Sub cmdNext_Click()
If Index = maxin Then
    Index = Index + 1
    Call Tip
    cmdNext.Enabled = False
    cmdBack.Enabled = True
Else
    cmdBack.Enabled = True
    Index = Index + 1
    Call Tip
End If
End Sub

Private Sub Form_Load()
Randomize
maxin = 11
Index = Int((maxin * Rnd) + 1)
If Not Index = 1 Then
    cmdBack.Enabled = True
End If
Call Tip
End Sub

Private Sub mnuClose_Click()
Unload Me
End Sub
