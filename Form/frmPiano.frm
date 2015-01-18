VERSION 5.00
Begin VB.Form frmPiano 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '단일 고정
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10080
   Icon            =   "frmPiano.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   10080
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   1
      Left            =   135
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   0
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   3
      Left            =   390
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   0
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   6
      Left            =   855
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   0
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   8
      Left            =   1095
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   0
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   10
      Left            =   1350
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   0
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   13
      Left            =   1815
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   0
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   15
      Left            =   2055
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   0
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   18
      Left            =   2535
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   0
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   20
      Left            =   2775
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   0
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   22
      Left            =   3015
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   0
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   25
      Left            =   3495
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   0
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   27
      Left            =   3750
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   0
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   30
      Left            =   4215
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   0
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   32
      Left            =   4455
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   0
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   34
      Left            =   4695
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   0
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   37
      Left            =   5175
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   0
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   39
      Left            =   5415
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   0
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   42
      Left            =   5895
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   0
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   44
      Left            =   6135
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   0
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   46
      Left            =   6375
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   0
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   58
      Left            =   8055
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   56
      Left            =   7815
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   54
      Left            =   7575
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   51
      Left            =   7095
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   49
      Left            =   6855
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   63
      Left            =   8775
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   61
      Left            =   8535
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   66
      Left            =   9240
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   68
      Left            =   9480
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   70
      Left            =   9720
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   71
      Left            =   9840
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   3
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   69
      Left            =   9600
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   4
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   67
      Left            =   9360
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   5
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   65
      Left            =   9120
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   6
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   64
      Left            =   8880
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   34
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   62
      Left            =   8640
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   35
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   60
      Left            =   8400
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   36
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   59
      Left            =   8160
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   37
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   57
      Left            =   7920
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   38
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   55
      Left            =   7680
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   39
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   53
      Left            =   7440
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   40
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   52
      Left            =   7200
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   41
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   50
      Left            =   6960
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   42
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   48
      Left            =   6720
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   43
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   47
      Left            =   6480
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   44
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   45
      Left            =   6240
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   45
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   43
      Left            =   6000
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   46
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   41
      Left            =   5760
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   47
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   40
      Left            =   5520
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   48
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   38
      Left            =   5280
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   49
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   36
      Left            =   5040
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   50
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   35
      Left            =   4800
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   51
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   33
      Left            =   4560
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   52
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   31
      Left            =   4320
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   53
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   29
      Left            =   4080
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   54
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   28
      Left            =   3840
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   55
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   26
      Left            =   3600
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   56
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   24
      Left            =   3360
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   57
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   23
      Left            =   3120
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   58
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   21
      Left            =   2880
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   59
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   19
      Left            =   2640
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   60
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   17
      Left            =   2400
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   61
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   16
      Left            =   2160
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   62
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   14
      Left            =   1920
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   63
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   12
      Left            =   1680
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   64
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   11
      Left            =   1440
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   65
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   9
      Left            =   1200
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   66
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   7
      Left            =   960
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   67
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   5
      Left            =   720
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   68
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   4
      Left            =   480
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   69
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   2
      Left            =   240
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   70
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   0
      Left            =   0
      MousePointer    =   10  '위쪽 화살표
      Style           =   1  '그래픽
      TabIndex        =   71
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "frmPiano"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub pKey_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

Dim Note As Integer

Note = KeyMap(KeyCode)
If Len(Note) Then
    If Not Note = lNote And Note Then
        PlayNote Note
    End If
End If

End Sub

Private Sub pKey_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

Dim Note As Integer

Note = KeyMap(KeyCode)
If Len(Note) Then
    StopNote (Note)
End If

End Sub

Private Sub pKey_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

PlayNote Index + 1

End Sub

Private Sub pKey_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

StopNote Index + 1

End Sub
