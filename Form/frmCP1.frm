VERSION 5.00
Begin VB.Form frmCP1 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "코드 패턴 - I VIm7 IIm7 V7"
   ClientHeight    =   510
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   510
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.ComboBox cbPattern2 
      Height          =   300
      Left            =   1440
      TabIndex        =   6
      Top             =   120
      Width           =   1935
   End
   Begin VB.ComboBox cbPattern4 
      Height          =   300
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox cbPattern3 
      Height          =   300
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox cbPattern1 
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lbLine 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "-"
      Height          =   180
      Index           =   2
      Left            =   4800
      TabIndex        =   5
      Top             =   150
      Width           =   90
   End
   Begin VB.Label lbLine 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "-"
      Height          =   180
      Index           =   1
      Left            =   3480
      TabIndex        =   4
      Top             =   150
      Width           =   90
   End
   Begin VB.Label lbLine 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "-"
      Height          =   180
      Index           =   0
      Left            =   1200
      TabIndex        =   3
      Top             =   150
      Width           =   90
   End
End
Attribute VB_Name = "frmCP1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
chkShape = True
cbPattern1.AddItem NotetoText(Key) & "M7"
cbPattern1.AddItem NotetoText(Key + 4) & "m7"
cbPattern1.AddItem NotetoText(Key + 9) & "m7"
cbPattern1.AddItem NotetoText(Key) & "6"

cbPattern2.AddItem NotetoText(Key + 9) & "m7"
cbPattern2.AddItem NotetoText(Key + 9) & "7"
cbPattern2.AddItem NotetoText(Key + 1) & "dim7"
chkShape = False
cbPattern2.AddItem NotetoText(Key + 3) & "dim7"
chkShape = True
cbPattern2.AddItem NotetoText(Key + 2) & "7"
chkShape = False
cbPattern2.AddItem NotetoText(Key + 3) & "7"
cbPattern2.AddItem NotetoText(Key + 8) & "7"
cbPattern2.AddItem NotetoText(Key + 3) & "m7b5"
cbPattern2.AddItem NotetoText(Key + 3) & "m7b5 - " & NotetoText(Key + 9) & "7"
cbPattern2.AddItem NotetoText(Key + 9) & "m7 - " & NotetoText(Key + 2) & "7"
cbPattern2.AddItem NotetoText(Key + 10) & "m7 - " & NotetoText(Key + 3) & "7"
cbPattern2.AddItem NotetoText(Key + 3) & "m7 - " & NotetoText(Key + 10) & "7"
chkShape = True
cbPattern1.ListIndex = 0
cbPattern2.ListIndex = 0
cbPattern3.Text = NotetoText(Key + 2) & "m7"
cbPattern4.Text = NotetoText(Key + 7) & "7"
End Sub
