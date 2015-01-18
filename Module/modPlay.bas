Attribute VB_Name = "modPlay"
Option Explicit
Private Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function MidOutClose Lib "winmm.dll" (ByVal hMidOut As Long) As Long
Public Declare Function MidOutOpen Lib "winmm.dll" (lphMidOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Public Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidOut As Long, ByVal dwMsg As Long) As Long
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Public rc, lNote, hMidi, Channel, numDevices, curDevice, KeyMap(255) As Long
Public Tempo, Pitch, PianoVelocity, Velocity, noteLong As Long
Public Related, Played As Boolean
Public mMScale, chdClassic, chdFunction As String
Public strStyle, Style, Key As Integer 'Style Count
Public ChkSave As Boolean

Public Sub PlayScale(ByVal Scales As String, States As Boolean, Velocitys As String, Slash As Boolean)
On Error Resume Next

PianoVelocity = Velocitys

If States = False Then
    If Scales = 2 Then
        StopNote Note(0)
        StopNote Note(1)
        StopNote Note(2)
    ElseIf Scales = 3 Then
        StopNote Note(0)
        StopNote Note(1)
        StopNote Note(2)
        StopNote Note(3)
    ElseIf Scales = 4 Then
        StopNote Note(0)
        StopNote Note(1)
        StopNote Note(2)
        StopNote Note(3)
        StopNote Note(4)
    ElseIf Scales = 5 Then
        StopNote Note(0)
        StopNote Note(1)
        StopNote Note(2)
        StopNote Note(3)
        StopNote Note(4)
        StopNote Note(5)
    End If
End If

If States = True Then
    If Slash = False Then
        If Scales = 2 Then
            If bBass = True Then PlayNote Note(0)
            PlayNote Note(0)
            PlayNote Note(1)
            PlayNote Note(2)
        ElseIf Scales = 3 Then
            If bBass = True Then PlayNote Note(0)
            PlayNote Note(0)
            PlayNote Note(1)
            'PlayNote Note(2)
            PlayNote Note(3)
        ElseIf Scales = 4 Then
            If bBass = True Then PlayNote Note(0)
            PlayNote Note(0)
            PlayNote Note(1)
            'PlayNote Note(2)
            PlayNote Note(3)
            PlayNote Note(4)
        ElseIf Scales = 5 Then
            If bBass = True Then PlayNote Note(0)
            PlayNote Note(0)
            PlayNote Note(1)
            'PlayNote Note(2)
            PlayNote Note(3)
            PlayNote Note(4)
            PlayNote Note(5)
        End If
    ElseIf Slash = True Then
        If Scales = 2 Then
            If bBass = True Then PlayNote Note(0)
            PlayNote Note(0)
            Timer Tempo / 32
            PlayNote Note(1)
            Timer Tempo / 32
            PlayNote Note(2)
        ElseIf Scales = 3 Then
            If bBass = True Then PlayNote Note(0)
            PlayNote Note(1)
            Timer Tempo / 48
            PlayNote Note(2)
            Timer Tempo / 48
            PlayNote Note(3)
            Timer Tempo / 48
        ElseIf Scales = 4 Then
            If bBass = True Then PlayNote Note(0)
            PlayNote Note(1)
            Timer Tempo / 64
            PlayNote Note(2)
            Timer Tempo / 64
            PlayNote Note(3)
            Timer Tempo / 64
            PlayNote Note(4)
            Timer Tempo / 64
        ElseIf Scales = 5 Then
            If bBass = True Then PlayNote Note(0)
            PlayNote Note(1)
            Timer Tempo / 80
            PlayNote Note(2)
            Timer Tempo / 80
            PlayNote Note(3)
            Timer Tempo / 80
            PlayNote Note(4)
            Timer Tempo / 80
            PlayNote Note(5)
            Timer Tempo / 80
        End If
    End If
End If

Exit Sub
Pass:
End Sub

Public Function PlayBass(ByVal Note As Long)
On Error Resume Next
Dim midiMsg As Long

If frmPlayer.chkBass.value = 1 Then
    midiMsg = &H90 + 11 + ((Pitch + Note) * &H100) + ((Velocity + 32) * &H10000)
    midiOutShortMsg hMidi, midiMsg
End If

If frmPiano.Visible = True Then frmPiano.pKey(Note - 1).BackColor = vbBlue

lNote = Note
End Function

Public Function PlayString(ByVal Note As Long)
On Error Resume Next
Dim midiMsg As Long

If frmPlayer.chkBass.value = 1 Then
    midiMsg = &H90 + 5 + ((Pitch + Note) * &H100) + ((Velocity + 50) * &H10000)
    midiOutShortMsg hMidi, midiMsg
End If

If frmPiano.Visible = True Then frmPiano.pKey(Note - 1).BackColor = vbYellow

lNote = Note
End Function

Public Function StopString(ByVal Note As Long)
On Error Resume Next
Dim midiMsg As Long

If Note = 88 Then
    'Sustain False
Else
    midiMsg = &H80 + ((Pitch + Note) * &H100) + 5
    midiOutShortMsg hMidi, midiMsg
    If frmPiano.Visible = True Then
        If frmPiano.pKey(Note - 1).Tag = "1" Then
            frmPiano.pKey(Note - 1).BackColor = vbWhite
        Else
            frmPiano.pKey(Note - 1).BackColor = vbBlack
        End If
    End If
    If frmPiano2.Visible = True Then frmPiano2.pKeyoff (Note - 1)
End If

If Note = lNote Then lNote = 0
End Function

Public Function StopBass(ByVal Note As Long)
On Error Resume Next
Dim midiMsg As Long

If Note = 88 Then
    'Sustain False
Else
    midiMsg = &H80 + ((Pitch + Note) * &H100) + 11
    midiOutShortMsg hMidi, midiMsg
If frmPiano.pKey(Note - 1).Tag = "1" Then
    If frmPiano.Visible = True Then frmPiano.pKey(Note - 1).BackColor = vbWhite
Else
    If frmPiano.Visible = True Then frmPiano.pKey(Note - 1).BackColor = vbBlack
End If
End If

If Note = lNote Then lNote = 0
End Function

Public Function PlayDrum(ByVal Note As Long)
On Error Resume Next
Dim midiMsg As Long
If frmPlayer.chkDrum.value = 1 Then
    midiMsg = &H90 + 9 + ((Pitch + Note) * &H100) + ((30 + 30) * &H10000) '9 : Percusion
    midiOutShortMsg hMidi, midiMsg
    If frmPiano.Visible = True Then frmPiano.pKey(Note - 1).BackColor = vbGreen
End If

lNote = Note
End Function

Public Function Timer(ByVal value As Long) As Boolean
Dim Tick
    Tick = GetTickCount
    Do While Tick + value > GetTickCount
        DoEvents
    Loop

'mdi_frmMain.tmrPlay.Interval = (Value)
'mdi_frmMain.tmrPlay.Enabled = True
'Do While True
'    If mdi_frmMain.tmrPlay.Enabled = False Then
'        Exit Do
'    End If
'    DoEvents
'Loop
End Function

Public Function Sustain(Active As Boolean) 'Piano Pedal
On Error Resume Next

If Active Then
    'midiOutShortMsg hMidi, (&HB0 + channel + &H4000 + &H7F0000)
Else
    'midiOutShortMsg hMidi, (&HB0 + channel + &H4000)
End If

End Function

Public Function PlayNote(ByVal Note As Long)
On Error Resume Next
Dim midiMsg As Long

If Note = 88 Then
    'Sustain True
Else
    midiMsg = &H90 + Channel + ((Pitch + Note) * &H100) + (PianoVelocity * &H10000)
    midiOutShortMsg hMidi, midiMsg
    If frmPiano.Visible = True Then
        frmPiano.pKey(Note - 1).BackColor = vbRed
    End If
    If frmPiano2.Visible = True Then frmPiano2.pKeyon (Note - 1)
End If

lNote = Note
End Function

Public Function StopNote(ByVal Note As Long)
On Error Resume Next
Dim midiMsg As Long

If Note = 88 Then
    'Sustain False
Else
    midiMsg = &H80 + ((Pitch + Note) * &H100) + Channel
    midiOutShortMsg hMidi, midiMsg
    If frmPiano.Visible = True Then
        If frmPiano.pKey(Note - 1).Tag = "1" Then
            frmPiano.pKey(Note - 1).BackColor = vbWhite
        Else
            frmPiano.pKey(Note - 1).BackColor = vbBlack
        End If
    End If
    If frmPiano2.Visible = True Then frmPiano2.pKeyoff (Note - 1)
End If

If Note = lNote Then lNote = 0
End Function


