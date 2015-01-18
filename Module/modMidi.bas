Attribute VB_Name = "modMidi"
Option Explicit

Public Const MAXPNAMELEN = 32

Type MIDIOUTCAPS
   wMid As Integer
   wPid As Integer
   vDriverVersion As Long
   szPname As String * MAXPNAMELEN
   wTechnology As Integer
   wVoices As Integer
   wNotes As Integer
   wChannelMask As Integer
   dwSupport As Long
End Type

Declare Function midiOutGetNumDevs Lib "winmm" () As Integer
Declare Function midiOutGetDevCaps Lib "winmm.dll" Alias "midiOutGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As MIDIOUTCAPS, ByVal uSize As Long) As Long
Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
Declare Function midiOutOpen Lib "winmm.dll" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long

Public Sub InitializInstrument(Value) 'Initializ Instrument
On Error GoTo ErrIns
Dim midiMsg As Long
If Value = 128 Then 'Percussion
    channel = 9
Else
    channel = 0
    midiMsg = (Value * &H100) + &HC0 + channel
    midiOutShortMsg hMidi, midiMsg
End If
Exit Sub
ErrIns:
MsgBox "초기화 실패", vbCritical, "미디 장치"
End Sub

Public Sub KeyMapping(Value)
On Error Resume Next
Dim temp() As String
Dim x As Long
    For x = 300 To 347
        temp = Split(LoadResString(x), ",")
        KeyMap(CLng(temp(0))) = CLng(temp(Value))
    Next x
KeyMap(16) = 88
End Sub

Public Sub InitializeMidi() 'Initialize Midi
On Error GoTo ErrIns
Dim x As Long
midiOutClose hMidi
rc = midiOutOpen(hMidi, curDevice, 0, 0, 0)
If rc = 4 Then
    MsgBox "초기화 실패", vbCritical, "미디 장치"
End If
Exit Sub
ErrIns:
MsgBox "초기화 실패", vbCritical, "미디 장치"
End Sub

Public Sub LoadDevice()
Dim caps As MIDIOUTCAPS
numDevices = midiOutGetNumDevs()
For i = -1 To (numDevices - 1)
If i = -1 Then
    midiOutGetDevCaps i, caps, Len(caps)
    mdi_frmMain.mnuOutput(i + 1).Caption = caps.szPname
Else
    midiOutGetDevCaps i, caps, Len(caps)
    Load mdi_frmMain.mnuOutput(i + 1)
    mdi_frmMain.mnuOutput(i + 1).Caption = caps.szPname
End If
Next i
End Sub
