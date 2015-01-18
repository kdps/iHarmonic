Attribute VB_Name = "modOutput"
Option Explicit
'***********************
'      Translated      *
'      by GioRock      *
'***********************
'***********************
'  Written by GioRock  *
'***********************
'***********************
'  Special Thanks to   *
'     David Madore     *
' david.madore@ens.fr  *
'***********************

Public Enum eNote
    B0 = 11
    C1 = 12
    Cd1 = 13
    D1 = 14
    Dd1 = 15
    E1 = 16
    F1 = 17
    Fd1 = 18
    G1 = 19
    Gd1 = 20
    A1 = 21
    Ad1 = 22
    B1 = 23
    C2 = 24
    Cd2 = 25
    D2 = 26
    Dd2 = 27
    E2 = 28
    F2 = 29
    Fd2 = 30
    G2 = 31
    Gd2 = 32
    A2 = 33
    Ad2 = 34
    B2 = 35
    C3 = 36
    Cd3 = 37
    D3 = 38
    Dd3 = 39
    E3 = 40
    F3 = 41
    Fd3 = 42
    G3 = 43
    Gd3 = 44
    A3 = 45
    Ad3 = 46
    B3 = 47
    C4 = 48
    Cd4 = 49
    D4 = 50
    Dd4 = 51
    E4 = 52
    F4 = 53
    Fd4 = 54
    G4 = 55
    Gd4 = 56
    A4 = 57
    Ad4 = 58
    B4 = 59
    C5 = 60
    Cd5 = 61
    D5 = 62
    Dd5 = 63
    E5 = 64
    F5 = 65
    Fd5 = 66
    G5 = 67
    Gd5 = 68
    A5 = 69
    Ad5 = 70
    B5 = 71
    C6 = 72
    Cd6 = 73
    D6 = 74
    Dd6 = 75
    E6 = 76
    F6 = 77
    Fd6 = 78
    G6 = 79
    Gd6 = 80
    A6 = 81
    Ad6 = 82
    B6 = 83
    C7 = 84
    Cd7 = 85
    D7 = 86
    Dd7 = 87
    E7 = 88
    F7 = 89
    Fd7 = 90
    G7 = 91
    Gd7 = 92
    A7 = 93
    Ad7 = 94
    B7 = 95
    C8 = 96
    Cd8 = 97
    D8 = 98
    Dd8 = 99
    E8 = 100
    F8 = 101
    Fd8 = 102
    G8 = 103
    Gd8 = 104
    A8 = 105
    Ad8 = 106
    B8 = 107
    C9 = 108
    Cd9 = 109
    D9 = 110
    Dd9 = 111
    E9 = 112
    F9 = 113
    Fd9 = 114
    G9 = 115
    Gd9 = 116
    A9 = 117
    Ad9 = 118
    B9 = 119
    C10 = 120
    Cd10 = 121
    D10 = 122
    Dd10 = 123
    E10 = 124
    F10 = 125
    Fd10 = 126
    G10 = 127
    [Acoustic Bass Drum] = 35
    [Bass Drum 1] = 36
    [Side Stick] = 37
    [Acoustic Snare] = 38
    [Hand Clap] = 39
    [Electric Snare] = 40
    [Low Floor Tom] = 41
    [Closed High Hat] = 42
    [High Floor Tom] = 43
    [Pedal High Hat] = 44
    [Low Tom] = 45
    [Open High Hat] = 46
    [Low Mid Tom] = 47
    [High Mid Tom] = 48
    [Crash Cymbal 1] = 49
    [High Tom] = 50
    [Ride Cymbal 1] = 51
    [Chinese Cymbal] = 52
    [Ride Bell] = 53
    [Tambourine] = 54
    [Splash Cymbal] = 55
    [Cowbell] = 56
    [Crash Cymbal 2] = 57
    [Vibraslap] = 58
    [Ride Cymbal 2] = 59
    [High Bongo] = 60
    [Low Bongo] = 61
    [Mute High Conga] = 62
    [Open High Conga] = 63
    [Low Conga] = 64
    [High Timbale] = 65
    [Low Timbale] = 66
    [High Agogo] = 67
    [Low Agogo] = 68
    [Cabasa] = 69
    [Maracas] = 70
    [Short Whistle] = 71
    [Long Whistle] = 72
    [Short Guiro] = 73
    [Long Guiro] = 74
    [Claves] = 75
    [High Wood Block] = 76
    [Low Wood Block] = 77
    [Mute Cuica] = 78
    [Open Cuica] = 79
    [Mute Triangle] = 80
    [Open Triangle] = 81
End Enum

Public Enum eInstrument
    [Acoustic Grand Piano] = 0
    [Bright Acoustic Piano] = 1
    [Electric Grand Piano] = 2
    [Honky Tonk Piano] = 3
    [Electric Piano 1] = 4
    [Electric Piano 2] = 5
    [Harpsichord] = 6
    [Clav] = 7
    [Celesta] = 8
    [Glockenspiel] = 9
    [Music Box] = 10
    [Vibraphone] = 11
    [Marimba] = 12
    [Xylophone] = 13
    [Tubular Bells] = 14
    [Dulcimer] = 15
    [Drawbar Organ] = 16
    [Percussive Organ] = 17
    [Rock Organ] = 18
    [Church Organ] = 19
    [Reed Organ] = 20
    [Accordian] = 21
    [Harmonica] = 22
    [Tango Accordian] = 23
    [Nylon Guitar] = 24
    [Steel Guitar] = 25
    [Jazz Electric Guitar] = 26
    [Clean Electric Guitar] = 27
    [Muted Electric Guitar] = 28
    [Overdriven Guitar] = 29
    [Distortion Guitar] = 30
    [Guitar Harmonics] = 31
    [Acoustic Bass] = 32
    [Finger Electric Bass] = 33
    [Pick Electric Bass] = 34
    [Fretless Bass] = 35
    [Slap Bass 1] = 36
    [Slap Bass 2] = 37
    [Synth Bass 1] = 38
    [Synth Bass 2] = 39
    [Violin] = 40
    [Viola] = 41
    [Cello] = 42
    [Contrabass] = 43
    [Tremolo Strings] = 44
    [Pizzicato Strings] = 45
    [Orchestral Strings] = 46
    [Timpani] = 47
    [String Ensemble 1] = 48
    [String Ensemble 2] = 49
    [Synth Strings 1] = 50
    [Synth Strings 2] = 51
    [Choir Aahs] = 52
    [Voice Oohs] = 53
    [Synth Voice] = 54
    [Orchestra Hit] = 55
    [Trumpet] = 56
    [Trombone] = 57
    [Tuba] = 58
    [Muted Trumpet] = 59
    [French Horn] = 60
    [Brass Section] = 61
    [Synth Brass 1] = 62
    [Synth Brass 2] = 63
    [Soprano Sax] = 64
    [Alto Sax] = 65
    [Tenor Sax] = 66
    [Baritone Sax] = 67
    [Oboe] = 68
    [English Horn] = 69
    [Bassoon] = 70
    [Clarinet] = 71
    [Piccolo] = 72
    [Flute] = 73
    [Recorder] = 74
    [Pan Flute] = 75
    [Blown Bottle] = 76
    [Shakuhachi] = 77
    [Whistle] = 78
    [Ocarina] = 79
    [Lead 1 Square] = 80
    [Lead 2 Sawtooth] = 81
    [Lead 3 Calliope] = 82
    [Lead 4 Chiff] = 83
    [Lead 5 Charang] = 84
    [Lead 6 Voice] = 85
    [Lead 7 Fifths] = 86
    [Lead 8 Bass Lead] = 87
    [Pad 1 New Age] = 88
    [Pad 2 Warm] = 89
    [Pad 3 Polysynth] = 90
    [Pad 4 Choir] = 91
    [Pad 5 Bowed] = 92
    [Pad 6 Metallic] = 93
    [Pad 7 Halo] = 94
    [Pad 8 Sweep] = 95
    [FX 1 Rain] = 96
    [FX 2 Soundtrack] = 97
    [FX 3 Crystal] = 98
    [FX 4 Atmosphere] = 99
    [FX 5 Brightness] = 100
    [FX 6 Goblins] = 101
    [FX 7 Echoes] = 102
    [FX 8 Sci Fi] = 103
    [Sitar] = 104
    [Banjo] = 105
    [Shamisen] = 106
    [Koto] = 107
    [Kalimba] = 108
    [Bagpipe] = 109
    [Fiddle] = 110
    [Shanai] = 111
    [Tinkle Bell] = 112
    [Agogo] = 113
    [Steel Drums] = 114
    [Woodblock] = 115
    [Taiko Drum] = 116
    [Melodic Tom] = 117
    [Synth Drum] = 118
    [Reverse Cymbal] = 119
    [Guitar Fret Noise] = 120
    [Breath Noise] = 121
    [Seashore] = 122
    [Bird Tweet] = 123
    [Telephone Ring] = 124
    [Helicopter] = 125
    [Applause] = 126
    [Gunshot] = 127
    [General MIDI Drum] = 1
End Enum

Public Enum ePitch
    Standard = 4096
    Baroque = 4148
End Enum

Public Enum eChannel
    [Channel 1] = 0
    [Channel 2] = 1
    [Channel 3] = 2
    [Channel 4] = 3
    [Channel 5] = 4
    [Channel 6] = 5
    [Channel 7] = 6
    [Channel 8] = 7
    [Channel 9] = 8
    [Channel 10 - Drum] = 9
    [Channel 11] = 10
    [Channel 12] = 11
    [Channel 13] = 12
    [Channel 14] = 13
    [Channel 15] = 14
    [Channel 16] = 15
End Enum

' MIDI Controller Numbers Constants
Public Enum eController
    MOD_WHEEL = 1
    BREATH_CONTROLLER = 2
    FOOT_CONTROLLER = 4
    PORTAMENTO_TIME = 5
    MAIN_VOLUME = 7
    BALANCE = 8
    PAN = 10
    EXPRESS_CONTROLLER = 11
    DAMPER_PEDAL = 64
    PORTAMENTO = 65
    SOSTENUTO = 66
    SOFT_PEDAL = 67
    HOLD_2 = 69
    EXTERNAL_FX_DEPTH = 91
    TREMELO_DEPTH = 92
    CHORUS_DEPTH = 93
    DETUNE_DEPTH = 94
    PHASER_DEPTH = 95
    DATA_INCREMENT = 96
    DATA_DECREMENT = 97
End Enum

Public Enum eChords
    Major = 0
    Minor = 1
    Major7 = 2
    Minor7 = 3
    sus4 = 4
End Enum

Private Type MidiTracks
    MaxLen As Long
    p As Long
    rs As Byte
    pd As Long
    time As Long
    d() As Byte
End Type
Private Type MidiContexts
    nbVoices As Long
    nbTracks As Long
    Midi() As MidiTracks
    gTime As Long
    SubDiv As Integer
End Type

Private mc As MidiContexts
Private hff As Integer

Private Const ALLOC_UNIT As Long = 8192
Private Const cMHeader As String = "MThd"
Private Const cTHeader As String = "MTrk"

Public Sub MidiNoteOn(ByVal Trk As Integer, ByVal Channel As eChannel, ByVal Note As eNote, ByVal Speed As Byte)
' Turn on note (at Speed) on a given channel.
    MidiPutCommand Trk, &H90 + Channel
    MidiAppend Trk, Note
    MidiAppend Trk, Speed
End Sub

Private Sub DefaultFatal()
    MsgBox Error$(Err), vbCritical
    MidiTerminate
End Sub


Private Sub MidiAppend(ByVal Trk As Integer, ByVal b As Byte)
    
    On Error GoTo GErr
    
    If mc.Midi(Trk).p >= mc.Midi(Trk).MaxLen Then
        mc.Midi(Trk).MaxLen = mc.Midi(Trk).MaxLen + ALLOC_UNIT
        ReDim Preserve mc.Midi(Trk).d(mc.Midi(Trk).MaxLen)
    End If
    
    mc.Midi(Trk).d(mc.Midi(Trk).p) = b
    mc.Midi(Trk).p = mc.Midi(Trk).p + 1
    
    Exit Sub
    
GErr:

    DefaultFatal
    
End Sub

Private Sub MidiInitTrack(ByVal Trk As Integer)
' Initialize and empty a MIDI track.

    On Error GoTo GErr
    
    mc.Midi(Trk).MaxLen = 0
    mc.Midi(Trk).p = 0
    mc.Midi(Trk).rs = &HFF
    mc.Midi(Trk).pd = 0
    mc.Midi(Trk).time = 0
    Erase mc.Midi(Trk).d
    
    Exit Sub
    
GErr:

    DefaultFatal
    
End Sub

Private Sub MidiAppendVar(ByVal Trk As Integer, ByVal v As Long)
Dim buffer As Long, value As Long
' Append a variable-length-encoded quantity to a MIDI track.

   value = v
   buffer = value And &H7F
   
   While value \ 128 > 0
      value = value \ 128
      buffer = buffer * 256
      buffer = buffer Or ((value And &H7F) Or &H80)
   Wend
   
   Do
      MidiAppend Trk, CByte(buffer And &HFF)
      If (buffer And &H80) Then
         buffer = buffer \ 256
      Else
         Exit Do
      End If
   Loop
   
End Sub

Private Sub MidiInit(ByVal nbVoices As Long)
Dim Trk As Long
'  Create and initialize all MIDI tracks.

    On Error GoTo GErr
    
    hff = 0
    mc.SubDiv = 120
    mc.gTime = 0
    mc.nbTracks = nbVoices + 1
    ReDim Preserve mc.Midi(nbVoices)
    
    For Trk = 0 To mc.nbTracks - 1
        MidiInitTrack Trk
    Next Trk
    
    Exit Sub
    
GErr:

    DefaultFatal
    
End Sub

Public Sub MidiInitialize(ByVal nbVoices As Long, ByVal sFileName As String)
    MidiInit nbVoices
    On Error GoTo GErr
    hff = FreeFile
    If Dir$(sFileName) <> "" Then: Kill sFileName
    Open sFileName For Binary Access Write As #hff
    Exit Sub
GErr:
    If hff Then: Close #hff
    DefaultFatal
End Sub

Private Sub MidiPutDelay(ByVal Trk As Integer)
' Dump MIDI track's pending delay.
    MidiAppendVar Trk, mc.Midi(Trk).pd
    mc.Midi(Trk).pd = 0
End Sub

Private Sub MidiPutCommand(ByVal Trk As Integer, ByVal b As Byte)
' Append a command to a MIDI track.
    MidiPutDelay Trk
    If b > &HF0 Or b <> mc.Midi(Trk).rs Then
        MidiAppend Trk, b
        mc.Midi(Trk).rs = b
    End If
End Sub

Public Sub MidiWait(ByVal Trk As Integer, ByVal lenght As Long)
' Increase a MIDI track's pending delay.
    mc.Midi(Trk).pd = mc.Midi(Trk).pd + lenght
    mc.Midi(Trk).time = mc.Midi(Trk).time + lenght
    If mc.Midi(Trk).time > mc.gTime Then
        mc.gTime = mc.Midi(Trk).time
    End If
End Sub

Public Sub MidiSync()
Dim Trk As Integer
' Synchronize all MIDI tracks' notion of time with the furthest one.
    
    For Trk = 0 To mc.nbTracks - 1
        MidiWait Trk, mc.gTime - mc.Midi(Trk).time
    Next Trk
    
End Sub

Private Sub MidiAllWait(ByVal lenght As Long)
' Increase all MIDI tracks' pending delay - synchronize.
    MidiWait 0, lenght
    MidiSync
End Sub

Public Sub MidiProgram(ByVal Trk As Integer, ByVal Channel As eChannel, ByVal Program As eInstrument)
' Set program (instrument) on a given channel.
    MidiPutCommand Trk, &HC0 + Channel
    MidiAppend Trk, Program
End Sub

Public Sub MidiController(ByVal Trk As Integer, ByVal Channel As eChannel, ByVal Controller As eController, ByVal value As Byte)
' Set controller value on a given channel.
    MidiPutCommand Trk, &HB0 + Channel
    MidiAppend Trk, Controller
    MidiAppend Trk, value
End Sub

Public Sub MidiNoteOff(ByVal Trk As Integer, ByVal Channel As eChannel, ByVal Note As eNote, ByVal Speed As Byte)
' Turn off note (at Speed) on a given channel.
    MidiPutCommand Trk, &H80 + Channel
    MidiAppend Trk, Note
    MidiAppend Trk, Speed
End Sub

Public Sub MidiPitchBend(ByVal Trk As Integer, ByVal Channel As eChannel, ByVal value As ePitch)
' Set pitchbend value on a given channel.
' this may be wrong...
    MidiPutCommand Trk, &HE0 + Channel
    MidiAppend Trk, CByte(value And &H7F)
    MidiAppend Trk, CByte(RotateBits(value, 7) And &H7F)
End Sub

Public Sub MidiSetSubDivision(ByVal SubDivTime As Integer)
' Set MIDI file global subdivision.
' (The tempo divided by this value determines the MIDI time unit.)
    mc.SubDiv = SubDivTime
End Sub

Private Sub MidiMetaTempo(ByVal v As Long)
' Set tempo (on track zero necessarily).
' This may occur at any point but it must occur before any note
' is played: this is ensured by midi_start().
    MidiPutCommand 0, &HFF
    MidiAppend 0, &H51
    MidiAppend 0, 3
    MidiAppend 0, CByte(RotateBits(v, 16) And &HFF)
    MidiAppend 0, CByte(RotateBits(v, 8) And &HFF)
    MidiAppend 0, CByte(v And &HFF)
End Sub

Private Sub MidiMetaBeat(ByVal n As Byte, ByVal d As Byte, ByVal c As Byte, ByVal b As Byte)
' Set beat.
' n / 2 ^ d is the division indication.
' c is the number of MIDI time units per beat.
' b is the number of demisemiquavers per beat.
    MidiPutCommand 0, &HFF
    MidiAppend 0, &H58
    MidiAppend 0, 4
    MidiAppend 0, n
    MidiAppend 0, d
    MidiAppend 0, c
    MidiAppend 0, b
End Sub

Private Sub MidiMetaKey(ByVal s As Byte, ByVal mj As Byte)
' Set key.
' s is number of sharps (or negative of number of flats.
' mj is 0 if major, 1 if minor.
    MidiPutCommand 0, &HFF
    MidiAppend 0, &H59
    MidiAppend 0, 2
    MidiAppend 0, s
    MidiAppend 0, mj
End Sub

Private Sub Midi_MetaText(ByVal Trk As Integer, ByVal t As Byte, ByVal MetaText As String)
Dim i As Long, lenght As Long
' Include generic text event t on track trk, with data MetaText.

    lenght = Len(MetaText)
    
    MidiPutCommand Trk, &HFF
    MidiAppend Trk, t
    MidiAppendVar Trk, lenght
    
    For i = 1 To lenght
        MidiAppend Trk, CByte(Asc(Mid$(MetaText, i, 1)))
    Next i
    
End Sub

Public Sub MidiMetaCopyRight(ByVal sCopyRight As String)
' Copyright text (only at beginning of track 0)
    Midi_MetaText 0, 2, sCopyRight
End Sub

Public Sub MidiMetaSequence(ByVal sSequence As String)
' Sequence text (only at beginning of track 0)
    Midi_MetaText 0, 3, sSequence
End Sub

Public Sub MidiMetaTrackName(ByVal Trk As Integer, ByVal sTrackName As String)
' Track name (only at beginning of track)
    Midi_MetaText Trk, 3, sTrackName
End Sub

Public Sub MidiMetaInstrument(ByVal Trk As Integer, ByVal sInstrumentName As String)
' Instrument name (only at beginning of track)
    Midi_MetaText Trk, 4, sInstrumentName
End Sub

Public Sub MidiMetaLyrics(ByVal Trk As Integer, ByVal sLyric As String)
' Lyrics text
    Midi_MetaText Trk, 5, sLyric
End Sub

Public Sub MidiMetaMarker(ByVal Trk As Integer, ByVal lenght As Long, ByVal sMarker As String)
' Marker text
    Midi_MetaText Trk, 6, sMarker
End Sub

Public Sub MidiMetaCuePoint(ByVal Trk As Integer, ByVal lenght As Long, ByVal sCuePoint As String)
'  Cue point
    Midi_MetaText Trk, 7, sCuePoint
End Sub

Public Sub MidiSingleNote(ByVal Trk As Integer, ByVal Channel As eChannel, ByVal Note As Byte, ByVal Speed As Byte, ByVal Duration As Byte)
' Convenience function:
' add a single note Note on track Trk,
' channel Channel and speed Speed for duration Duration.
    MidiNoteOn Trk, Channel, Note, Speed
    MidiWait Trk, Duration
    MidiNoteOn Trk, Channel, Note, 0
End Sub

Public Sub MidiStart(ByVal lTempo As Long, ByVal bn As Byte, ByVal bd As Byte, ByVal bc As Byte, ByVal bb As Byte, ByVal ks As Byte, ByVal km As Byte)
' Start MIDI file: add tempo, beat and key events to the
' controller track.
    MidiMetaTempo lTempo
    MidiMetaBeat bn, bd, bc, bb
    MidiMetaKey ks, km
End Sub

Public Sub MidiFinish()
Dim Trk As Integer
' Finish MIDI file: add track end event to every track.

    MidiSync
    
    For Trk = 0 To mc.nbTracks - 1
        MidiPutCommand Trk, &HFF
        MidiAppend Trk, &H2F
        MidiAppend Trk, &H0
    Next Trk
    
    MidiDump
    
    MidiTerminate
    
End Sub

Private Sub MidiDump()
Dim Trk As Integer
Dim i As Long

    ' Put Midi Header
    Put #hff, , cMHeader
    WriteDWORD 6
    
    ' Put Midi format
    WriteWORD 1
    ' Put Midi Tracks number
    WriteWORD mc.nbTracks
    ' Put Subdivision value
    WriteWORD mc.SubDiv
    
    ' Save all stored data in increasing order
    For Trk = 0 To mc.nbTracks - 1
        ' Put Midi track Header
        Put #hff, , cTHeader
        ' Put structure length
        WriteDWORD mc.Midi(Trk).p
        For i = 0 To mc.Midi(Trk).p - 1
            Put #hff, , mc.Midi(Trk).d(i)
        Next i
    Next Trk
    
End Sub

Private Sub WriteWORD(ByVal w As Integer)
Dim b As Byte
' Write a 16-bit quantity in big-endian format.

    b = CByte(RotateBits(w, 8) And &HFF)
    Put #hff, , b
    b = CByte(w And &HFF)
    Put #hff, , b
    
End Sub

Private Sub WriteDWORD(ByVal dw As Long)
Dim b As Byte
' Write a 32-bit quantity in big-endian format.

    b = CByte(RotateBits(dw, 24) And &HFF)
    Put #hff, , b
    b = CByte(RotateBits(dw, 16) And &HFF)
    Put #hff, , b
    b = CByte(RotateBits(dw, 8) And &HFF)
    Put #hff, , b
    b = CByte(dw And &HFF)
    Put #hff, , b

End Sub


Private Sub MidiTerminate()
Dim Trk As Integer
' Terminate MIDI session.
' Free all resources associated with the MIDI data.

    On Local Error Resume Next
    
    For Trk = 0 To mc.nbTracks - 1
        Erase mc.Midi(Trk).d
    Next Trk
    
    Erase mc.Midi
    
    Close #hff
    
End Sub

Public Sub MidiMetaText(ByVal Trk As Integer, ByVal sText As String)
' Text (only at beginning of track)
    Midi_MetaText Trk, 0, sText
End Sub

Private Function RotateBits(ByVal lValue As Long, ByVal nRot As Byte) As Long
Dim lTemp As Long
Dim lRot As Long
'***********************
'  Written by GioRock  *
'***********************
' Same uses >> in C or C++

    lTemp = lValue
        
    Select Case nRot
        Case 8
            lRot = 256
        Case 16
            lRot = (256 ^ 2)
        Case 24
            lRot = (256 ^ 3)
        Case Else
            lRot = &HF7
    End Select
    
    lTemp = lTemp \ lRot
    
    RotateBits = lTemp
    
End Function

