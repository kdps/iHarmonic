Attribute VB_Name = "modPlaySong"
Option Explicit
Public Repeat, rPlay As Boolean
Public bBass As Boolean
Public Inv As Integer
Public Play() As String
Public Note() As String
Public Active As Boolean
Public i, l, z, k, p, t, s, e As Long
Public Style As Integer
Dim NoteBass, NoteBass2, TopBass, strStyle As Long
Dim ListItem() As String
Dim Backup() As String
Dim strPlay, lStyle As Long
Dim midiMsg As Long

Public Function PlaySong()

On Error GoTo ReturnEnd

Repeat = False
strStyle = 0
bBass = True
lStyle = 0
p = 0

For z = t To mdi_frmMain.ActiveForm.lstChord.ListItems.Count

'CalcAll ListItem(0), ListItem(1)
ListItem() = Split(mdi_frmMain.ActiveForm.lstChord.ListItems(z).ListSubItems(1).Text, "/") 'Split Note

If z > 2 Then
    mdi_frmMain.ActiveForm.lstChord.ListItems(z - 2).Bold = False
    mdi_frmMain.ActiveForm.lstChord.ListItems(z - 1).Bold = True
    frmPlayer.sldPlay.value = z
End If

Select Case ListItem(0)
    Case "Style"
        Style = ListItem(1)
        GoTo Pass
    Case "Comment"
        Select Case ListItem(3)
            Case "[:"
                s = z
                GoTo Pass
            Case ":]"
                If Repeat = False Then
                    e = z - 1
                    z = s - 1
                    Repeat = True
                End If
            Case "┌"
                If Repeat = True Then
                    z = e
                End If
            Case "A"
                lStyle = 0
                strStyle = 0
            Case "B"
                lStyle = 1
                strStyle = 0
            Case "C"
                lStyle = 2
                strStyle = 0
            Case "D"
                lStyle = 3
                strStyle = 0
            Case Else
                GoTo Pass
        End Select
        GoTo Pass
    Case "Time"
        p = Str(ListItem(1)) - 1
        GoTo Pass
End Select

mdi_frmMain.ActiveForm.labChord = mdi_frmMain.ActiveForm.lstChord.ListItems(z).Text '& chdClassic

If (ListItem(0) - Key) > 7 Then
    ListItem(0) = ListItem(0) - 12
ElseIf (ListItem(0) - Key) < 0 Then
    ListItem(0) = ListItem(0) + 12
End If

Note() = Split(CalcNote((12 * (frmPlayer.sldOctave.value + 1)) + ListItem(0), ListItem(1), ListItem(2)), ",") 'Split Note

i = UBound(Note)

If UBound(ListItem) > 2 Then
    NoteBass = 12 * (frmPlayer.sldOctave.value + 1) + ListItem(3)
    NoteBass2 = 12 * (frmPlayer.sldOctave.value - 1) + Note(2)
    TopBass = 12 * (frmPlayer.sldOctave.value + 1) + ListItem(0)

    If ((ListItem(0) - Key) < 1 And ListItem(3) = 1) Or ((ListItem(0) - Key) > 7 And ListItem(3) = 1) Then
        ListItem(0) = ListItem(0) - 12
    ElseIf ((ListItem(0) - Key)) > 7 And ListItem(3) = 2 Then
        ListItem(0) = ListItem(0) - 12
    End If

    If (ListItem(3) - ListItem(0)) < 0 Then
        Inv = ListItem(0) - ListItem(3)
    Else
        Inv = ListItem(3) - ListItem(0)
    End If
    
End If

If UBound(ListItem) > 3 And frmPlayer.chkVoice.value = 1 Then

    If Not Inv = 0 And Inv < 5 And Len(Inv) Then   'Invention 1 C/E 1~4
        If i = 2 Then
            If frmPlayer.chkOpen.value = 1 Then
                ReDim Backup(3)
                Note(0) = Note(0) + 12 'C
                'E
                Note(2) = Note(2) 'G
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Note(0) = Backup(2)
                Note(1) = Backup(3) + 12
                Note(2) = Backup(1)
            Else
                ReDim Backup(3)
                Note(0) = Note(0) + 12 'C
                'E
                Note(2) = Note(2) 'G
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Note(0) = Backup(2)
                Note(1) = Backup(3)
                Note(2) = Backup(1)
            End If
        ElseIf i = 3 Then
            If frmPlayer.chkOpen.value = 1 Then
                ReDim Backup(4)
                Note(0) = Note(0) + 12 'C
                'E
                Note(2) = Note(2) 'G
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Backup(4) = Note(3)
                Note(0) = Backup(2)
                Note(1) = Backup(3) + 12
                Note(2) = Backup(1)
                Note(3) = Backup(4) + 12
            Else
                ReDim Backup(4)
                Note(0) = Note(0) + 12 'C
                'E
                Note(2) = Note(2) 'G
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Backup(4) = Note(3)
                Note(0) = Backup(2)
                Note(1) = Backup(3)
                Note(2) = Backup(1)
                Note(3) = Backup(4)
            End If
        End If
    End If
    
    If Not Inv = 4 And Inv < 8 And Inv > 4 Then  'Invention 2 C/G 5~7
        If i = 2 Then
            If frmPlayer.chkOpen.value = 1 Then
                ReDim Backup(3)
                Note(0) = Note(0) + 12 'C
                Note(1) = Note(1) + 12 'E
                'G
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Note(0) = Backup(3) - 12
                Note(1) = Backup(1)
                Note(2) = Backup(2) - 12
            Else
                ReDim Backup(3)
                Note(0) = Note(0) + 12 'C
                Note(1) = Note(1) + 12 'E
                'G
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Note(0) = Backup(3)
                Note(1) = Backup(1)
                Note(2) = Backup(2)
            End If
        ElseIf i = 3 Then
            If frmPlayer.chkOpen.value = 1 Then
                ReDim Backup(4)
                Note(0) = Note(0) + 12 'C
                Note(1) = Note(1) + 12 'E
                'G
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Backup(4) = Note(3)
                Note(0) = Backup(3) - 12
                Note(1) = Backup(1)
                Note(2) = Backup(2) - 12
                Note(3) = Backup(4)
            Else
                ReDim Backup(4)
                Note(0) = Note(0) + 12 'C
                Note(1) = Note(1) + 12 'E
                'G
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Backup(4) = Note(3)
                Note(0) = Backup(3)
                Note(1) = Backup(1)
                Note(2) = Backup(2)
                Note(3) = Backup(4)
            End If
        End If
    End If
    
    If Not Inv = 7 And Inv < 12 And Inv > 7 Then     'Invention 3 CM7/B 8~11
        If i > 2 Then
            If frmPlayer.chkOpen.value = 1 Then
                ReDim Backup(4)
                Note(0) = Note(0) 'C
                Note(1) = Note(1) 'E
                Note(2) = Note(2) 'G
                Note(3) = Note(3) - 12 'B
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Backup(4) = Note(3)
                Note(0) = Backup(4) - 12
                Note(1) = Backup(1)
                Note(2) = Backup(2) - 12
                Note(3) = Backup(3)
            Else
                ReDim Backup(4)
                Note(0) = Note(0) 'C
                Note(1) = Note(1) 'E
                Note(2) = Note(2) 'G
                Note(3) = Note(3) - 12 'B
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Backup(4) = Note(3)
                Note(0) = Backup(4)
                Note(1) = Backup(1)
                Note(2) = Backup(2)
                Note(3) = Backup(3)
            End If
        End If
    End If
    
End If

If UBound(ListItem) > 3 And frmPlayer.chkVoice.value = 0 Then

    If ListItem(4) = "1" Then 'Invention 1 C/E
        If i = 2 Then
            If frmPlayer.chkOpen.value = 1 Then
                ReDim Backup(3)
                Note(0) = Note(0) + 12 'C
                Note(2) = Note(2) 'G
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Note(0) = Backup(2)
                Note(1) = Backup(3) + 12
                Note(2) = Backup(1)
            Else
                ReDim Backup(3)
                Note(0) = Note(0) + 12 'C
                Note(2) = Note(2) 'G
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Note(0) = Backup(2)
                Note(1) = Backup(3)
                Note(2) = Backup(1)
            End If
        ElseIf i = 3 Then
            If frmPlayer.chkOpen.value = 1 Then
                ReDim Backup(4)
                Note(0) = Note(0) + 12 'C
                Note(2) = Note(2) 'G
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Backup(4) = Note(3)
                Note(0) = Backup(2)
                Note(1) = Backup(3) + 12
                Note(2) = Backup(1)
                Note(3) = Backup(4) + 12
            Else
                ReDim Backup(4)
                Note(0) = Note(0) + 12 'C
                Note(2) = Note(2) 'G
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Backup(4) = Note(3)
                Note(0) = Backup(2)
                Note(1) = Backup(3)
                Note(2) = Backup(1)
                Note(3) = Backup(4)
            End If
        End If
    End If
    
    If ListItem(4) = "2" Then  'Invention 2 C/G
        If i = 2 Then
            If frmPlayer.chkOpen.value = 1 Then
                ReDim Backup(3)
                Note(0) = Note(0) + 12 'C
                Note(1) = Note(1) + 12 'E
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Note(0) = Backup(3) - 12
                Note(1) = Backup(1)
                Note(2) = Backup(2) - 12
            Else
                ReDim Backup(3)
                Note(0) = Note(0) + 12 'C
                Note(1) = Note(1) + 12 'E
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Note(0) = Backup(3) - 12
                Note(1) = Backup(1) - 12
                Note(2) = Backup(2) - 12
            End If
        ElseIf i = 3 Then
            If frmPlayer.chkOpen.value = 1 Then
                ReDim Backup(4)
                Note(0) = Note(0) + 12 'C
                Note(1) = Note(1) + 12 'E
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Backup(4) = Note(3)
                Note(0) = Backup(3) - 12
                Note(1) = Backup(1)
                Note(2) = Backup(2) - 12
                Note(3) = Backup(4)
            Else
                ReDim Backup(4)
                Note(0) = Note(0) + 12 'C
                Note(1) = Note(1) + 12 'E
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Backup(4) = Note(3)
                Note(0) = Backup(3) - 12
                Note(1) = Backup(1) - 12
                Note(2) = Backup(2) - 12
                Note(3) = Backup(4) - 12
            End If
        End If
    End If
    
    If ListItem(4) = "3" Then     'Invention 3 CM7/B
        If i > 2 Then
            If frmPlayer.chkOpen.value = 1 Then
                ReDim Backup(4)
                Note(0) = Note(0) 'C
                Note(1) = Note(1) 'E
                Note(2) = Note(2) 'G
                Note(3) = Note(3) - 12 'B
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Backup(4) = Note(3)
                Note(0) = Backup(4) - 12
                Note(1) = Backup(1)
                Note(2) = Backup(2) - 12
                Note(3) = Backup(3)
            Else
                ReDim Backup(4)
                Note(0) = Note(0) 'C
                Note(1) = Note(1) 'E
                Note(2) = Note(2) 'G
                Note(3) = Note(3) - 12 'B
                Backup(1) = Note(0)
                Backup(2) = Note(1)
                Backup(3) = Note(2)
                Backup(4) = Note(3)
                Note(0) = Backup(4) - 12
                Note(1) = Backup(1) - 12
                Note(2) = Backup(2) - 12
                Note(3) = Backup(3) - 12
            End If
        End If
    End If
End If

If UBound(ListItem) = "3" Then 'Invention 1 C/C
    If frmPlayer.chkOpen.value = 1 Then
        If i = 2 Then
            Note(1) = Note(1) + 12
        ElseIf i = 3 Then
            Note(1) = Note(1) + 12
            Note(2) = Note(2) + 12
        End If
    End If
End If

rPlay = True

StylePlay

Do
DoEvents
Loop Until rPlay = False

Dim x As Long
For x = 1 To 71 '(stop all notes)
    StopNote x
    StopBass x
Next x

Pass:

For x = 1 To 71 '(stop all notes)
    StopNote x
    StopBass x
Next x

Next z

Exit Function

ReturnEnd:

End Function

Public Function StylePlay()
Dim middles As String
midiMsg = (1 * &H100) + &HC0 + Channel
midiOutShortMsg hMidi, midiMsg

Select Case Style
Case 1
PlayNote NoteBass - 24

If lStyle = 0 Then
    For l = 0 To p
        If i = 2 Then
            PlayNote Note(1)
            PlayNote Note(2)
            Timer Tempo * 2
            StopNote Note(1)
            StopNote Note(2)
            PlayNote Note(0)
            Timer Tempo * 2
            StopNote Note(0)
        ElseIf i = 3 Then
            PlayNote Note(1)
            PlayNote Note(2)
            PlayNote Note(3)
            Timer Tempo * 2
            StopNote Note(1)
            StopNote Note(2)
            StopNote Note(3)
            PlayNote Note(0)
            Timer Tempo * 2
            StopNote Note(0)
        ElseIf i = 4 Then
            PlayNote Note(1)
            PlayNote Note(2)
            PlayNote Note(3)
            PlayNote Note(4)
            Timer Tempo * 2
            StopNote Note(1)
            StopNote Note(2)
            StopNote Note(3)
            StopNote Note(4)
            PlayNote Note(0)
            Timer Tempo * 2
            StopNote Note(0)
        ElseIf i = 5 Then
            PlayNote Note(1)
            PlayNote Note(2)
            PlayNote Note(3)
            PlayNote Note(4)
            PlayNote Note(5)
            Timer Tempo * 2
            StopNote Note(1)
            StopNote Note(2)
            StopNote Note(3)
            StopNote Note(4)
            StopNote Note(5)
            PlayNote Note(0)
            Timer Tempo * 2
            StopNote Note(0)
        End If
    Next l
    GoTo Pass
ElseIf lStyle = 1 Then
    For l = 0 To p
        If i = 2 Then
            PlayNote Note(0)
            PlayNote Note(1)
            PlayNote Note(2)
            Timer Tempo * 4
            StopNote Note(0)
            StopNote Note(1)
            StopNote Note(2)
        ElseIf i = 3 Then
            PlayNote Note(0)
            PlayNote Note(1)
            PlayNote Note(2)
            PlayNote Note(3)
            Timer Tempo * 4
            StopNote Note(0)
            StopNote Note(1)
            StopNote Note(2)
            StopNote Note(3)
        ElseIf i = 4 Then
            PlayNote Note(0)
            PlayNote Note(1)
            PlayNote Note(2)
            PlayNote Note(3)
            PlayNote Note(4)
            Timer Tempo * 4
            StopNote Note(0)
            StopNote Note(1)
            StopNote Note(2)
            StopNote Note(3)
            StopNote Note(4)
        ElseIf i = 5 Then
            PlayNote Note(0)
            PlayNote Note(1)
            PlayNote Note(2)
            PlayNote Note(3)
            PlayNote Note(4)
            PlayNote Note(5)
            Timer Tempo * 4
            StopNote Note(0)
            StopNote Note(1)
            StopNote Note(2)
            StopNote Note(3)
            StopNote Note(4)
            StopNote Note(5)
        End If
    Next l
    GoTo Pass
ElseIf lStyle = 2 Then
    For l = 0 To p
        PlayNote NoteBass
        PlayNote NoteBass - 12
        PlayScale i, True, 86, False
        Timer Tempo * 2
        PlayScale i, False, 86, False
        Timer Tempo * 2
        StopNote NoteBass
        StopNote NoteBass - 12
        PlayScale i, True, 86, False
        Timer Tempo * 2
        PlayScale i, False, 86, False
        PlayNote NoteBass
        PlayNote NoteBass - 12
        Timer Tempo * 2
        StopNote NoteBass
        StopNote NoteBass - 12
    Next l
ElseIf lStyle = 3 Then
        If p > 0 Then
            For l = 0 To p - 1
                If i = 2 Then
                    PlayNote Note(0)
                    Timer Tempo * 2
                    StopNote Note(0)
                    PlayNote Note(1)
                    Timer Tempo * 2
                    StopNote Note(1)
                    PlayNote Note(2)
                    Timer Tempo * 4
                    StopNote Note(2)
                ElseIf i = 3 Then
                    PlayNote Note(3)
                    PlayNote Note(0)
                    Timer Tempo * 2
                    StopNote Note(0)
                    StopNote Note(3)
                    PlayNote Note(1)
                    Timer Tempo * 2
                    StopNote Note(1)
                    PlayNote Note(2)
                    Timer Tempo * 4
                    StopNote Note(2)
                ElseIf i = 4 Then
                    PlayNote Note(3)
                    PlayNote Note(4)
                    PlayNote Note(0)
                    Timer Tempo * 2
                    StopNote Note(0)
                    StopNote Note(4)
                    StopNote Note(3)
                    PlayNote Note(1)
                    Timer Tempo * 2
                    StopNote Note(1)
                    PlayNote Note(2)
                    Timer Tempo * 4
                    StopNote Note(2)
                ElseIf i = 5 Then
                    PlayNote Note(3)
                    PlayNote Note(4)
                    PlayNote Note(5)
                    PlayNote Note(0)
                    Timer Tempo * 2
                    StopNote Note(0)
                    StopNote Note(3)
                    StopNote Note(4)
                    StopNote Note(5)
                    PlayNote Note(1)
                    Timer Tempo * 2
                    StopNote Note(1)
                    PlayNote Note(2)
                    Timer Tempo * 4
                    StopNote Note(2)
                End If
            Next l
        ElseIf p = 0 Then
            PlayScale i, True, 72, False
            Timer Tempo * 4
            PlayScale i, False, 72, False
        End If
End If


Case 2
midiMsg = (4 * &H100) + &HC0 + Channel
midiOutShortMsg hMidi, midiMsg

midiMsg = (32 * &H100) + &HC0 + 11
midiOutShortMsg hMidi, midiMsg

bBass = True
If lStyle = 0 Then
    If strStyle = 0 Then
        PlayBass NoteBass - 12
        PlayScale i, True, 86, True
        PlayDrum 19
        PlayDrum 14
        PlayDrum 12
        Timer Tempo
        PlayDrum 19
        Timer Tempo
        PlayDrum 19
        PlayDrum 15
        PlayDrum 14
        Timer Tempo
        PlayDrum 19
        PlayDrum 12
        StopBass NoteBass - 12
        PlayBass NoteBass2 - 24
        Timer Tempo
        PlayDrum 19
        PlayDrum 12
        StopBass NoteBass2 - 24
        PlayBass NoteBass2 - 24
        Timer Tempo
        PlayDrum 19
        PlayDrum 14
        Timer Tempo
        PlayDrum 19
        PlayDrum 15
        Timer Tempo
        PlayDrum 19
        StopBass NoteBass2 - 24
        PlayBass NoteBass - 12
        Timer Tempo
        PlayScale i, False, 86, False
        strStyle = strStyle + 1
        GoTo Pass
    ElseIf strStyle = 1 Then
        PlayScale i, True, 86, False
        StopBass NoteBass - 12
        PlayBass NoteBass - 12
        PlayDrum 19
        PlayDrum 14
        PlayDrum 12
        Timer Tempo
        PlayDrum 19
        Timer Tempo
        PlayDrum 19
        PlayDrum 15
        Timer Tempo
        PlayDrum 19
        PlayDrum 14
        PlayDrum 12
        Timer Tempo
        PlayDrum 19
        PlayDrum 12
        StopBass NoteBass - 12
        PlayBass NoteBass2 - 24
        Timer Tempo
        PlayDrum 19
        Timer Tempo
        PlayDrum 19
        PlayDrum 15
        PlayDrum 14
        StopBass NoteBass2 - 24
        PlayBass NoteBass2 - 24
        Timer Tempo
        PlayDrum 19
        PlayBass NoteBass - 12
        Timer Tempo
        PlayScale i, False, 86, False
        StopBass NoteBass - 12
        strStyle = 0
        GoTo Pass
    End If
End If

If lStyle = 1 Then
    If strStyle = 0 Then
        PlayBass NoteBass - 12
        PlayScale i, True, 86, True
        PlayDrum 19
        PlayDrum 14
        PlayDrum 12
        Timer Tempo
        PlayDrum 19
        Timer Tempo
        PlayDrum 19
        PlayDrum 15
        PlayDrum 14
        Timer Tempo
        PlayDrum 19
        PlayDrum 12
        StopBass NoteBass - 12
        PlayBass NoteBass2 - 24
        Timer Tempo
        PlayScale i, False, 86, False
        PlayDrum 19
        PlayDrum 12
        StopBass NoteBass2 - 24
        PlayBass NoteBass2 - 24
        Timer Tempo
        PlayScale i, True, 86, False
        PlayDrum 19
        PlayDrum 14
        Timer Tempo
        PlayScale i, False, 86, False
        PlayDrum 19
        PlayDrum 15
        Timer Tempo
        PlayDrum 19
        StopBass NoteBass2 - 24
        PlayBass NoteBass - 12
        Timer Tempo
        strStyle = strStyle + 1
        GoTo Pass
    ElseIf strStyle = 1 Then
        StopBass NoteBass - 12
        PlayBass NoteBass - 12
        PlayDrum 19
        PlayDrum 14
        PlayDrum 12
        Timer Tempo
        PlayScale i, True, 86, True
        PlayDrum 19
        Timer Tempo
        PlayScale i, False, 86, False
        PlayDrum 19
        PlayDrum 15
        Timer Tempo
        PlayDrum 19
        PlayDrum 14
        PlayDrum 12
        Timer Tempo
        PlayScale i, True, 86, False
        PlayDrum 19
        PlayDrum 12
        StopBass NoteBass - 12
        PlayBass NoteBass2 - 24
        Timer Tempo
        PlayDrum 19
        Timer Tempo
        PlayDrum 19
        PlayDrum 15
        PlayDrum 14
        StopBass NoteBass2 - 24
        PlayBass NoteBass2 - 24
        Timer Tempo
        PlayDrum 19
        PlayBass NoteBass - 12
        Timer Tempo
        PlayScale i, False, 86, False
        StopBass NoteBass - 12
        strStyle = 0
        GoTo Pass
    End If
End If

If lStyle = 2 Then
    If strStyle = 0 Then
        PlayScale i, True, 86, False
        PlayDrum 58
        PlayDrum 19
        Timer Tempo
        PlayDrum 19
        Timer Tempo
        PlayDrum 57
        PlayDrum 15
        PlayDrum 19
        Timer Tempo
        PlayDrum 57
        PlayDrum 19
        Timer Tempo
        PlayScale i, False, 86, False
        PlayDrum 58
        PlayDrum 19
        Timer Tempo
        PlayScale i, True, 86, False
        PlayDrum 19
        Timer Tempo
        PlayScale i, False, 86, False
        PlayDrum 57
        PlayDrum 15
        PlayDrum 19
        Timer Tempo
        PlayDrum 57
        PlayDrum 19
        Timer Tempo
        strStyle = 0
        GoTo Pass
    End If
End If

If lStyle = 3 Then
    If strStyle = 0 Then
        PlayBass NoteBass - 12
        PlayScale i, True, 86, False
        PlayDrum 58
        PlayDrum 19
        PlayDrum 14
        PlayDrum 12
        Timer Tempo
        PlayDrum 19
        Timer Tempo
        PlayDrum 57
        PlayDrum 19
        PlayDrum 15
        PlayDrum 14
        Timer Tempo
        PlayDrum 57
        PlayDrum 19
        PlayDrum 12
        StopBass NoteBass - 12
        PlayBass NoteBass2 - 24
        Timer Tempo
        PlayDrum 58
        PlayScale i, False, 86, False
        PlayDrum 19
        PlayDrum 12
        StopBass NoteBass2 - 24
        PlayBass NoteBass2 - 24
        Timer Tempo
        PlayScale i, True, 86, False
        PlayDrum 19
        PlayDrum 14
        Timer Tempo
        PlayDrum 57
        PlayScale i, False, 86, False
        PlayDrum 19
        PlayDrum 15
        Timer Tempo
        PlayDrum 57
        PlayDrum 19
        StopBass NoteBass2 - 24
        PlayBass NoteBass - 12
        Timer Tempo
        strStyle = strStyle + 1
        GoTo Pass
    ElseIf strStyle = 1 Then
        StopBass NoteBass - 12
        PlayBass NoteBass - 12
        PlayDrum 58
        PlayDrum 19
        PlayDrum 14
        PlayDrum 12
        Timer Tempo
        PlayScale i, True, 86, False
        PlayDrum 19
        Timer Tempo
        PlayScale i, False, 86, False
        PlayDrum 57
        PlayDrum 19
        PlayDrum 15
        Timer Tempo
        PlayDrum 57
        PlayDrum 19
        PlayDrum 14
        PlayDrum 12
        Timer Tempo
        PlayScale i, True, 86, False
        PlayDrum 58
        PlayDrum 19
        PlayDrum 12
        StopBass NoteBass - 12
        PlayBass NoteBass2 - 24
        Timer Tempo
        PlayDrum 57
        PlayDrum 19
        Timer Tempo
        PlayDrum 57
        PlayDrum 19
        PlayDrum 15
        PlayDrum 14
        StopBass NoteBass2 - 24
        PlayBass NoteBass2 - 24
        Timer Tempo
        PlayDrum 19
        PlayBass NoteBass - 12
        Timer Tempo
        PlayScale i, False, 86, False
        StopBass NoteBass - 12
        strStyle = 0
        GoTo Pass
    End If
End If

Case 3

bBass = False

midiMsg = (34 * &H100) + &HC0 + Channel
midiOutShortMsg hMidi, midiMsg

midiMsg = (7 * &H100) + &HC0 + 5
midiOutShortMsg hMidi, midiMsg

midiMsg = (8 * &H100) + &HC0 + 11
midiOutShortMsg hMidi, midiMsg
If strStyle = 0 Then
    If lStyle = 0 Then
        PlayBass NoteBass - 12
        PlayScale i, True, 86, False
        PlayDrum 19
        PlayDrum 12
        Timer Tempo * 2
        PlayDrum 19
        Timer Tempo * 2
        PlayScale i, False, 86, False
        PlayScale i, True, 86, False
        PlayDrum 39
        PlayDrum 37
        PlayDrum 31
        PlayDrum 19
        PlayDrum 17
        PlayDrum 12
        Timer Tempo * 2
        PlayScale i, False, 86, False
        PlayScale i, True, 86, False
        PlayDrum 24
        PlayDrum 19
        Timer Tempo
        PlayBass NoteBass - 12
        PlayScale i, False, 86, False
        PlayScale i, True, 86, False
        PlayDrum 39
        PlayDrum 37
        PlayDrum 17
        Timer Tempo
        PlayBass NoteBass - 12
        PlayDrum 24
        PlayDrum 18
        Timer Tempo
        PlayScale i, False, 86, False
        PlayScale i, True, 86, False
        PlayDrum 39
        PlayDrum 37
        PlayDrum 17
        Timer Tempo
        PlayScale i, False, 86, False
        PlayScale i, True, 86, False
        PlayDrum 24
        PlayDrum 19
        Timer Tempo * 2
        PlayScale i, False, 86, False
        PlayScale i, True, 86, False
        PlayDrum 39
        PlayDrum 37
        PlayDrum 31
        PlayDrum 19
        PlayDrum 16
        PlayDrum 12
        Timer Tempo * 2
        PlayBass NoteBass - 12
        PlayDrum 19
        Timer Tempo * 2
        strStyle = 0
        GoTo Pass
    ElseIf lStyle = 1 Then
        PlayString NoteBass - 12
        PlayBass NoteBass - 12
        PlayScale i, True, 86, False
        PlayDrum 19
        PlayDrum 12
        Timer Tempo * 2
        PlayDrum 19
        Timer Tempo * 2
        PlayScale i, False, 86, False
        PlayScale i, True, 86, False
        PlayDrum 39
        PlayDrum 37
        PlayDrum 31
        PlayDrum 19
        PlayDrum 17
        PlayDrum 12
        Timer Tempo * 2
        PlayScale i, False, 86, False
        PlayScale i, True, 86, False
        PlayDrum 24
        PlayDrum 19
        Timer Tempo
        PlayBass NoteBass - 12
        PlayScale i, False, 86, False
        PlayScale i, True, 86, False
        PlayDrum 39
        PlayDrum 37
        PlayDrum 17
        Timer Tempo
        PlayBass NoteBass - 12
        PlayDrum 24
        PlayDrum 18
        Timer Tempo
        PlayScale i, False, 86, False
        PlayScale i, True, 86, False
        PlayDrum 39
        PlayDrum 37
        PlayDrum 17
        Timer Tempo
        PlayScale i, False, 86, False
        PlayScale i, True, 86, False
        PlayDrum 24
        PlayDrum 19
        Timer Tempo * 2
        PlayScale i, False, 86, False
        PlayScale i, True, 86, False
        PlayDrum 39
        PlayDrum 37
        PlayDrum 31
        PlayDrum 19
        PlayDrum 16
        PlayDrum 12
        Timer Tempo * 2
        PlayBass NoteBass - 12
        PlayDrum 19
        Timer Tempo
        PlayString NoteBass2 - 12
        Timer Tempo
        strStyle = 0
        GoTo Pass
    End If
End If

Case 4
midiMsg = (1 * &H100) + &HC0 + Channel
midiOutShortMsg hMidi, midiMsg
midiMsg = (1 * &H100) + &HC0 + 11
midiOutShortMsg hMidi, midiMsg

For l = 0 To p
If strStyle = 0 Then
    If lStyle = 0 Then
        PlayBass NoteBass - 12
        PlayDrum 47
        PlayDrum 26
        PlayDrum 14
        Timer Tempo + (Tempo / 2)
        PlayDrum 46
        PlayDrum 21
        Timer Tempo / 2
        PlayDrum 47
        PlayDrum 15
        StopBass NoteBass - 12
        PlayScale i, True, 86, False
        PlayBass NoteBass - 12
        Timer Tempo + (Tempo / 2)
        PlayDrum 46
        PlayDrum 21
        Timer Tempo / 2
        PlayScale i, False, 86, False
        StopBass NoteBass - 12
        strStyle = strStyle + 1
    ElseIf lStyle = 1 Then
        PlayBass NoteBass - 12
        PlayDrum 47
        PlayDrum 26
        PlayDrum 14
        Timer Tempo + (Tempo / 2)
        PlayDrum 46
        PlayDrum 21
        Timer Tempo / 2
        PlayDrum 47
        PlayDrum 15
        StopBass NoteBass - 12
        PlayScale i, True, 86, False
        PlayBass NoteBass - 12
        Timer Tempo + (Tempo / 2)
        PlayDrum 46
        PlayDrum 21
        Timer Tempo / 2
        PlayScale i, False, 86, False
        StopBass NoteBass - 12
        strStyle = strStyle + 1
    ElseIf lStyle = 2 Then
        StopBass NoteBass - 12
        PlayScale i, True, 86, False
        PlayDrum 47
        PlayDrum 36
        PlayDrum 14
        Timer Tempo + (Tempo / 2)
        PlayDrum 46
        PlayDrum 36
        Timer Tempo / 2
        PlayDrum 47
        PlayDrum 36
        Timer Tempo + (Tempo / 2)
        PlayDrum 46
        PlayDrum 36
        PlayDrum 13
        Timer Tempo / 2
        strStyle = strStyle + 1
    ElseIf lStyle = 3 Then
        PlayBass NoteBass - 12
        PlayDrum 47
        PlayDrum 26
        PlayDrum 14
        Timer Tempo + (Tempo / 2)
        PlayDrum 46
        PlayDrum 21
        Timer Tempo / 2
        PlayDrum 47
        PlayDrum 16
        PlayDrum 15
        StopBass NoteBass - 12
        PlayScale i, True, 86, False
        PlayBass NoteBass - 12
        Timer Tempo + (Tempo / 2)
        PlayDrum 46
        PlayDrum 21
        Timer Tempo / 2
        PlayScale i, False, 86, False
        StopBass NoteBass - 12
        strStyle = strStyle + 1
    End If
ElseIf strStyle = 1 Then
    If lStyle = 0 Then
        PlayDrum 46
        PlayBass NoteBass - 12
        Timer Tempo + (Tempo / 2)
        PlayScale i, True, 86, False
        PlayDrum 47
        PlayDrum 21
        PlayDrum 15
        Timer Tempo / 2
        PlayDrum 46
        PlayDrum 15
        PlayScale i, False, 86, False
        StopBass NoteBass - 12
        PlayBass NoteBass - 12
        Timer Tempo + (Tempo / 2)
        PlayDrum 46
        PlayDrum 21
        Timer Tempo / 2
        StopBass NoteBass - 12
        strStyle = strStyle + 1
    ElseIf lStyle = 1 Then
        PlayDrum 46
        PlayBass NoteBass - 12
        Timer Tempo + (Tempo / 2)
        PlayDrum 47
        PlayDrum 21
        PlayDrum 15
        Timer Tempo / 2
        PlayScale i, True, 86, False
        PlayDrum 46
        PlayDrum 15
        StopBass NoteBass - 12
        PlayBass NoteBass - 12
        Timer Tempo + (Tempo / 2)
        PlayDrum 46
        PlayDrum 21
        Timer Tempo / 2
        PlayScale i, False, 86, False
        StopBass NoteBass - 12
        strStyle = strStyle + 1
    ElseIf lStyle = 2 Then
        PlayScale i, True, 86, False
        PlayDrum 47
        PlayDrum 36
        PlayDrum 15
        Timer Tempo + (Tempo / 2)
        PlayDrum 46
        PlayDrum 36
        Timer Tempo / 2
        PlayDrum 47
        PlayDrum 36
        Timer Tempo + (Tempo / 2)
        PlayDrum 46
        PlayDrum 36
        PlayDrum 14
        Timer Tempo / 2
        strStyle = 0
    ElseIf lStyle = 3 Then
        PlayDrum 46
        PlayBass NoteBass - 12
        Timer Tempo + (Tempo / 2)
        PlayScale i, True, 86, False
        PlayDrum 47
        PlayDrum 21
        PlayDrum 15
        Timer Tempo / 2
        PlayDrum 46
        PlayDrum 15
        PlayScale i, False, 86, False
        StopBass NoteBass - 12
        PlayBass NoteBass - 12
        Timer Tempo + (Tempo / 2)
        PlayDrum 46
        PlayDrum 21
        Timer Tempo / 2
        StopBass NoteBass - 12
        strStyle = strStyle + 1
    End If
ElseIf strStyle = 2 Then
    If lStyle = 0 Then
        PlayDrum 47
        PlayDrum 14
        PlayBass NoteBass - 12
        PlayScale i, True, 86, False
        Timer Tempo + (Tempo / 2)
        PlayDrum 46
        PlayDrum 21
        Timer (Tempo / 2)
        PlayDrum 46
        PlayDrum 15
        PlayScale i, False, 86, False
        StopBass NoteBass - 12
        PlayBass NoteBass - 12
        Timer Tempo + (Tempo / 2)
        PlayDrum 46
        strStyle = strStyle + 1
    ElseIf lStyle = 1 Then
        PlayDrum 47
        PlayDrum 14
        PlayBass NoteBass - 12
        PlayScale i, True, 86, False
        Timer Tempo + (Tempo / 2)
        PlayDrum 46
        PlayDrum 21
        Timer (Tempo / 2)
        PlayDrum 46
        PlayDrum 15
        PlayScale i, False, 86, False
        StopBass NoteBass - 12
        PlayBass NoteBass - 12
        Timer Tempo + (Tempo / 2)
        PlayDrum 46
        strStyle = strStyle + 1
    ElseIf lStyle = 3 Then
        PlayDrum 47
        PlayDrum 14
        PlayBass NoteBass - 12
        PlayScale i, True, 86, False
        Timer Tempo + (Tempo / 2)
        PlayDrum 46
        PlayDrum 21
        Timer (Tempo / 2)
        PlayDrum 46
        PlayDrum 15
        PlayScale i, False, 86, False
        StopBass NoteBass - 12
        PlayBass NoteBass - 12
        Timer Tempo + (Tempo / 2)
        PlayDrum 46
        strStyle = strStyle + 1
    End If
ElseIf strStyle = 3 Then
    If lStyle = 0 Then
        PlayBass NoteBass - 12
        PlayScale i, True, 86, False
        Timer Tempo / 2
        PlayDrum 46
        PlayScale i, False, 86, False
        Timer Tempo + (Tempo / 2)
        PlayDrum 47
        PlayDrum 21
        PlayDrum 15
        StopBass NoteBass - 12
        PlayBass NoteBass - 12
        Timer Tempo / 2
        StopBass NoteBass - 12
        PlayScale i, True, 86, False
        PlayBass NoteBass - 12
        PlayDrum 46
        PlayDrum 15
        Timer Tempo + (Tempo / 2)
        PlayDrum 46
        PlayDrum 21
        PlayDrum 15
        Timer Tempo / 2
        PlayScale i, False, 86, False
        StopBass NoteBass - 12
        strStyle = 0
    ElseIf lStyle = 1 Then
        PlayBass NoteBass - 12
        Timer Tempo / 2
        PlayScale i, True, 86, False
        PlayDrum 46
        Timer Tempo + (Tempo / 2)
        PlayDrum 47
        PlayDrum 21
        PlayDrum 15
        StopBass NoteBass - 12
        PlayBass NoteBass - 12
        Timer Tempo / 2
        PlayScale i, False, 86, False
        PlayScale i, True, 86, False
        StopBass NoteBass - 12
        PlayBass NoteBass - 12
        PlayDrum 46
        PlayDrum 15
        Timer Tempo + (Tempo / 2)
        PlayDrum 46
        PlayDrum 21
        PlayDrum 15
        Timer Tempo / 2
        PlayScale i, False, 86, False
        StopBass NoteBass - 12
        strStyle = 0
    ElseIf lStyle = 3 Then
        PlayBass NoteBass - 12
        PlayScale i, True, 86, False
        Timer Tempo / 2
        PlayDrum 46
        PlayScale i, False, 86, False
        Timer Tempo + (Tempo / 2)
        PlayDrum 47
        PlayDrum 21
        PlayDrum 15
        StopBass NoteBass - 12
        PlayBass NoteBass - 12
        Timer Tempo / 2
        StopBass NoteBass - 12
        PlayScale i, True, 86, False
        PlayBass NoteBass - 12
        PlayDrum 46
        PlayDrum 16
        PlayDrum 15
        Timer Tempo + (Tempo / 2)
        PlayDrum 46
        PlayDrum 21
        PlayDrum 15
        Timer Tempo / 2
        PlayScale i, False, 86, False
        StopBass NoteBass - 12
        strStyle = 0
    End If
End If
Next l
GoTo Pass
End Select

Pass:
rPlay = False

End Function
