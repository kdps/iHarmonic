Attribute VB_Name = "modExport"
Option Explicit
Public Style As Integer
Dim NoteBass, NoteBass2, TopBass, strStyle As Long
Dim ListItem() As String
Dim Backup() As String
Dim strPlay, lStyle As Long
Dim midiMsg As Long

Public Function Export()
On Error GoTo Failed
Dim bStop As Boolean
Dim i As Integer
Dim lTempo As Long
Dim sFileMidi As String
Dim sText As String
Dim Notes As eNote
Const NUMLOOP As Integer = 100
Const Tempo As Byte = 121

'Screen.MousePointer = 11

bStop = False

sFileMidi = App.Path + "\Midi.mid"

Call MidiInitialize(6, sFileMidi)

Call MidiSetSubDivision(Tempo)
lTempo = CLng(CSng(1000 / Tempo) * 60000)
Call MidiStart(lTempo, 4, 2, Tempo, 8, 0, 0)

sText = "iHarmonic Export File"
Call MidiMetaSequence(sText)
Call MidiMetaText(0, "iHarmonic")
Call MidiMetaCopyRight("iHarmonic")

sText = "Midi File"
Call MidiMetaTrackName(1, sText)
sText = "Piano"
Call MidiMetaInstrument(1, sText)
Call MidiPitchBend(1, [Channel 1], Standard)
Call MidiProgram(1, [Channel 1], [Acoustic Grand Piano])
Call MidiController(1, [Channel 1], MAIN_VOLUME, 100)
Call MidiController(1, [Channel 1], EXTERNAL_FX_DEPTH, 64)
Call MidiController(1, [Channel 1], CHORUS_DEPTH, 64)

t = 1

For z = t To mdi_frmMain.ActiveForm.lstChord.ListItems.Count

'CalcAll ListItem(0), ListItem(1)
ListItem() = Split(mdi_frmMain.ActiveForm.lstChord.ListItems(z).ListSubItems(1).Text, "/") 'Split Note

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


If (ListItem(0) - Key) > 7 Then
    ListItem(0) = ListItem(0)
ElseIf (ListItem(0) - Key) < 0 Then
    ListItem(0) = ListItem(0) + 12
End If

Note() = Split(CalcNote((12 * (2 + 1)) + ListItem(0), ListItem(1), ListItem(2)), ",") 'Split Note

i = UBound(Note)

If UBound(ListItem) > 2 Then
    NoteBass = 12 * (2 + 1) + ListItem(3)
    NoteBass2 = 12 * (2 - 1) + Note(2)
    TopBass = 12 * (2 + 1) + ListItem(0)

    If ((ListItem(0) - Key) < 1 And ListItem(3) = 1) Or ((ListItem(0) - Key) > 7 And ListItem(3) = 1) Then
        ListItem(0) = ListItem(0)
    ElseIf ((ListItem(0) - Key)) > 7 And ListItem(3) = 2 Then
        ListItem(0) = ListItem(0)
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

Dim q As Long

For q = 0 To UBound(Note)
    Call MidiNoteOn(1, [Channel 1], Note(q) - 1 + 24, 127)
Next q

Call MidiWait(1, 480) '1 per 60

For q = 0 To UBound(Note)
    Call MidiNoteOn(1, [Channel 1], Note(q) - 1 + 24, 0)
Next q

Pass:

Next z

Call MidiSync
Call MidiFinish
bStop = True
Exit Function
Failed:
MsgBox "미디 익스포트 실패!", vbCritical
End Function
