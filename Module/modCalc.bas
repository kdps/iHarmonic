Attribute VB_Name = "modCalc"
Option Explicit
Public chkShape As Boolean

Public Function NotetoText(ByVal txtNote As Integer) As String 'Convert Velocity to Note Name
On Error Resume Next

Dim NoteText() As String
Dim NoteText2() As String

Do
DoEvents
txtNote = txtNote - 12
Loop Until txtNote < 12

Do
DoEvents
txtNote = txtNote + 12
Loop Until txtNote > 0

If txtNote > 12 Then
    Do
    DoEvents
    txtNote = txtNote - 12
    Loop Until txtNote < 12
End If

If chkShape = True Then
    NoteText() = Split("C,C#,D,D#,E,F,F#,G,G#,A,A#,B", ",")
    NotetoText = NoteText(txtNote - 1)
Else
    NoteText2() = Split("C,Db,D,Eb,E,F,Gb,G,Ab,A,Bb,B", ",")
    NotetoText = NoteText2(txtNote - 1)
End If

End Function

Public Function NotetoRoma(txtNote) As String 'Convert Velocity to Note Name

On Error Resume Next

Dim NoteRoma() As String

NoteRoma() = Split("I,bII,II,bIII,III,IV,#IV,V,bVI,VI,bVII,VII", ",")

Do
DoEvents
txtNote = txtNote + 12
Loop Until txtNote > 0
    
If txtNote > 12 Then
    Do
    DoEvents
    txtNote = txtNote - 12
    Loop Until txtNote < 12
End If

NotetoRoma = NoteRoma(txtNote - 1)

End Function


Public Function TexttoNote(txtNote) As String

Select Case txtNote
Case "C"
    TexttoNote = 1
Case "C#"
    TexttoNote = 2
Case "Db"
    TexttoNote = 2
Case "D"
    TexttoNote = 3
Case "D#"
    TexttoNote = 4
Case "Eb"
    TexttoNote = 4
Case "E"
    TexttoNote = 5
Case "Fb"
    TexttoNote = 5
Case "F"
    TexttoNote = 6
Case "E#"
    TexttoNote = 6
Case "F#"
    TexttoNote = 7
Case "Gb"
    TexttoNote = 7
Case "G"
    TexttoNote = 8
Case "G#"
    TexttoNote = 9
Case "Ab"
    TexttoNote = 9
Case "A"
    TexttoNote = 10
Case "A#"
    TexttoNote = 11
Case "Bb"
    TexttoNote = 11
Case "B"
    TexttoNote = 12
Case "Cb"
    TexttoNote = 12
End Select

End Function

Public Function AllNotePos(chdRoot, chdType) 'Calculate Chord Note

On Error Resume Next

Dim Note() As String
Dim Tmp As String
Dim i, z As Long

'Array Note
Select Case chdType
    Case ""
        ReDim Note(3)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 7 'Perfect 5
    Case "2"
        ReDim Note(3)
        Note(1) = chdRoot
        Note(2) = chdRoot + 2 'Major 2
        Note(3) = chdRoot + 7 'Perfect 5
    Case "Maj"
        ReDim Note(3)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 7 'Perfect 5
    Case "△"
        ReDim Note(3)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 7 'Perfect 5
    Case "M"
        ReDim Note(3)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 7 'Perfect 5
    Case "sus4"
        ReDim Note(3)
        Note(1) = chdRoot
        Note(2) = chdRoot + 5 'Perfect 4
        Note(3) = chdRoot + 7 'Perfect 5
    Case "-"
        ReDim Note(3)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 7 'Perfect 5
    Case "min"
        ReDim Note(3)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 7 'Perfect 5
    Case "m"
        ReDim Note(3)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 7 'Perfect 5
    Case "+"
        ReDim Note(3)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 8 'Augmented 5
    Case "aug"
        ReDim Note(3)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 8 'Augmented 5
    Case "M7"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 11 'Major 7
    Case "M7sus4"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 5 'Perfect 4
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 11 'Major 7
    Case "Maj7"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 11 'Major 7
    Case "M7b5"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 6 'Perfect 5
        Note(4) = chdRoot + 11 'Major 7
    Case "M7(b5)"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 6 'Perfect 5
        Note(4) = chdRoot + 11 'Major 7
    Case "M7-5"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 6 'Perfect 5
        Note(4) = chdRoot + 11 'Major 7
    Case "M7+5"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 8 'Augmented 5
        Note(4) = chdRoot + 11 'Major 7
    Case "mM7"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Major 3
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 11 'Major 7
    Case "m7"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 10 'Minor 7
    Case "m7+5"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 8 'Augmented 5
        Note(4) = chdRoot + 10 'Minor 7
    Case "dom"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 10 'Minor 7
    Case "7sus4"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 5 'Perfect 4
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 10 'Minor 7
    Case "7"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 10 'Minor 7
    Case "+7"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 8 'Minor 6
        Note(4) = chdRoot + 10 'Minor 7
    Case "7sus4"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 5 'Perfect 4
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 10 'Minor 7
    Case "Dim"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 6 'Diminished 5
    Case "dim"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 6 'Diminished 5
    Case "o"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 6 'Diminished 5
    Case "ø"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 6 'Diminished 5
        Note(4) = chdRoot + 10 'Major 7
    Case "m7b5"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 6 'Diminished 5
        Note(4) = chdRoot + 10 'Major 7
    Case "Dim7"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 6 'Diminished 5
        Note(4) = chdRoot + 9 'Minor 7
    Case "dim7"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 6 'Diminished 5
        Note(4) = chdRoot + 9 'Minor 7
    Case "Dim7sus4"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 5 'Minor 3
        Note(3) = chdRoot + 6 'Diminished 5
        Note(4) = chdRoot + 9 'Minor 7
    Case "dim7sus4"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 5 'Minor 3
        Note(3) = chdRoot + 6 'Diminished 5
        Note(4) = chdRoot + 9 'Minor 7
    Case "M6"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 9 'Major 6
    Case "6"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 9 'Major 6
    Case "m6"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 9 'Major 6
End Select

'Merge Note
    For z = 1 To UBound(Note)
        
        If Note(z) <> "" Then
            If Tmp = "" Then
                Tmp = Tmp & Note(z)
            Else
                Tmp = Tmp & "," & Note(z)
            End If
        End If
        
    Next z

AllNotePos = Tmp 'Output Temp

Tmp = "" 'Cleanup Temp

End Function

Public Function NotePos(chdRoot, chdType, chdPos) 'Calculate Chord Note
On Error Resume Next
Dim Note() As String
Dim Tmp As String
Dim i, z As Long

'Array Note
Select Case chdType
    Case ""
        ReDim Note(3)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 7 'Perfect 5
    Case "2"
        ReDim Note(3)
        Note(1) = chdRoot
        Note(2) = chdRoot + 2 'Major 2
        Note(3) = chdRoot + 7 'Perfect 5
    Case "Maj"
        ReDim Note(3)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 7 'Perfect 5
    Case "△"
        ReDim Note(3)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 7 'Perfect 5
    Case "M"
        ReDim Note(3)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 7 'Perfect 5
    Case "sus4"
        ReDim Note(3)
        Note(1) = chdRoot
        Note(2) = chdRoot + 5 'Perfect 4
        Note(3) = chdRoot + 7 'Perfect 5
    Case "-"
        ReDim Note(3)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 7 'Perfect 5
    Case "min"
        ReDim Note(3)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 7 'Perfect 5
    Case "m"
        ReDim Note(3)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 7 'Perfect 5
    Case "+"
        ReDim Note(3)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 8 'Augmented 5
    Case "aug"
        ReDim Note(3)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 8 'Augmented 5
    Case "M7"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 11 'Major 7
    Case "M7sus4"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 5 'Perfect 4
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 11 'Major 7
    Case "Maj7"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 11 'Major 7
    Case "M7b5"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 6 'Perfect 5
        Note(4) = chdRoot + 11 'Major 7
    Case "M7(b5)"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 6 'Perfect 5
        Note(4) = chdRoot + 11 'Major 7
    Case "M7-5"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 6 'Perfect 5
        Note(4) = chdRoot + 11 'Major 7
    Case "M7+5"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 8 'Augmented 5
        Note(4) = chdRoot + 11 'Major 7
    Case "mM7"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Major 3
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 11 'Major 7
    Case "m7"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 10 'Minor 7
    Case "m7+5"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 8 'Augmented 5
        Note(4) = chdRoot + 10 'Minor 7
    Case "dom"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 10 'Minor 7
    Case "7sus4"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 5 'Perfect 4
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 10 'Minor 7
    Case "7"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 10 'Minor 7
    Case "+7"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 8 'Minor 6
        Note(4) = chdRoot + 10 'Minor 7
    Case "7sus4"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 5 'Perfect 4
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 10 'Minor 7
    Case "Dim"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 6 'Diminished 5
    Case "dim"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 6 'Diminished 5
    Case "o"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 6 'Diminished 5
    Case "ø"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 6 'Diminished 5
        Note(4) = chdRoot + 10 'Major 7
    Case "m7b5"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 6 'Diminished 5
        Note(4) = chdRoot + 10 'Major 7
    Case "Dim7"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 6 'Diminished 5
        Note(4) = chdRoot + 9 'Minor 7
    Case "dim7"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 6 'Diminished 5
        Note(4) = chdRoot + 9 'Minor 7
    Case "Dim7sus4"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 5 'Minor 3
        Note(3) = chdRoot + 6 'Diminished 5
        Note(4) = chdRoot + 9 'Minor 7
    Case "dim7sus4"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 5 'Minor 3
        Note(3) = chdRoot + 6 'Diminished 5
        Note(4) = chdRoot + 9 'Minor 7
    Case "M6"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 9 'Major 6
    Case "6"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 9 'Major 6
    Case "m6"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 9 'Major 6
End Select

'Merge Note
NotePos = Note(chdPos) 'Output Temp

Tmp = "" 'Cleanup Temp

End Function

Public Function CalcKey(keyRoot, keyType) 'Calculate Key Note

On Error Resume Next

Dim Note() As String
Dim Tmp$
Dim i, z As Long

'Array Note
Select Case keyType
    Case "Major"
        ReDim Note(8)
        Note(1) = keyRoot
        Note(2) = keyRoot + 2
        Note(3) = keyRoot + 4
        Note(4) = keyRoot + 5
        Note(6) = keyRoot + 7
        Note(7) = keyRoot + 9
        Note(8) = keyRoot + 11
    Case "Major Pentatonic"
        ReDim Note(5)
        Note(1) = keyRoot
        Note(2) = keyRoot + 2
        Note(3) = keyRoot + 4
        Note(4) = keyRoot + 7
        Note(5) = keyRoot + 9
    Case "Minor"
        ReDim Note(8)
        Note(1) = keyRoot
        Note(2) = keyRoot + 2
        Note(3) = keyRoot + 3
        Note(4) = keyRoot + 5
        Note(6) = keyRoot + 7
        Note(7) = keyRoot + 8
        Note(8) = keyRoot + 10
    Case "Minor Pentatonic"
        ReDim Note(8)
        Note(1) = keyRoot
        Note(2) = keyRoot + 2
        Note(3) = keyRoot + 5
        Note(4) = keyRoot + 7
        Note(5) = keyRoot + 10
    Case "Harmonic Minor"
        ReDim Note(8)
        Note(1) = keyRoot
        Note(2) = keyRoot + 2
        Note(3) = keyRoot + 3
        Note(4) = keyRoot + 5
        Note(6) = keyRoot + 7
        Note(7) = keyRoot + 8
        Note(8) = keyRoot + 11
    Case "Melodic Minor"
        ReDim Note(8)
        Note(1) = keyRoot
        Note(2) = keyRoot + 2
        Note(3) = keyRoot + 3
        Note(4) = keyRoot + 5
        Note(6) = keyRoot + 7
        Note(7) = keyRoot + 9
        Note(8) = keyRoot + 11
    Case "Chromatic"
        ReDim Note(11)
        Note(1) = keyRoot
        Note(2) = keyRoot + 1
        Note(3) = keyRoot + 2
        Note(4) = keyRoot + 3
        Note(6) = keyRoot + 4
        Note(7) = keyRoot + 5
        Note(8) = keyRoot + 6
        Note(9) = keyRoot + 7
        Note(10) = keyRoot + 8
        Note(11) = keyRoot + 9
    Case "Blues"
        ReDim Note(6)
        Note(1) = keyRoot
        Note(2) = keyRoot + 1
        Note(3) = keyRoot + 5
        Note(4) = keyRoot + 6
        Note(5) = keyRoot + 7
        Note(6) = keyRoot + 8
    Case "Whole"
        ReDim Note(6)
        Note(1) = keyRoot
        Note(2) = keyRoot + 2
        Note(3) = keyRoot + 4
        Note(4) = keyRoot + 6
        Note(5) = keyRoot + 8
        Note(6) = keyRoot + 10
    Case "Ionian"
        ReDim Note(8)
        Note(1) = keyRoot
        Note(2) = keyRoot + 2
        Note(3) = keyRoot + 4
        Note(4) = keyRoot + 5
        Note(6) = keyRoot + 7
        Note(7) = keyRoot + 9
        Note(8) = keyRoot + 11
    Case "Lydian"
        ReDim Note(8)
        Note(1) = keyRoot
        Note(2) = keyRoot + 2
        Note(3) = keyRoot + 4
        Note(4) = keyRoot + 6
        Note(6) = keyRoot + 7
        Note(7) = keyRoot + 9
        Note(8) = keyRoot + 11
    Case "Mixolydian"
        ReDim Note(8)
        Note(1) = keyRoot
        Note(2) = keyRoot + 2
        Note(3) = keyRoot + 4
        Note(4) = keyRoot + 5
        Note(6) = keyRoot + 7
        Note(7) = keyRoot + 9
        Note(8) = keyRoot + 10
    Case "Aeorian"
        ReDim Note(8)
        Note(1) = keyRoot
        Note(2) = keyRoot + 2
        Note(3) = keyRoot + 3
        Note(4) = keyRoot + 5
        Note(6) = keyRoot + 7
        Note(7) = keyRoot + 8
        Note(8) = keyRoot + 10
    Case "Dorian"
        ReDim Note(8)
        Note(1) = keyRoot
        Note(2) = keyRoot + 2
        Note(3) = keyRoot + 3
        Note(4) = keyRoot + 5
        Note(6) = keyRoot + 7
        Note(7) = keyRoot + 9
        Note(8) = keyRoot + 10
    Case "Phrygian"
        ReDim Note(8)
        Note(1) = keyRoot
        Note(2) = keyRoot + 1
        Note(3) = keyRoot + 3
        Note(4) = keyRoot + 5
        Note(6) = keyRoot + 7
        Note(7) = keyRoot + 8
        Note(8) = keyRoot + 10
    Case "Locrian"
        ReDim Note(8)
        Note(1) = keyRoot
        Note(2) = keyRoot + 1
        Note(3) = keyRoot + 3
        Note(4) = keyRoot + 5
        Note(6) = keyRoot + 6
        Note(7) = keyRoot + 8
        Note(8) = keyRoot + 10
End Select

'Merge Note
For z = 1 To UBound(Note)
    
    If Note(z) <> "" Then
        If Tmp = "" Then
            Tmp = Tmp & Note(z)
        Else
            Tmp = Tmp & "," & Note(z)
        End If
    End If
    
Next z

CalcKey = Tmp 'Output Temp

Tmp = "" 'Cleanup Temp

End Function

Public Function CalcNote(chdRoot, chdType, chdTention) 'Calculate Chord Note
On Error Resume Next

Dim Note() As String
Dim Tention() As String
Dim Tmp$
Dim i, z As Long

'Array Note
Select Case chdType
    Case ""
        ReDim Note(3)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 7 'Perfect 5
    Case "2"
        ReDim Note(3)
        Note(1) = chdRoot
        Note(2) = chdRoot + 2 'Major 2
        Note(3) = chdRoot + 7 'Perfect 5
    Case "Maj"
        ReDim Note(3)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 7 'Perfect 5
    Case "△"
        ReDim Note(3)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 7 'Perfect 5
    Case "M"
        ReDim Note(3)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 7 'Perfect 5
    Case "sus4"
        ReDim Note(3)
        Note(1) = chdRoot
        Note(2) = chdRoot + 5 'Perfect 4
        Note(3) = chdRoot + 7 'Perfect 5
    Case "-"
        ReDim Note(3)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 7 'Perfect 5
    Case "min"
        ReDim Note(3)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 7 'Perfect 5
    Case "m"
        ReDim Note(3)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 7 'Perfect 5
    Case "+"
        ReDim Note(3)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 8 'Augmented 5
    Case "aug"
        ReDim Note(3)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 8 'Augmented 5
    Case "M7"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 11 'Major 7
    Case "M7sus4"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 5 'Perfect 4
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 11 'Major 7
    Case "Maj7"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 11 'Major 7
    Case "M7b5"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 6 'Perfect 5
        Note(4) = chdRoot + 11 'Major 7
    Case "M7(b5)"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 6 'Perfect 5
        Note(4) = chdRoot + 11 'Major 7
    Case "M7-5"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 6 'Perfect 5
        Note(4) = chdRoot + 11 'Major 7
    Case "M7+5"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 8 'Augmented 5
        Note(4) = chdRoot + 11 'Major 7
    Case "mM7"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Major 3
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 11 'Major 7
    Case "m7"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 10 'Minor 7
    Case "m7+5"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 8 'Augmented 5
        Note(4) = chdRoot + 10 'Minor 7
    Case "dom"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 10 'Minor 7
    Case "7sus4"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 5 'Perfect 4
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 10 'Minor 7
    Case "7"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 10 'Minor 7
    Case "+7"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 8 'Minor 6
        Note(4) = chdRoot + 10 'Minor 7
    Case "7sus4"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 5 'Perfect 4
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 10 'Minor 7
    Case "Dim"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 6 'Diminished 5
    Case "dim"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 6 'Diminished 5
    Case "o"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 6 'Diminished 5
    Case "ø"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 6 'Diminished 5
        Note(4) = chdRoot + 10 'Major 7
    Case "m7b5"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 6 'Diminished 5
        Note(4) = chdRoot + 10 'Major 7
    Case "Dim7"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 6 'Diminished 5
        Note(4) = chdRoot + 9 'Minor 7
    Case "dim7"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 6 'Diminished 5
        Note(4) = chdRoot + 9 'Minor 7
    Case "Dim7sus4"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 5 'Minor 3
        Note(3) = chdRoot + 6 'Diminished 5
        Note(4) = chdRoot + 9 'Minor 7
    Case "dim7sus4"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 5 'Minor 3
        Note(3) = chdRoot + 6 'Diminished 5
        Note(4) = chdRoot + 9 'Minor 7
    Case "M6"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 9 'Major 6
    Case "6"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 9 'Major 6
    Case "m6"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 9 'Major 6
End Select

'Array Tention
If InStr(chdTention, ",") Then

    Tention() = Split(chdTention, ",")

    For i = 0 To UBound(Tention)
        Select Case Tention(i)
            Case "b5"
                Note(2) = Note(2) - 1
            Case "b3"
                Note(2) = Note(1) - 1
            Case "7"
                ReDim Preserve Note(UBound(Note) + 1)
                Note(UBound(Note)) = (chdRoot) + 11
            Case "b9" 'b2
                ReDim Preserve Note(UBound(Note) + 1)
                Note(UBound(Note)) = (chdRoot) + 13
            Case "9" '2
                ReDim Preserve Note(UBound(Note) + 1)
                Note(UBound(Note)) = (chdRoot) + 14
            Case "#9" '#2
                ReDim Preserve Note(UBound(Note) + 1)
                Note(UBound(Note)) = (chdRoot) + 15
            Case "11" ''4
                ReDim Preserve Note(UBound(Note) + 1)
                Note(UBound(Note)) = (chdRoot) + 17
            Case "#11"
                ReDim Preserve Note(UBound(Note) + 1)
                Note(UBound(Note)) = (chdRoot) + 18
            Case "b13" '6
                ReDim Preserve Note(UBound(Note) + 1)
                Note(UBound(Note)) = (chdRoot - 12) + 20
            Case "13"
                ReDim Preserve Note(UBound(Note) + 1)
                Note(UBound(Note)) = (chdRoot - 12) + 21
        End Select
    Next i
    
Else

'Array a Tention
    Select Case chdTention
        Case "b5"
            Note(2) = Note(2) - 1
        Case "b3"
            Note(2) = Note(1) - 1
        Case "7"
            ReDim Preserve Note(UBound(Note) + 1)
            Note(UBound(Note)) = (chdRoot - 12) + 11
        Case "b9" 'b2
            ReDim Preserve Note(UBound(Note) + 1)
            Note(UBound(Note)) = (chdRoot) + 13
        Case "9" '2
            ReDim Preserve Note(UBound(Note) + 1)
            Note(UBound(Note)) = (chdRoot) + 14
        Case "#9" '#2
            ReDim Preserve Note(UBound(Note) + 1)
            Note(UBound(Note)) = (chdRoot) + 15
        Case "11" ''4
            ReDim Preserve Note(UBound(Note) + 1)
            Note(UBound(Note)) = (chdRoot) + 17
        Case "#11"
            ReDim Preserve Note(UBound(Note) + 1)
            Note(UBound(Note)) = (chdRoot) + 18
        Case "b13" '6
            ReDim Preserve Note(UBound(Note) + 1)
            Note(UBound(Note)) = (chdRoot) + 20
        Case "13"
            ReDim Preserve Note(UBound(Note) + 1)
            Note(UBound(Note)) = (chdRoot) + 21
    End Select
End If

'Merge Note
For z = 1 To UBound(Note)
    
    If Note(z) <> "" Then
    
        If Tmp = "" Then
            Tmp = Tmp & Note(z)
        Else
            Tmp = Tmp & "," & Note(z)
        End If
        
    End If
    
Next z

CalcNote = Tmp 'Output Temp

Tmp = "" 'Cleanup Temp

End Function
