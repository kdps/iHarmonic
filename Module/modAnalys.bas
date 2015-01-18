Attribute VB_Name = "modAnalys"
Option Explicit
Public Sub CalcAll(posRoot, posChd)

Dim txtUnknown As String
Dim txtDS As String
Dim txtTSDm As String
Dim txtPassDim As String
Dim txtSecondary As String
Dim txtSecondarym As String
Dim txtSubDominant7th As String
Dim txtSubDominant7thm As String
Dim txtSubDominantMinor As String
Dim txtSubDominantMinor2 As String
Dim txtTonic As String
Dim txtTonicm As String
Dim txtSubDominant As String
Dim txtSubDominantm As String
Dim txtDominant As String
Dim txtDominantm As String
Dim txtPass As String

txtPass = "Passing Tone"
txtDS = "Dominant/Sub Dominant"
txtTSDm = "Tonic/Sub Dominant Minor"
txtPassDim = "Passing Diminished"
txtSecondary = "Secondary Dominant"
txtSubDominant7th = "Substitute Dominnt 7th"
txtSubDominantMinor = "Secondary Dominant Minor"
txtTonic = "Tonic"
txtSubDominant = "Sub Dominant"
txtDominant = "Dominant"
txtUnknown = "Unknown"
txtTonicm = "Tonic Minor"
txtSubDominantm = "Sub Dominant(Minor)"
txtDominantm = "Dominant (Minor)"
txtSecondarym = "Secondary Dominant(Minor)"
txtSubDominant7thm = "Sub Dominant (Minor)"
txtSubDominantMinor2 = "Sub Dominant Minor"

If mMScale = "Major" Or Played = True Then
    If (posRoot) - Key & posChd = "07" And Related = True Then 'Related II-V
        chdClassic = "I7, I7/IV7"
        chdFunction = "V7" & "(" & NotetoText((posRoot - 12) + 5) & "), " & txtSecondary
        Related = False
    ElseIf (posRoot) - Key & posChd = "57" And Related = True Then
        chdClassic = "IV7 , IV7/VII7"
        chdFunction = "V7" & "(" & NotetoText((posRoot - 12) + 10) & "), " & txtSubDominant7th
        Related = False
    ElseIf ((posRoot) - Key & posChd = "107" Or (posRoot) - Key & posChd = "10") And Related = True Then
        chdClassic = "bVII7"
        chdFunction = "V7" & "(" & NotetoText((posRoot - 12) + 3) & "), " & txtSubDominantMinor
        Related = False
    ElseIf ((posRoot) - Key & posChd = "37" Or (posRoot) - Key & posChd = "3") And Related = True Then
        chdClassic = "bIII7, bIII7/V7"
        chdFunction = "V7" & "(" & NotetoText((posRoot - 12) + 8) & "), " & txtSubDominant7th
        Related = False
    ElseIf ((posRoot) - Key & posChd = "87" Or (posRoot) - Key & posChd = "8") And Related = True Then
        chdClassic = "bVI7"
        chdFunction = "V7" & "(" & NotetoText((posRoot - 12) + 1) & ")"
        Related = False
    ElseIf ((posRoot) - Key & posChd = "17" Or (posRoot) - Key & posChd = "1") And Related = True Then
        chdClassic = "#I7"
        chdFunction = "V7" & "(" & NotetoText((posRoot - 12) + 6) & ")"
        Related = False
    ElseIf ((posRoot) - Key & posChd = "67" Or (posRoot) - Key & posChd = "6") And Related = True Then
        chdClassic = "#IV7"
        chdFunction = "V7" & "(" & NotetoText((posRoot - 12) + 11) & ")"
        Related = False
    ElseIf ((posRoot) - Key & posChd = "117" Or (posRoot) - Key & posChd = "11") And Related = True Then
        chdClassic = "VII7, VII7/IIIm"
        chdFunction = "V7" & "(" & NotetoText((posRoot - 12) + 4) & "), " & txtSecondary
        Related = False
    ElseIf ((posRoot) - Key & posChd = "47" Or (posRoot) - Key & posChd = "4") And Related = True Then
        chdClassic = "III7, III7/VIm"
        chdFunction = "V7" & "(" & NotetoText((posRoot - 12) + 9) & "), " & txtSecondary
        Related = False
    ElseIf ((posRoot) - Key & posChd = "97" Or (posRoot) - Key & posChd = "9") And Related = True Then
        chdClassic = "VI7, IIm"
        chdFunction = "V7" & "(" & NotetoText((posRoot - 12) + 2) & "), " & txtSecondary
        Related = False
    ElseIf ((posRoot) - Key & posChd = "27" Or (posRoot) - Key & posChd = "2") And Related = True Then
        chdClassic = "II7, II7/V"
        chdFunction = "V7" & "(" & NotetoText((posRoot - 12) + 7) & "), " & txtSecondary
        Related = False
    ElseIf ((posRoot) - Key & posChd = "7min" Or (posRoot) - Key & posChd = "7m") Then 'Check Related II-V
        chdClassic = "Vmin"
        chdFunction = "IIm7" & "(" & NotetoText((posRoot - 7) + 5) & ")"
        Related = True
    ElseIf (posRoot) - Key & posChd = "7m7" Then
        chdClassic = "Vm7"
        chdFunction = "IIm7" & "(" & NotetoText((posRoot - 7) + 5) & ")"
        Related = True
    ElseIf (posRoot) - Key & posChd = "0min" Or (posRoot) - Key & posChd = "0m" Then
        chdClassic = "Imin"
        chdFunction = "IIm7" & "(" & NotetoText((posRoot - 7) + 10) & ")"
        Related = True
    ElseIf (posRoot) - Key & posChd = "0m7" Then
        chdClassic = "Im7"
        chdFunction = "IIm7" & "(" & NotetoText((posRoot - 7) + 10) & ")"
        Related = True
    ElseIf (posRoot) - Key & posChd = "5min" Or (posRoot) - Key & posChd = "5m" Then
        chdClassic = "IVm"
        chdFunction = "IIm7" & "(" & NotetoText((posRoot - 7) + 3) & ") / " & txtSubDominantMinor
        Related = True
    ElseIf (posRoot) - Key & posChd = "5m7" Then
        chdClassic = "IVm7"
        chdFunction = "IIm7" & "(" & NotetoText((posRoot - 7) + 3) & ") / " & txtSubDominantMinor
        Related = True
    ElseIf (posRoot) - Key & posChd = "10min" Or (posRoot) - Key & posChd = "10m" Then
        chdClassic = "bVIIm"
        chdFunction = "IIm7" & "(" & NotetoText((posRoot - 7) + 8) & ")"
        Related = True
    ElseIf (posRoot) - Key & posChd = "10m7" Then
        chdClassic = "bVIIm7"
        chdFunction = "IIm7" & "(" & NotetoText((posRoot - 7) + 8) & ")"
        Related = True
    ElseIf (posRoot) - Key & posChd = "3min" Or (posRoot) - Key & posChd = "3m" Then
        chdClassic = "bIIIm"
        chdFunction = "IIm7" & "(" & NotetoText((posRoot - 7) + 1) & ")"
        Related = True
    ElseIf (posRoot) - Key & posChd = "3m7" Then
        chdClassic = "bIIIm7"
        chdFunction = "IIm7" & "(" & NotetoText((posRoot - 7) + 1) & ")"
        Related = True
    ElseIf (posRoot) - Key & posChd = "8min" Or (posRoot) - Key & posChd = "8m" Then
        chdClassic = "bVIm"
        chdFunction = "IIm7" & "(" & NotetoText((posRoot - 7) + 6) & ")"
        Related = True
    ElseIf (posRoot) - Key & posChd = "8m7" Then
        chdClassic = "bVIm7"
        chdFunction = "IIm7" & "(" & NotetoText((posRoot - 7) + 6) & ")"
        Related = True
    ElseIf (posRoot) - Key & posChd = "1min" Or (posRoot) - Key & posChd = "1m" Then
        chdClassic = "#Im"
        chdFunction = "IIm7" & "(" & NotetoText((posRoot - 7) + 11) & ")"
        Related = True
    ElseIf (posRoot) - Key & posChd = "1m7" Then
        chdClassic = "#Im7"
        chdFunction = "IIm7" & "(" & NotetoText((posRoot - 7) + 11) & ")"
        Related = True
    ElseIf (posRoot) - Key & posChd = "6min" Or (posRoot) - Key & posChd = "6m" Then
        chdClassic = "#IVm(II)"
        chdFunction = "IIm7" & "(" & NotetoText((posRoot - 7) + 4) & ") / " & txtTonic
        Related = True
    ElseIf (posRoot) - Key & posChd = "6m7" Then
        chdClassic = "#IVm7(II7)"
        chdFunction = "IIm7" & "(" & NotetoText((posRoot - 7) + 4) & ") / " & txtTonic
        Related = True
    ElseIf (posRoot) - Key & posChd = "11min" Or (posRoot) - Key & posChd = "11m" Then
        chdClassic = "VIIm"
        chdFunction = "IIm7" & "(" & NotetoText((posRoot - 7) + 9) & ")"
        Related = True
    ElseIf (posRoot) - Key & posChd = "11m7" Then
        chdClassic = "VIIm7"
        chdFunction = "IIm7" & "(" & NotetoText((posRoot - 7) + 9) & ")"
        Related = True
    ElseIf (posRoot) - Key & posChd = "4min" Or (posRoot) - Key & posChd = "4m" Then
        chdClassic = "IIIm7"
        chdFunction = "IIm7" & "(" & NotetoText((posRoot - 7) + 2) & "), " & txtTonic
        Related = True
    ElseIf (posRoot) - Key & posChd = "4m7" Then
        chdClassic = "IIIm7"
        chdFunction = "IIm7" & "(" & NotetoText((posRoot - 7) + 2) & "), " & txtTonic
        Related = True
    ElseIf (posRoot) - Key & posChd = "9min" Or (posRoot) - Key & posChd = "9m" Then
        chdClassic = "VIm"
        chdFunction = "IIm7" & "(" & NotetoText((posRoot - 7) + 7) & "), " & txtTonic
        Related = True
    ElseIf (posRoot) - Key & posChd = "9m7" Then
        chdClassic = "VIm7"
        chdFunction = "IIm7" & "(" & NotetoText((posRoot - 7) + 7) & "), " & txtTonic
        Related = True
    ElseIf (posRoot) - Key & posChd = "0Maj" Or (posRoot) - Key & posChd = "0" Then
        chdClassic = "I"
        chdFunction = txtTonic
    ElseIf (posRoot) - Key & posChd = "2min" Or (posRoot) - Key & posChd = "2m" Then
        chdClassic = "IIm"
        chdFunction = txtSubDominant
    ElseIf (posRoot) - Key & posChd = "4min" Or (posRoot) - Key & posChd = "4m" Then
        chdClassic = "IIIm"
        chdFunction = txtTonic
    ElseIf (posRoot) - Key & posChd = "5Maj" Or (posRoot) - Key & posChd = "5" Then
        chdClassic = "IV"
        chdFunction = txtSubDominant
    ElseIf (posRoot) - Key & posChd = "7Maj" Or (posRoot) - Key & posChd = "7" Then
        chdClassic = "V"
        chdFunction = txtDominant
    ElseIf (posRoot) - Key & posChd = "9min" Or (posRoot) - Key & posChd = "9m" Then
        chdClassic = "VIm"
        chdFunction = txtTonic
    ElseIf (posRoot) - Key & posChd = "11m7b5" Or (posRoot) - Key & posChd = "11dim" Then
        chdClassic = "VIIm7b5"
        chdFunction = txtDS
    ElseIf (posRoot) - Key & posChd = "0M7" Then
        chdClassic = "IM7"
        chdFunction = txtTonic
    ElseIf (posRoot) - Key & posChd = "2m7" Or (posRoot) - Key & posChd = "2m" Then
        chdClassic = "IIm7"
        chdFunction = txtSubDominant
    ElseIf (posRoot) - Key & posChd = "4m7" Or (posRoot) - Key & posChd = "4m" Then
        chdClassic = "IIIm7"
        chdFunction = txtTonic
    ElseIf (posRoot) - Key & posChd = "5M7" Or (posRoot) - Key & posChd = "5" Then
        chdClassic = "IVM7"
        chdFunction = txtSubDominant
    ElseIf (posRoot) - Key & posChd = "5m" Then
        chdClassic = "IVm"
        chdFunction = txtSubDominantMinor2
    ElseIf (posRoot) - Key & posChd = "5m7" Then
        chdClassic = "IVm7"
        chdFunction = txtSubDominantMinor2
    ElseIf (posRoot) - Key & posChd = "77" Or (posRoot) - Key & posChd = "7" Then
        chdClassic = "V7"
        chdFunction = txtDominant
    ElseIf (posRoot) - Key & posChd = "77sus4" Then
        chdClassic = "V7sus4"
        chdFunction = txtSubDominant
    ElseIf (posRoot) - Key & posChd = "7sus4" Then
        chdClassic = "Vsus4"
        chdFunction = txtSubDominant
    ElseIf (posRoot) - Key & posChd = "9m7" Or (posRoot) - Key & posChd = "9m" Then
        chdClassic = "VIm7"
        chdFunction = txtTonic
    ElseIf (posRoot) - Key & posChd = "11Dim" Then
        chdClassic = "VIIDim"
        chdFunction = txtDS
    ElseIf (posRoot) - Key & posChd = "2Maj" Or (posRoot) - Key & posChd = "2" Then
        chdClassic = "II/V7"
        chdFunction = txtSecondary
    ElseIf (posRoot) - Key & posChd = "4Maj" Or (posRoot) - Key & posChd = "4" Then
        chdClassic = "III/VIm"
        chdFunction = txtDS
    ElseIf (posRoot) - Key & posChd = "9Maj" Or (posRoot) - Key & posChd = "9" Then
        chdClassic = "VI/IIm"
        chdFunction = txtSecondary
    ElseIf (posRoot) - Key & posChd = "11Maj" Or (posRoot) - Key & posChd = "11" Then
        chdClassic = "VII/IIIm"
        chdFunction = txtSecondary
    ElseIf (posRoot) - Key & posChd = "07" Or (posRoot) - Key & posChd = "07" Then
        chdClassic = "I7/IV"
        chdFunction = txtSecondary
    ElseIf (posRoot) - Key & posChd = "27" Or (posRoot) - Key & posChd = "2" Then
        chdClassic = "II7/V7"
        chdFunction = txtSecondary
    ElseIf (posRoot) - Key & posChd = "47" Or (posRoot) - Key & posChd = "4" Then
        chdClassic = "III7/VIm"
        chdFunction = txtSecondary
    ElseIf (posRoot) - Key & posChd = "97" Or (posRoot) - Key & posChd = "9" Then
        chdClassic = "VI7/IIm"
        chdFunction = txtSecondary
    ElseIf (posRoot) - Key & posChd = "117" Or (posRoot) - Key & posChd = "11" Then
        chdClassic = "VII7/IIIm"
        chdFunction = txtSecondary
    ElseIf (posRoot) - Key & posChd = "6m7b5" Or (posRoot) - Key & posChd = "6Dim" Then
        chdClassic = "#IVm7b5"
        chdFunction = txtTonic
    ElseIf (posRoot) - Key & posChd = "5m6" Then
        chdClassic = "IVm6"
        chdFunction = txtSubDominantMinor2
    ElseIf (posRoot) - Key & posChd = "5m7" Or (posRoot) - Key & posChd = "5m" Then
        chdClassic = "IVm7"
        chdFunction = txtSubDominantMinor2
    ElseIf (posRoot) - Key & posChd = "86" Then
        chdClassic = "bVI6"
        chdFunction = txtSubDominantMinor2
    ElseIf (posRoot) - Key & posChd = "8M7" Or (posRoot) - Key & posChd = "8" Then
        chdClassic = "bVIM7"
        chdFunction = txtSubDominantMinor2
    ElseIf (posRoot) - Key & posChd = "1M7" Then
        chdClassic = "bIIM7"
        chdFunction = txtSubDominantMinor2
    ElseIf (posRoot) - Key & posChd = "2m7b5" Or (posRoot) - Key & posChd = "2Dim" Then
        chdClassic = "IIm7b5"
        chdFunction = txtSubDominantMinor2
    ElseIf (posRoot) - Key & posChd = "11Maj" Or (posRoot) - Key & posChd = "11" Then
        chdClassic = "bVII"
        chdFunction = txtSubDominantMinor2
    ElseIf (posRoot) - Key & posChd = "107" Or (posRoot) - Key & posChd = "10" Then
        chdClassic = "bVII7"
        chdFunction = txtSubDominantMinor2
    ElseIf (posRoot) - Key & posChd = "1Maj" Or (posRoot) - Key & posChd = "1" Then
        chdClassic = "bII"
        chdFunction = txtSubDominant7th
    ElseIf (posRoot) - Key & posChd = "17" Then
        chdClassic = "bII7"
        chdFunction = txtSubDominant7th
    ElseIf (posRoot) - Key & posChd = "57" Then
        chdClassic = "IV7/VII7/IIIm7"
        chdFunction = txtSubDominant7th
    ElseIf (posRoot) - Key & posChd = "4Maj" Or (posRoot) - Key & posChd = "4" Then
        chdClassic = "bIII/V7/I"
        chdFunction = txtSubDominant7th
    ElseIf (posRoot) - Key & posChd = "47" Then
        chdClassic = "bIII7/V7/I"
        chdFunction = txtSubDominant7th
    ElseIf (posRoot) - Key & posChd = "06" Then
        chdClassic = "I6 / VIm7/I"
        chdFunction = txtTonic
    ElseIf (posRoot) - Key & posChd = "0Dim" Then
        chdClassic = "V7 (Deck) I6"
        chdFunction = "Dim(" & txtPass & ") Chord"
    ElseIf (posRoot) - Key & posChd = "0m7b5" Then
        chdClassic = "Im7b5"
        chdFunction = txtSubDominant
    ElseIf (posRoot) - Key & posChd = "11Dim" Then
        chdClassic = "VIIDim - C"
        chdFunction = txtPassDim
    ElseIf (posRoot) - Key & posChd = "1Dim" Then
        chdClassic = "#IDim - IIm"
        chdFunction = txtPassDim
    ElseIf (posRoot) - Key & posChd = "3Dim" Then
        chdClassic = "#IIDim - IIIm"
        chdFunction = txtPassDim
    ElseIf (posRoot) - Key & posChd = "4Dim" Then
        chdClassic = "IIIDim - IV"
        chdFunction = txtPassDim
    ElseIf (posRoot) - Key & posChd = "6Dim" Then
        chdClassic = "#IVDim - V"
        chdFunction = txtPassDim
    ElseIf (posRoot) - Key & posChd = "8Dim" Then
        chdClassic = "#VDim - VIm"
        chdFunction = txtPassDim
    ElseIf (posRoot) - Key & posChd = "0+" Then
        chdClassic = "Caug/F"
        chdFunction = txtPass
    ElseIf (posRoot) - Key & posChd = "4+" Then
        chdClassic = "Eaug/F"
        chdFunction = txtPass
    ElseIf (posRoot) - Key & posChd = "7+" Then
        chdClassic = "Gaug/F"
        chdFunction = txtPass
    ElseIf (posRoot) - Key & posChd = "4" Then
        chdClassic = "E/F"
        chdFunction = txtPass
    Else
        chdClassic = NotetoRoma(posRoot) & posChd
        chdFunction = txtUnknown
    End If
ElseIf mMScale = "Minor" Or Played = True Then
    If (posRoot) - Key & posChd = "0min" Or (posRoot) - Key & posChd = "0m" Then
        chdClassic = "Im"
        chdFunction = txtTonicm
    ElseIf (posRoot) - Key & posChd = "0m7" Then
        chdClassic = "Im7"
        chdFunction = txtTonicm
    ElseIf (posRoot) - Key & posChd = "0m6" Then
        chdClassic = "Im6"
        chdFunction = txtTonicm
    ElseIf (posRoot) - Key & posChd = "9m7b5" Or (posRoot) - Key & posChd = "9Dim" Then
        chdClassic = "VIm7b5"
        chdFunction = txtTonicm
    ElseIf (posRoot) - Key & posChd = "8Maj" Or (posRoot) - Key & posChd = "8" Then
        chdClassic = "bVIM"
        chdFunction = txtTSDm
    ElseIf (posRoot) - Key & posChd = "8M7" Then
        chdClassic = "bVIM7"
        chdFunction = txtTSDm
    ElseIf (posRoot) - Key & posChd = "3M" Or (posRoot) - Key & posChd = "3" Then
        chdClassic = "bIIIM"
        chdFunction = txtTonicm
    ElseIf (posRoot) - Key & posChd = "3M7" Then
        chdClassic = "bIIIM7"
        chdFunction = txtTonicm
    ElseIf (posRoot) - Key & posChd = "2m7b5" Or (posRoot) - Key & posChd = "2Dim" Then
        chdClassic = "IIm7b5"
        chdFunction = txtSubDominantMinor2
    ElseIf (posRoot) - Key & posChd = "2min" Or (posRoot) - Key & posChd = "2m" Then
        chdClassic = "IIm"
        chdFunction = txtSubDominantMinor2
    ElseIf (posRoot) - Key & posChd = "2m7" Then
        chdClassic = "IIm7"
        chdFunction = txtSubDominantMinor2
    ElseIf (posRoot) - Key & posChd = "5m6" Then
        chdClassic = "IVm6"
        chdFunction = txtSubDominantMinor2
    ElseIf (posRoot) - Key & posChd = "9min" Or (posRoot) - Key & posChd = "9m" Then
        chdClassic = "IVm6/VI"
        chdFunction = txtSubDominantMinor2
    ElseIf (posRoot) - Key & posChd = "5min" Or (posRoot) - Key & posChd = "5m" Then
        chdClassic = "IVm"
        chdFunction = txtSubDominantMinor2
    ElseIf (posRoot) - Key & posChd = "5m7" Then
        chdClassic = "IVm7"
        chdFunction = txtSubDominantMinor2
    ElseIf (posRoot) - Key & posChd = "7Maj" Then
        chdClassic = "V"
        chdFunction = txtDominantm
    ElseIf (posRoot) - Key & posChd = "77" Or (posRoot) - Key & posChd = "7" Then
        chdClassic = "V7"
        chdFunction = txtDominantm
    ElseIf (posRoot) - Key & posChd = "10Dim" Then
        chdClassic = "bVIIDim"
        chdFunction = txtDominantm
    ElseIf (posRoot) - Key & posChd = "10m7b5" Then
        chdClassic = "VIIm7b5"
        chdFunction = txtDominantm
    ElseIf (posRoot) - Key & posChd = "9Maj" Or (posRoot) - Key & posChd = "9" Then
        chdClassic = "bVII"
        chdFunction = txtDominantm
    ElseIf (posRoot) - Key & posChd = "97" Or (posRoot) - Key & posChd = "9" Then
        chdClassic = "bVII7"
        chdFunction = txtDominantm
    ElseIf (posRoot) - Key & posChd = "7min" Then
        chdClassic = "Vm"
        chdFunction = txtDominantm
    ElseIf (posRoot) - Key & posChd = "7m7" Then
        chdClassic = "Vm7"
        chdFunction = txtDominantm
    ElseIf (posRoot) - Key & posChd = "07" Then
        chdClassic = "I7/IVm7"
        chdFunction = txtSecondarym
    ElseIf (posRoot) - Key & posChd = "2Maj" Or (posRoot) - Key & posChd = "2" Then
        chdClassic = "II/V7"
        chdFunction = txtSecondarym
    ElseIf (posRoot) - Key & posChd = "27" Or (posRoot) - Key & posChd = "2" Then
        chdClassic = "II7/V7"
        chdFunction = txtSecondarym
    ElseIf (posRoot) - Key & posChd = "37" Or (posRoot) - Key & posChd = "3" Then
        chdClassic = "bIII7/bVIM7"
        chdFunction = txtSecondarym
    ElseIf (posRoot) - Key & posChd = "17" Or (posRoot) - Key & posChd = "1" Then
        chdClassic = "bII7/Im"
        chdFunction = txtSubDominant7thm
    ElseIf (posRoot) - Key & posChd = "67" Or (posRoot) - Key & posChd = "6" Then
        chdClassic = "bV7/IVm"
        chdFunction = txtSubDominant7thm
    ElseIf (posRoot) - Key & posChd = "87" Or (posRoot) - Key & posChd = "8" Then
        chdClassic = "bVI7/V7"
        chdFunction = txtSubDominant7thm
    ElseIf (posRoot) - Key & posChd = "97" Or (posRoot) - Key & posChd = "9" Then
        chdClassic = "VI7/bVIM7"
        chdFunction = txtSubDominant7thm
    ElseIf (posRoot) - Key & posChd = "1Maj" Or (posRoot) - Key & posChd = "1" Then
        chdClassic = "bIIM/bVI7"
        chdFunction = txtSubDominant7thm
    ElseIf (posRoot) - Key & posChd = "1M7" Or (posRoot) - Key & posChd = "1" Then
        chdClassic = "bIIM7/bVI7"
        chdFunction = txtSubDominant7thm
    Else
        chdClassic = NotetoRoma(posRoot) & posChd
        chdFunction = txtUnknown
    End If
End If
End Sub

