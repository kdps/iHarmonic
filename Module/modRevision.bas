Attribute VB_Name = "modRevision"
Option Explicit

Public Function SplitNote(chdRoot, chdType, chdTention) 'Split Note by Array
On Error Resume Next
Dim Note() As String
Dim i As Long
Note() = Split(CalcNote(chdRoot, chdType, chdTention), ",") 'Split Note
End Function

Public Function RevisionChord(Chord As String, i As Long)
On Error Resume Next
Dim Chd, txtRoot, txtKind, txtInvention, txtInvention2, strInput, strTention As String
Dim n, l As Long

Chord = Replace(Chord, " ", "") 'Remove spaces
Note() = Split(Chord, "/") 'Separated by a special character "/"

If Mid$(Note(0), 2, 1) = "#" Or Mid$(Note(0), 2, 1) = "b" Then
    txtRoot = Mid$(Note(0), 1, 2)
    txtKind = Mid$(Note(0), 3, Len(Note(0)) - 2)
Else
    txtRoot = Mid$(Note(0), 1, 1)
    txtKind = Mid$(Note(0), 2, Len(Note(0)) - 1)
End If

If UBound(Note) > 0 Then
    txtInvention = Note(1)
    If InStr(txtInvention, "#") Or InStr(txtInvention, "b") Then
        txtInvention = Mid$(txtInvention, 1, 2)
    Else
        txtInvention = Mid$(txtInvention, 1, 1)
    End If
End If

If UBound(Note) > 1 Then
    txtInvention2 = Note(2)
End If

If Mid$(txtKind, 2, 1) = "#" Or Mid$(txtKind, 2, 1) = "b" Then
    strInput = Mid$(txtKind, 3, Len(txtKind) - 2)
Else
    strInput = Mid$(txtKind, 2, Len(txtKind) - 1)
End If

If Not InStr(strInput, "(") = 0 Then
    n = InStr(strInput, "(") + 1
    l = Len(Mid$(strInput, n)) + 1
    strTention = Mid$(strInput, n, Len(strInput) - (n))
    If UBound(Note) > 0 Then
        mdi_frmMain.ActiveForm.lstChord.ListItems(i) = txtRoot & txtKind & "/" & txtInvention
        mdi_frmMain.ActiveForm.lstChord.ListItems(i).ListSubItems(1).Text = TexttoNote(txtRoot) & "/" & Mid$(strInput, 3, Len(strInput) - (2 + l)) & "/" & strTention & "/" & TexttoNote(txtInvention) & "/" & txtInvention2
    ElseIf UBound(Note) > 1 Then
        mdi_frmMain.ActiveForm.lstChord.ListItems(i).Text = txtRoot & txtKind & "/" & txtInvention
        mdi_frmMain.ActiveForm.lstChord.ListItems(i).ListSubItems(1).Text = TexttoNote(txtRoot) & "/" & Mid$(strInput, 3, Len(strInput) - (2 + l)) & "/" & strTention & "/" & TexttoNote(txtInvention) & "/" & txtInvention2
    End If
Else
    If UBound(Note) > 0 Then
        mdi_frmMain.ActiveForm.lstChord.ListItems(i) = txtRoot & txtKind & "/" & txtInvention
        mdi_frmMain.ActiveForm.lstChord.ListItems(i).ListSubItems(1).Text = TexttoNote(txtRoot) & "/" & txtKind & "/" & "0" & "/" & TexttoNote(txtInvention)
    ElseIf UBound(Note) > 1 Then
        mdi_frmMain.ActiveForm.lstChord.ListItems(i).Text = txtRoot & txtKind & "/" & txtInvention
        mdi_frmMain.ActiveForm.lstChord.ListItems(i).ListSubItems(1).Text = TexttoNote(txtRoot) & "/" & txtKind & "/" & "0" & "/" & TexttoNote(txtInvention) & "/" & txtInvention2
    End If
End If

End Function

Public Function ResetPos(x As Long, Listtext1 As String, Listtext2 As String)
On Error Resume Next
Dim i As Long

If mdi_frmMain.ActiveForm.lstChord.ListItems.Count = 0 Then
    mdi_frmMain.ActiveForm.lstChord.ListItems.Add , , Listtext1
    mdi_frmMain.ActiveForm.lstChord.ListItems(mdi_frmMain.ActiveForm.lstChord.ListItems.Count).ListSubItems.Add.Text = Listtext2
    mdi_frmMain.ActiveForm.lstChord.ListItems(mdi_frmMain.ActiveForm.lstChord.ListItems.Count).ForeColor = vbBlue
Else
    x = mdi_frmMain.ActiveForm.lstChord.ListItems.Count - x
    mdi_frmMain.ActiveForm.lstChord.ListItems.Add , , ""
    mdi_frmMain.ActiveForm.lstChord.ListItems(mdi_frmMain.ActiveForm.lstChord.ListItems.Count).Text = mdi_frmMain.ActiveForm.lstChord.ListItems(mdi_frmMain.ActiveForm.lstChord.ListItems.Count - 1).Text
    mdi_frmMain.ActiveForm.lstChord.ListItems(mdi_frmMain.ActiveForm.lstChord.ListItems.Count).ListSubItems.Add.Text = mdi_frmMain.ActiveForm.lstChord.ListItems(mdi_frmMain.ActiveForm.lstChord.ListItems.Count - 1).ListSubItems(1).Text
    For i = 2 To x - 1
        mdi_frmMain.ActiveForm.lstChord.ListItems(mdi_frmMain.ActiveForm.lstChord.ListItems.Count - (i - 1)).Text = mdi_frmMain.ActiveForm.lstChord.ListItems(mdi_frmMain.ActiveForm.lstChord.ListItems.Count - i).Text
        mdi_frmMain.ActiveForm.lstChord.ListItems(mdi_frmMain.ActiveForm.lstChord.ListItems.Count - (i - 1)).ListSubItems(1).Text = mdi_frmMain.ActiveForm.lstChord.ListItems(mdi_frmMain.ActiveForm.lstChord.ListItems.Count - i).ListSubItems(1).Text
    Next i
    mdi_frmMain.ActiveForm.lstChord.ListItems(mdi_frmMain.ActiveForm.lstChord.ListItems.Count - (x - 1)).Text = Listtext1
    mdi_frmMain.ActiveForm.lstChord.ListItems(mdi_frmMain.ActiveForm.lstChord.ListItems.Count - (x - 1)).ListSubItems(1).Text = Listtext2
    mdi_frmMain.ActiveForm.lstChord.ListItems(mdi_frmMain.ActiveForm.lstChord.ListItems.Count - (x - 1)).ForeColor = vbMagenta
End If

End Function

Public Function AddChord(strChord As String)
On Error Resume Next
Dim strTention As String
Dim Chd As String
Dim Rtxt As Integer
Dim n, i As Integer

i = mdi_frmMain.ActiveForm.lstChord.ListItems.Count


If Not InStr(strChord, "(") = 0 Then
    n = InStr(strChord, "(") + 1
    l = Len(Mid$(strChord, n)) + 1
    strTention = Mid$(strChord, n, Len(strChord) - (n))
Else
    strTention = ""
End If

If Mid$(strChord, 2, 1) = "#" Or Mid$(strChord, 2, 1) = "b" Then
    If strTention <> "" Then
        mdi_frmMain.ActiveForm.lstChord.ListItems.Add , , Mid$(strChord, 1, 2) & Mid$(strChord, 3, Len(strChord) - 2) & "/" & Mid$(strChord, 1, 2)
        mdi_frmMain.ActiveForm.lstChord.ListItems(mdi_frmMain.ActiveForm.lstChord.ListItems.Count).ListSubItems.Add.Text = TexttoNote(Mid$(strChord, 1, 2)) & "/" & Mid$(strChord, 3, Len(strChord) - (2 + l)) & "/" & strTention & "/" & TexttoNote(Mid$(strChord, 1, 2))
    Else
        mdi_frmMain.ActiveForm.lstChord.ListItems.Add , , Mid$(strChord, 1, 2) & Mid$(strChord, 3, Len(strChord) - 2) & "/" & Mid$(strChord, 1, 2)
        mdi_frmMain.ActiveForm.lstChord.ListItems(mdi_frmMain.ActiveForm.lstChord.ListItems.Count).ListSubItems.Add.Text = TexttoNote(Mid$(strChord, 1, 2)) & "/" & Mid$(strChord, 3, Len(strChord) - 2) & "/" & "0" & "/" & TexttoNote(Mid$(strChord, 1, 2))
    End If
Else
    If strTention <> "" Then
        mdi_frmMain.ActiveForm.lstChord.ListItems.Add , , Mid$(strChord, 1, 1) & Mid$(strChord, 2, Len(strChord) - 1) & "/" & Mid$(strChord, 1, 1)
        mdi_frmMain.ActiveForm.lstChord.ListItems(mdi_frmMain.ActiveForm.lstChord.ListItems.Count).ListSubItems.Add.Text = TexttoNote(Mid$(strChord, 1, 1)) & "/" & Mid$(strChord, 2, Len(strChord) - (1 + l)) & "/" & strTention & "/" & TexttoNote(Mid$(strChord, 1, 1))
    Else
        mdi_frmMain.ActiveForm.lstChord.ListItems.Add , , Mid$(strChord, 1, 1) & Mid$(strChord, 2, Len(strChord) - 1) & "/" & Mid$(strChord, 1, 1)
        mdi_frmMain.ActiveForm.lstChord.ListItems(mdi_frmMain.ActiveForm.lstChord.ListItems.Count).ListSubItems.Add.Text = TexttoNote(Mid$(strChord, 1, 1)) & "/" & Mid$(strChord, 2, Len(strChord) - 1) & "/" & "0" & "/" & TexttoNote(Mid$(strChord, 1, 1))
    End If
End If

End Function
