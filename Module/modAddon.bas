Attribute VB_Name = "modAddon"

Public Function PreviewChord(Actives As Boolean, strRoot As String, strType As String, strTention As String, strInv As String)
Dim Note() As String
Dim i As Long

Note() = Split(CalcNote((12 * 3) + TexttoNote(strRoot), strType, Mid(strTention, 2)), ",") 'Split Note

Select Case strInv
Case 1
    Note(0) = Note(0) + 12
    Note(1) = Note(1)
    Note(2) = Note(2)
Case 2
    Note(0) = Note(0) + 12
    Note(1) = Note(1) + 12
    Note(2) = Note(2)
Case 3
    Note(0) = Note(0) + 12
    Note(1) = Note(1) + 12
    Note(2) = Note(2) + 12
    Note(3) = Note(3)
End Select

PlayScale UBound(Note), True, 70, False

End Function
