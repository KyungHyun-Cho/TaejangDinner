Attribute VB_Name = "Module1"
Public Function npe(str As String)
For i = 1 To LenB(str) / 2
npe = npe & "djt" & Replace(Replace(Replace(Replace(Asc(Mid(str, i, 1)), "10", "aih"), "20", "ibd"), "11", "zvg"), "4", "uwe")
Next
End Function


Public Function npd(str As String)
On Error Resume Next
Dim temp As String
Dim Arr() As String
Dim j As Integer
temp = Replace(Replace(Replace(Replace(str, "aih", "10"), "ibd", "20"), "zvg", "11"), "uwe", "4")
Arr = Split(temp, "djt")
For i = 1 To UBound(Arr)
npd = npd & Chr(Arr(i))
Next
End Function
