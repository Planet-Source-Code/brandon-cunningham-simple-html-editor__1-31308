Attribute VB_Name = "Path"
'This code is Intellectual Propriety of Brandon Cunningham
Public Function lpath(fname As String) As String
Dim arrFname As Variant
Dim i As Integer
arrFname = Split(fname, "\")
For i = LBound(arrFname) To UBound(arrFname) - 1
If i = 0 Then
lpath = arrFname(i)
Else
lpath = lpath & "\" & arrFname(i)
End If
Next i
End Function
