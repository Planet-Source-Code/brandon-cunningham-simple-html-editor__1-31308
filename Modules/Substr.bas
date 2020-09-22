Attribute VB_Name = "Substr"
'This code is Intellectual Propriety of Brandon Cunningham.
Public Function xsubstr(string1 As String, pos As Integer, countback As Integer) As String
LeftStr = Left(string1, pos)
xsubstr = Right(LeftStr, countback)
End Function
