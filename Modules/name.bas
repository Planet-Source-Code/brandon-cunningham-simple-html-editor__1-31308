Attribute VB_Name = "Name"
'This code is Intellectual Propriety of Brandon Cunningham.
Public Function namex(File As String) As String
filex = Split(File, "\")
x = UBound(filex)
namex = filex(x)
Rem modified from original, this version returns name w/ extension
End Function
