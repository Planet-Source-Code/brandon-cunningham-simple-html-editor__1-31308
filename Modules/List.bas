Attribute VB_Name = "List"
'This code is Intellectual Propriety of Brandon Cunningham.
Public Function ListAssign(List_1 As Object, list_2 As Object)
Dim x As Variant
For x = 0 To List_1.ListCount - 1
list_2.AddItem List_1.List(x)
Next x
End Function

Public Function ListCopy(List_1 As Object)
Dim Count As Long
Dim Copy As String
For Count = 0 To List_1.ListCount - 1
If Count = 0 Then
Copy = List_1.List(Count)
Else
Copy = Copy & Chr(13) + Chr(10) & List_1.List(Count)
End If
Next Count
Clipboard.Clear
Clipboard.SetText Copy
End Function
Public Function ListLoadfile(List_1 As Object, file_1 As String)
On Error GoTo gracefulexit:
List_1.Clear
Dim lstInput As String
Open file_1 For Input As #1
While Not EOF(1)
Input #1, lstInput$
List_1.AddItem lstInput$
Wend
Close #1
gracefulexit:
Exit Function
End Function
Public Function ListSaveFile(List_1 As Object, file_1 As String)
Dim lngSave As Integer
lngSave = 0
    Open file_1 For Output As #1
        For lngSave = 0 To List_1.ListCount
            Print #1, List_1.List(lngSave)
        Next lngSave
    Close #1
End Function

Public Function ListText(List_1 As Object) As String
Dim Count As Long
For Count = 0 To List_1.ListCount - 1
If Count = 0 Then
ListText = List_1.List(Count)
Else
ListText = ListText & Chr(13) + Chr(10) & List_1.List(Count)
End If
Next Count
End Function
