Attribute VB_Name = "onDragDrop"
'This code is Intellectual Propriety of Brandon Cunningham.
Public Function onDragDropx(filex As String)
Dim mainstring, x As String
Dim i As Integer
If filex = "" Then
Exit Function
Else
End If
Call ListLoadfile(Form1.List1, CStr(filex))
For i = 0 To Form1.List1.ListCount
If i = 0 Then
mainstring = Form1.List1.List(0)
Else
mainstring = mainstring & Chr(13) & Chr(10) & Form1.List1.List(i)
End If
Next i
MDIForm1.StatusBar1.SimpleText = lPath(filex, "\", False)
x = namex(CStr(filex))
Form1.Caption = x
Form1.txtCodeWin.Text = mainstring
Dim lngSave As Long
    Open MDIForm1.StatusBar1.SimpleText & "\temp.html" For Output As #1
        For lngSave& = 0 To Form1.List1.ListCount - 1
            Print #1, Form1.List1.List(lngSave&)
        Next lngSave&
    Close #1
    Form2.Browser1.Navigate (MDIForm1.StatusBar1.SimpleText & "\temp.html")
End Function
