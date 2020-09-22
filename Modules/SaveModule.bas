Attribute VB_Name = "SaveModule"
'This code is Intellectual Propriety of Brandon Cunningham.
Public Function Saveit()
If Form1.Caption = "Untitled" Then
Call SaveAsit
Else
End If
ListSaveFile Form1.List1, MDIForm1.StatusBar1.SimpleText & "\" & Form1.Caption
End Function

Public Function SaveAsit()
Dim filex, arrcurtext As Variant
Dim mainstring, x, curtext As String
Dim i As Integer
MDIForm1.OpenDialog.DialogTitle = "Select a HTML file to open"
MDIForm1.OpenDialog.Filter = "HTML Files|*.html;*.htm|All Files|*.*"
MDIForm1.OpenDialog.ShowSave
filex = MDIForm1.OpenDialog.FileName
If filex = "" Then
Exit Function
Else
End If
curtext = Form1.txtCodeWin.text
arrcurtext = Split(curtext, Chr(13) + Chr(10))
Form1.List1.Clear
For i = LBound(arrcurtext) To UBound(arrcurtext)
Form1.List1.AddItem arrcurtext(i)
Next i
Call ListSaveFile(Form1.List1, CStr(filex))
MDIForm1.StatusBar1.SimpleText = lpath(filex, "\", False)
x = namex(CStr(filex))
Form1.Caption = x
Dim lngSave As Long
    Open MDIForm1.StatusBar1.SimpleText & "\temp.html" For Output As #1
        For lngSave& = 0 To Form1.List1.ListCount - 1
            Print #1, Form1.List1.List(lngSave&)
        Next lngSave&
    Close #1
    Form2.Browser1.Navigate (MDIForm1.StatusBar1.SimpleText & "\temp.html")
End Function
