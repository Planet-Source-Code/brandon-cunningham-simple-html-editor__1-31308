Attribute VB_Name = "Refresh"
'This code is Intellectual Propriety of Brandon Cunningham.
Option Explicit
Public Function RefreshPage()
Dim browserloc, curtext As String 'list2text and list1text will contain the ENTIRE contents of each list
Dim arrcurtext As Variant 'Array for inserting the text of txtCodeWin into List1
Dim i As Integer 'Counter for arrcurtext
'Load string from txtCodeWin to list1
curtext = Form1.txtCodeWin.Text
arrcurtext = Split(curtext, Chr(13) + Chr(10))
Form1.List1.Clear
For i = LBound(arrcurtext) To UBound(arrcurtext)
Form1.List1.AddItem arrcurtext(i)
Next i

curtext = Form1.txtCodeWin.Text
arrcurtext = Split(curtext, Chr(13) + Chr(10))
Form1.List2.Clear
For i = LBound(arrcurtext) To UBound(arrcurtext)
Form1.List2.AddItem arrcurtext(i)
Next i

Call ListSaveFile(Form1.List1, MDIForm1.StatusBar1.SimpleText & "\temp.html")
'MsgBox "refreshing" ' debug tool, comment out
'Check to make sure they are on the temp page, they may have linked out.
'if not on temp page goto temp page
Form2.Browser1.Refresh 'Refresh browser (DO NOT NAVIGATE!  Navigate resets scrollbar)
Wait 0.2 'Small wait, browser has to refresh
Form1.txtCodeWin.SetFocus 'set focus back to txtCodeWin
End Function
