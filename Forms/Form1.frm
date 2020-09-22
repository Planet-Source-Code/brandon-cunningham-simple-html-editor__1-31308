VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   3255
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCodeWin 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   0
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   0
      Width           =   4695
   End
   Begin VB.TextBox flgx 
      Enabled         =   0   'False
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Text            =   "false"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   450
      Left            =   2880
      TabIndex        =   1
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   500
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This code is intellectual propriety of Brandon Cunningham.
Option Explicit

Private Sub Form_Resize()
On Error GoTo subend:
Dim a, b, cp, dp, c, d, diff1, diff2 As Integer
a = MDIForm1.Left
b = MDIForm1.Top
cp = MDIForm1.Width
dp = MDIForm1.Height
Form1.Move 0, 0, cp - 184
c = Form1.Width
d = Form1.Height
diff1 = dp - d
diff2 = diff1 - 500
txtCodeWin.Move 0, 0, c - 100, d - 350
Form2.Move 0, d, cp - 175, diff2 - Val(MDIForm1.heightflag.text)
subend:
Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not Form1.txtCodeWin.text = "" Then
Call asksave2
Else
End If
End Sub

Private Sub asksave2()
Dim x As Integer
x = MsgBox("Would you like to save this file before exiting?", vbYesNo)
If x = 6 Then
Call Saveit
Else
End If
End Sub

Private Sub txtCodeWin_Change()
'Comments are outdated and inaccurate.
If flgx.text = "True" Then
Exit Sub
Else
End If
Dim list2text, list1text, curtext, browserloc As String 'list2text and list1text will contain the ENTIRE contents of each list
Dim arrcurtext As Variant 'Array for inserting the text of txtCodeWin into List1
Dim i As Integer 'Counter for arrcurtext
Wait 1 'wait command serves multiple purposes. _
Each time a new character is entered into txtCodeWin a new txtCodeWin_Change() _
event occurs.  Each new txtCodeWin_Change() pre-empts previous ones. _
By placing a wait command here each new event has to wait, causing a loop that waits _
until the user is done typing. _

'Browser refreshing takes up quite a bit of system resources, _
and saving and reading files causes disk thrashing so we only need to do this once when the _
user is finished typing. _

'Load string from txtCodeWin to list1
curtext = txtCodeWin.text
arrcurtext = Split(curtext, Chr(13) + Chr(10))
List1.Clear
For i = LBound(arrcurtext) To UBound(arrcurtext)
List1.AddItem arrcurtext(i)
Next i

list1text = ListText(List1)
list2text = ListText(List2)

'if the length of list1 and list2 are not *almost* equal Then

If list1text = list2text Then 'Compare list1 and list2
Exit Sub
Else
End If

curtext = txtCodeWin.text
arrcurtext = Split(curtext, Chr(13) + Chr(10))
List2.Clear
For i = LBound(arrcurtext) To UBound(arrcurtext)
List2.AddItem arrcurtext(i)
Next i

'Have list1 save the temp file
Call ListSaveFile(List1, MDIForm1.StatusBar1.SimpleText & "\temp.html")
'have list2 read from the temp file so that the len of list2 minus the EOF marker is equal to the text in list1
'MsgBox "refreshing" ' debug tool, comment out
'Check to make sure they are on the temp page, they may have linked out.
Form2.Browser1.Refresh 'Refresh browser (DO NOT NAVIGATE!  Navigate resets scrollbar)
Wait 0.2 'Small wait, browser has to refresh
Form1.txtCodeWin.SetFocus 'set focus back to txtCodeWin
End Sub

Private Sub txtCodeWin_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Form1.Top = 0
Form1.Left = 0
End Sub

Private Sub txtCodeWin_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
If Not Form1.txtCodeWin.text = "" Then
Call asksave
Else
End If
Dim filex As String
filex = CStr(Data.Files(1))
Call onDragDropx(filex)
End Sub

Private Sub asksave()
Dim x As Integer
x = MsgBox("Would you like to save this file before opening a new document?", vbYesNo)
If x = 6 Then
Call Saveit
Else
End If
End Sub
