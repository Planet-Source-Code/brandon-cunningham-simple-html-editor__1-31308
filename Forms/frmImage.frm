VERSION 5.00
Begin VB.Form frmImage 
   Caption         =   "Image"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3375
   ControlBox      =   0   'False
   LinkMode        =   1  'Source
   LinkTopic       =   "Form4"
   ScaleHeight     =   4155
   ScaleWidth      =   3375
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkDefaults 
      Caption         =   "Use Default Height and Width"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   1800
      Value           =   1  'Checked
      Width           =   2655
   End
   Begin VB.CheckBox chkWPerc 
      Caption         =   "Percent"
      Height          =   195
      Left            =   2040
      TabIndex        =   8
      Top             =   3120
      Width           =   855
   End
   Begin VB.CheckBox chkHPerc 
      Caption         =   "Percent"
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txtAlt 
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Frame Frame5 
      Caption         =   "Mouseover Message"
      Height          =   615
      Left            =   240
      TabIndex        =   13
      Top             =   1080
      Width           =   2895
   End
   Begin VB.TextBox txtWidth 
      Height          =   285
      Left            =   360
      TabIndex        =   7
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Frame Frame4 
      Caption         =   "Width"
      Height          =   615
      Left            =   240
      TabIndex        =   12
      Top             =   2880
      Width           =   2895
   End
   Begin VB.TextBox txtHeight 
      Height          =   285
      Left            =   360
      TabIndex        =   5
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Height"
      Height          =   615
      Left            =   240
      TabIndex        =   11
      Top             =   2160
      Width           =   2895
   End
   Begin VB.TextBox txtSource 
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   2655
   End
   Begin VB.Frame Frame2 
      Caption         =   "Image Source"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "Image"
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This code is Intellectual Propriety of Brandon Cunningham.
Private Sub cmdCancel_Click()
cmdOK.SetFocus
txtSource.text = ""
chkDefaults.Value = 1
txtHeight.text = ""
chkHPerc.Value = 0
txtWidth.text = ""
chkWPerc.Value = 0
txtAlt.text = ""
frmImage.Hide
End Sub

Private Sub cmdOK_Click()
Dim x As String
x = Form1.txtCodeWin.SelText
Form1.flgx.text = "True"
Form1.txtCodeWin.SelText = "<img src=" & Chr(34) & txtSource.text & Chr(34)

If chkDefaults.Value = 1 Then
txtHeight.text = ""
chkHPerc.Value = 0
txtWidth.text = ""
chkWPerc = 0
Else
End If


If Not txtHeight = "" Then
Form1.txtCodeWin.SelText = " height=" & Chr(34) & txtHeight.text & Chr(34)
Else
chkHPerc.Value = 0
End If

If chkHPerc.Value = 1 Then
Form1.txtCodeWin.SelStart = Form1.txtCodeWin.SelStart - 1
Form1.txtCodeWin.SelText = "%"
Form1.txtCodeWin.SelStart = Form1.txtCodeWin.SelStart + 1
Else
End If

If Not txtWidth = "" Then
Form1.txtCodeWin.SelText = " width=" & Chr(34) & txtWidth.text & Chr(34)
Else
chkWPerc.Value = 0
End If

If chkWPerc.Value = 1 Then
Form1.txtCodeWin.SelStart = Form1.txtCodeWin.SelStart - 1
Form1.txtCodeWin.SelText = "%"
Form1.txtCodeWin.SelStart = Form1.txtCodeWin.SelStart + 1
Else
End If

If Not txtAlt.text = "" Then
Form1.txtCodeWin.SelText = " alt=" & Chr(34) & txtAlt.text & Chr(34)
Else
End If
Form1.txtCodeWin.SelText = ">" & x
Form1.txtCodeWin.SelStart = Form1.txtCodeWin.SelStart - Len(x)
Form1.txtCodeWin.SelLength = Len(x)
Form1.txtCodeWin.Refresh
Form1.flgx.text = "False"
cmdOK.SetFocus
txtSource.text = ""
chkDefaults.Value = 1
txtHeight.text = ""
chkHPerc.Value = 0
txtWidth.text = ""
chkWPerc.Value = 0
txtAlt.text = ""
frmImage.Hide
Call RefreshPage
End Sub

