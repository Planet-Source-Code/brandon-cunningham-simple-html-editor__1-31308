VERSION 5.00
Begin VB.Form frmNewPage 
   Caption         =   "New Page"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3735
   ControlBox      =   0   'False
   LinkMode        =   1  'Source
   LinkTopic       =   "Form4"
   ScaleHeight     =   346
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   249
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Forget It."
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "That's What I Want!"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   15
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox txtVLinkColor 
      Height          =   285
      Left            =   2040
      TabIndex        =   14
      Top             =   4080
      Width           =   1335
   End
   Begin VB.TextBox txtALinkColor 
      Height          =   285
      Left            =   360
      TabIndex        =   13
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Frame Frame7 
      Caption         =   "Visited Link Color"
      Height          =   615
      Left            =   1920
      TabIndex        =   12
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Frame Frame6 
      Caption         =   "Active Link Color"
      Height          =   615
      Left            =   240
      TabIndex        =   11
      Top             =   3840
      Width           =   1575
   End
   Begin VB.TextBox txtBGFile 
      Height          =   285
      Left            =   360
      TabIndex        =   10
      Top             =   3360
      Width           =   3015
   End
   Begin VB.Frame Frame5 
      Caption         =   "Background File"
      Height          =   615
      Left            =   240
      TabIndex        =   9
      Top             =   3120
      Width           =   3255
   End
   Begin VB.TextBox txtBGColor 
      Height          =   285
      Left            =   360
      TabIndex        =   8
      Top             =   2640
      Width           =   3015
   End
   Begin VB.Frame Frame4 
      Caption         =   "Background Color"
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   2400
      Width           =   3255
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Left            =   360
      TabIndex        =   6
      Top             =   1920
      Width           =   3015
   End
   Begin VB.Frame Frame3 
      Caption         =   "Description"
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   3255
   End
   Begin VB.TextBox txtKeywords 
      Height          =   285
      Left            =   360
      TabIndex        =   4
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Frame Frame2 
      Caption         =   "Keywords"
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   3255
   End
   Begin VB.TextBox txtPgTitle 
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   3000
   End
   Begin VB.Frame Frame1 
      Caption         =   "Page Title"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   3255
   End
   Begin VB.Frame frameNewPage 
      Caption         =   "New Page Properties"
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmNewPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This code is Intellectual Propriety of Brandon Cunningham.
Private Sub cmdCancel_Click()
txtPgTitle.text = ""
txtKeywords.text = ""
txtDescription.text = ""
txtBGColor.text = ""
txtBGFile.text = ""
txtALinkColor.text = ""
txtVLinkColor.text = ""
cmdOK.SetFocus
frmNewPage.Hide
End Sub

Private Sub cmdOK_Click()
Form1.flgx.text = "True"
Dim x As String
Form1.txtCodeWin.SelText = "<html>" & Chr(13) + Chr(10) & _
"<head>" & Chr(13) + Chr(10) & _
"<title>" & txtPgTitle.text & Chr(13) + Chr(10) & _
"</title>" & Chr(13) + Chr(10) & _
"<meta name=" & Chr(34) & "Keywords" & Chr(34) & " content=" & Chr(34) & txtKeywords.text & Chr(34) & ">" & Chr(13) + Chr(10) & _
"<meta name=" & Chr(34) & "Description" & Chr(34) & " content=" & Chr(34) & txtDescription.text & Chr(34) & ">" & Chr(13) + Chr(10) & _
"<meta name=" & Chr(34) & "Generator" & Chr(34) & " content=" & Chr(34) & "Simple HTML Editor" & Chr(34) & ">" & Chr(13) + Chr(10) & _
"</head>" & Chr(13) + Chr(10) & _
"<body bgcolor=" & Chr(34) & txtBGColor.text & Chr(34) & " background=" & Chr(34) & txtBGFile.text & Chr(34) & " alink=" & Chr(34) & txtALinkColor.text & Chr(34) & " vlink=" & Chr(34) & txtVLinkColor.text & Chr(34) & ">" & Chr(13) + Chr(10) & Chr(13) + Chr(10) & _
"</body>" & Chr(13) + Chr(10) & _
"</html>"
Form1.txtCodeWin.SelStart = Form1.txtCodeWin.SelStart - 18
Form1.txtCodeWin.Refresh
Form1.flgx.text = "False"
txtPgTitle.text = ""
txtKeywords.text = ""
txtDescription.text = ""
txtBGColor.text = ""
txtBGFile.text = ""
txtALinkColor.text = ""
txtVLinkColor.text = ""
cmdOK.SetFocus
frmNewPage.Hide
Call RefreshPage
End Sub

