VERSION 5.00
Begin VB.Form frmLink 
   Caption         =   "Link"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3375
   ControlBox      =   0   'False
   LinkMode        =   1  'Source
   LinkTopic       =   "Form4"
   ScaleHeight     =   3795
   ScaleWidth      =   3375
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtText 
      Height          =   285
      Left            =   360
      TabIndex        =   8
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Frame Frame5 
      Caption         =   "Link Text"
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   2520
      Width           =   2895
   End
   Begin VB.TextBox txtTarget 
      Height          =   285
      Left            =   360
      TabIndex        =   6
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Frame Frame4 
      Caption         =   "Target Frame"
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox txtAnchor 
      Height          =   285
      Left            =   360
      TabIndex        =   4
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Frame Frame3 
      Caption         =   "Anchor"
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   2895
   End
   Begin VB.TextBox txtLinkTo 
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   2655
   End
   Begin VB.Frame Frame2 
      Caption         =   "Link to ..."
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "Link"
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This code is Intellectual Propriety of Brandon Cunningham.
Private Sub cmdCancel_Click()
cmdOK.SetFocus
txtLinkTo.text = ""
txtAnchor.text = ""
txtTarget.text = ""
txtText.text = ""
frmLink.Hide
End Sub

Private Sub cmdOK_Click()
txtText.text = Form1.txtCodeWin.SelText
Form1.flgx.text = "True"
If txtAnchor.text = "" Then ' if there is not an Anchor
Form1.txtCodeWin.SelText = "<a href=" & Chr(34) & txtLinkTo.text & Chr(34) & " target=" & Chr(34) & txtTarget.text & Chr(34) & ">" & txtText.text & "</a>"
Else 'if there is an Anchor
Form1.txtCodeWin.SelText = "<a href=" & Chr(34) & txtLinkTo.text & Chr(34) & "#" & txtAnchor.text & " target=" & Chr(34) & txtTarget.text & Chr(34) & ">" & txtText.text & "</a>"
End If
Form1.txtCodeWin.SelStart = Form1.txtCodeWin.SelStart - Len(txtText.text) - 4
Form1.txtCodeWin.SelLength = Len(txtText.text)
Form1.txtCodeWin.Refresh
Form1.flgx.text = "False"
cmdOK.SetFocus
txtLinkTo.text = ""
txtAnchor.text = ""
txtTarget.text = ""
txtText.text = ""
frmLink.Hide
Call RefreshPage
End Sub
