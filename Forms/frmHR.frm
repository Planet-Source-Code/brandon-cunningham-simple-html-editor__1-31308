VERSION 5.00
Begin VB.Form frmHR 
   Caption         =   "Horizontal Rule"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3375
   ControlBox      =   0   'False
   LinkMode        =   1  'Source
   LinkTopic       =   "Form3"
   ScaleHeight     =   3840
   ScaleWidth      =   3375
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   12
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtColor 
      Height          =   285
      Left            =   360
      TabIndex        =   11
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Frame Frame5 
      Caption         =   "Color"
      Height          =   615
      Left            =   240
      TabIndex        =   10
      Top             =   2520
      Width           =   2895
   End
   Begin VB.TextBox txtSize 
      Height          =   285
      Left            =   360
      TabIndex        =   9
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Frame Frame4 
      Caption         =   "Size (Height)"
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Width           =   2895
   End
   Begin VB.OptionButton optRight 
      Caption         =   "Right"
      Height          =   195
      Left            =   2160
      TabIndex        =   7
      Top             =   1320
      Width           =   735
   End
   Begin VB.OptionButton optCenter 
      Caption         =   "Center"
      Height          =   195
      Left            =   1200
      TabIndex        =   6
      Top             =   1320
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.OptionButton optLeft 
      Caption         =   "Left"
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   1320
      Width           =   615
   End
   Begin VB.Frame Frame3 
      Caption         =   "Alignment"
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   2895
   End
   Begin VB.CheckBox chkPerc 
      Caption         =   "Percentage"
      Height          =   195
      Left            =   1800
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtWidth 
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Width"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "Horizontal Rule"
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmHR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This code is Intellectual Propriety of Brandon Cunningham.
Private Sub cmdCancel_Click()
txtWidth.text = ""
chkPerc.Value = 0
optCenter.Value = True
txtSize.text = ""
txtColor.text = ""
cmdOK.SetFocus
frmHR.Hide
End Sub

Private Sub cmdOK_Click()
Form1.flgx.text = "True"
Dim blnFlag As Boolean
Form1.txtCodeWin.SelText = "<hr"
If Not txtWidth.text = "" Then
Form1.txtCodeWin.SelText = " width=" & Chr(34) & txtWidth.text & Chr(34)
blnFlag = True
Else
blnFlag = False
End If

If blnFlag = True And chkPerc.Value = 1 Then
Form1.txtCodeWin.SelStart = Form1.txtCodeWin.SelStart - 1
Form1.txtCodeWin.SelText = "%"
Form1.txtCodeWin.SelStart = Form1.txtCodeWin.SelStart + 1
Else
End If

If optLeft.Value = True Then
Form1.txtCodeWin.SelText = " align=" & Chr(34) & "left" & Chr(34)
Else
End If

If optRight.Value = True Then
Form1.txtCodeWin.SelText = " align=" & Chr(34) & "right" & Chr(34)
Else
End If

If Not txtSize.text = "" Then
Form1.txtCodeWin.SelText = " size=" & Chr(34) & txtSize.text & Chr(34)
Else
End If

If Not txtColor.text = "" Then
Form1.txtCodeWin.SelText = " color=" & Chr(34) & txtColor.text & Chr(34)
Else
End If

Form1.txtCodeWin.SelText = ">"

txtWidth.text = ""
chkPerc.Value = 0
optCenter.Value = True
txtSize.text = ""
txtColor.text = ""
cmdOK.SetFocus
frmHR.Hide

Form1.flgx.text = "False"

Form1.txtCodeWin.Refresh

Call RefreshPage

End Sub

