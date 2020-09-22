VERSION 5.00
Begin VB.Form frmOptions 
   Caption         =   "Options"
   ClientHeight    =   1395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   1395
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   1095
      Left            =   3360
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.OptionButton optRefresh 
      Caption         =   "Refresh After Each Character Typed"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   2895
   End
   Begin VB.OptionButton optWait 
      Caption         =   "Wait Until Finished Typing"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Value           =   -1  'True
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Refresh Options"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdOK_Click()
frmOptions.Hide
If optWait.Value = True Then
InifileWrite "she.ini", "options", "refresh", "False"
Else
InifileWrite "she.ini", "options", "refresh", "True"
End If
End Sub
