VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkMode        =   1  'Source
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.ListBox aboutbox 
      BackColor       =   &H8000000F&
      Height          =   2010
      ItemData        =   "Form3.frx":0000
      Left            =   240
      List            =   "Form3.frx":0019
      TabIndex        =   2
      Top             =   360
      Width           =   4095
   End
   Begin VB.Frame Frame1 
      Caption         =   "About Simple HTML Editor"
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   2760
      Width           =   1095
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This code is Intellectual Propriety of Brandon Cunningham.
Private Sub Command1_Click()
frmAbout.Hide
Rem do nothing
End Sub
