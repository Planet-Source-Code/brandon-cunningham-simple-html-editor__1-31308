VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Simple HTML Editor"
   ClientHeight    =   5820
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7635
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ColorList 
      Left            =   5520
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   12
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0278
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":04F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0768
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":09E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0C58
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0ED0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1148
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":13C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1638
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":18B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1B28
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1DA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2018
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2290
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2508
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog FontDialog 
      Left            =   3960
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select Font"
      FontName        =   "Times New Roman"
      FontSize        =   12
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   5550
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   476
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   840
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2780
            Key             =   ""
            Object.Tag             =   "&New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2C78
            Key             =   ""
            Object.Tag             =   "O&pen"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":3178
            Key             =   ""
            Object.Tag             =   "&Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":36A4
            Key             =   ""
            Object.Tag             =   "Cu&t"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":3BBC
            Key             =   ""
            Object.Tag             =   "&Copy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":40D4
            Key             =   ""
            Object.Tag             =   "&Paste"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":460C
            Key             =   ""
            Object.Tag             =   "&Change Font"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":4930
            Key             =   ""
            Object.Tag             =   "&Bold"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":4E04
            Key             =   ""
            Object.Tag             =   "&Italic"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":52BC
            Key             =   ""
            Object.Tag             =   "&Underline"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":5798
            Key             =   ""
            Object.Tag             =   "&Center"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":5C44
            Key             =   ""
            Object.Tag             =   "&Right"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":5F08
            Key             =   ""
            Object.Tag             =   "&Bullets"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":63CC
            Key             =   ""
            Object.Tag             =   "&Tab"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":6894
            Key             =   ""
            Object.Tag             =   "&Break Line"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":6D5C
            Key             =   ""
            Object.Tag             =   "&Horizontal Rule"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":71FC
            Key             =   ""
            Object.Tag             =   "&Link"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":7754
            Key             =   ""
            Object.Tag             =   "&Image"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":7CC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":7FFE
            Key             =   ""
            Object.Tag             =   "Help &Contents"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":8336
            Key             =   ""
            Object.Tag             =   "&About"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog OpenDialog 
      Left            =   3000
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   420
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageCombo ColorCombo 
         Height          =   330
         Left            =   2760
         TabIndex        =   7
         Top             =   45
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
         Text            =   "Black"
         ImageList       =   "ColorList"
      End
      Begin VB.ListBox List1 
         Height          =   255
         Left            =   5625
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   90
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSComctlLib.ImageCombo SizeCombo 
         Height          =   330
         Left            =   4005
         TabIndex        =   5
         Top             =   45
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
         Text            =   "12"
      End
      Begin MSComctlLib.ImageCombo FontCombo 
         Height          =   330
         Left            =   120
         TabIndex        =   4
         Top             =   45
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
         Text            =   "Times New Roman"
         ImageList       =   "ImageList1"
      End
      Begin VB.TextBox heightflag 
         Height          =   285
         Left            =   5850
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   90
         Visible         =   0   'False
         Width           =   420
      End
      Begin MSComctlLib.Toolbar Toolbar3 
         Height          =   390
         Left            =   4680
         TabIndex        =   2
         Top             =   0
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Bold"
               Object.ToolTipText     =   "Bold"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Italic"
               Object.ToolTipText     =   "Italic"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Underline"
               Object.ToolTipText     =   "Underline"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Center"
               Object.ToolTipText     =   "Center"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Right"
               Object.ToolTipText     =   "Right"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Bullets"
               Object.ToolTipText     =   "Bullets"
               ImageIndex      =   13
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ChangeFont"
            Object.ToolTipText     =   "Change Font"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Tab"
            Object.ToolTipText     =   "Tab"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BreakLine"
            Object.ToolTipText     =   "Break Line"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "HR"
            Object.ToolTipText     =   "Horizontal Rule"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Link"
            Object.ToolTipText     =   "Link"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Image"
            Object.ToolTipText     =   "Image"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "About"
            Object.ToolTipText     =   "About"
            ImageIndex      =   21
         EndProperty
      EndProperty
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu new 
         Caption         =   "&New"
      End
      Begin VB.Menu nsep1 
         Caption         =   "-"
      End
      Begin VB.Menu Open 
         Caption         =   "O&pen"
      End
      Begin VB.Menu nsep2 
         Caption         =   "-"
      End
      Begin VB.Menu Save 
         Caption         =   "&Save"
      End
      Begin VB.Menu SaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "&Edit"
      Begin VB.Menu Cut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu Copy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu Paste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu Delete 
         Caption         =   "De&lete"
      End
      Begin VB.Menu SelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu Fonts 
      Caption         =   "Fo&nts"
      Begin VB.Menu ChangeFont 
         Caption         =   "&Change Font"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu Bold 
         Caption         =   "&Bold"
      End
      Begin VB.Menu Italic 
         Caption         =   "&Italic"
      End
      Begin VB.Menu Underline 
         Caption         =   "&Underline"
      End
      Begin VB.Menu newsep2 
         Caption         =   "-"
      End
      Begin VB.Menu Strike 
         Caption         =   "&Strike"
      End
   End
   Begin VB.Menu Format 
      Caption         =   "Fo&rmat"
      Begin VB.Menu text 
         Caption         =   "&Text"
         Begin VB.Menu Center 
            Caption         =   "&Center"
         End
         Begin VB.Menu Right 
            Caption         =   "&Right"
         End
         Begin VB.Menu Bullets 
            Caption         =   "&Bullets"
         End
      End
      Begin VB.Menu newsep3 
         Caption         =   "-"
      End
      Begin VB.Menu Lines 
         Caption         =   "&Lines"
         Begin VB.Menu Tab 
            Caption         =   "&Tab"
         End
         Begin VB.Menu BreakLine 
            Caption         =   "&Break Line"
         End
         Begin VB.Menu HR 
            Caption         =   "&Horizontal Rule"
         End
         Begin VB.Menu sep 
            Caption         =   "-"
         End
         Begin VB.Menu PageBreak 
            Caption         =   "&Page Break"
         End
      End
   End
   Begin VB.Menu Insert 
      Caption         =   "&Insert"
      Begin VB.Menu Link 
         Caption         =   "&Link"
      End
      Begin VB.Menu newsep4 
         Caption         =   "-"
      End
      Begin VB.Menu Image 
         Caption         =   "&Image"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
      Begin VB.Menu About 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This code is Intellectual Propriety of Brandon Cunningham.
Private Sub About_Click()
frmAbout.Show 1
End Sub

Private Sub Bold_Click()
Dim x As String
x = Form1.txtCodeWin.SelText
Form1.flgx.text = "True"
Form1.txtCodeWin.SelText = "<b>" & x & "</b>"
Form1.txtCodeWin.SelStart = Form1.txtCodeWin.SelStart - 4 - Len(x)
Form1.txtCodeWin.SelLength = Len(x)
Form1.txtCodeWin.Refresh
Form1.flgx.text = "False"
Call RefreshPage
End Sub

Private Sub BreakLine_Click()
Dim x As String
x = Form1.txtCodeWin.SelText
Form1.flgx.text = "True"
Form1.txtCodeWin.SelText = "<br>" & Chr(13) + Chr(10) & x
Form1.txtCodeWin.SelStart = Form1.txtCodeWin.SelStart - Len(x)
Form1.txtCodeWin.SelLength = Len(x)
Form1.txtCodeWin.Refresh
Form1.flgx.text = "False"
Call RefreshPage
End Sub

Private Sub Bullets_Click()
Dim x As String
Dim i As Integer
On Error GoTo lierror:
Form1.flgx.text = "True"
x = InputBox("Insert Number of Bullets")
Form1.txtCodeWin.SelText = "<ul>" & Chr(13) + Chr(10)
For i = 1 To Val(x)
Form1.txtCodeWin.SelText = "<li>" & InputBox("Insert Bullet Item #" & i) & Chr(13) & Chr(10)
Next i
Form1.txtCodeWin.SelText = "</ul>"
lierror:
Form1.txtCodeWin.Refresh
Form1.flgx.text = "False"
Call RefreshPage
End Sub

Private Sub Center_Click()
Dim x As String
x = Form1.txtCodeWin.SelText
Form1.flgx.text = "True"
Form1.txtCodeWin.SelText = "<center>" & x & "</center>"
Form1.txtCodeWin.SelStart = Form1.txtCodeWin.SelStart - 9 - Len(x)
Form1.txtCodeWin.SelLength = Len(x)
Form1.txtCodeWin.Refresh
Form1.flgx.text = "False"
Call RefreshPage
End Sub
Private Sub ChangeFont_Click()
   FontDialog.CancelError = True
   On Error GoTo ErrHandler
Dim x As String
Dim Color As String
x = Form1.txtCodeWin.SelText
Form1.flgx.text = "True"
FontDialog.Flags = cdlCFBoth Or cdlCFEffects
FontDialog.ShowFont
Color = CStr(FontDialog.Color)
Select Case Color
Case "0"
Color = "#000000"
Case "128"
Color = "#CC0000"
Case "32768"
Color = "#00CC00"
Case "32896"
Color = "#CCCC00"
Case "8388608"
Color = "#0000CC"
Case Is = "8388736"
Color = "#CC00CC"
Case Is = "8421376"
Color = "#00CCCC"
Case Is = "8421504"
Color = "#999999"
Case Is = "12632256"
Color = "#CCCCCC"
Case Is = "255"
Color = "#FF0000"
Case Is = "65280"
Color = "#00FF00"
Case Is = "65535"
Color = "#FFFF00"
Case Is = "16711680"
Color = "#0000FF"
Case Is = "16711935"
Color = "#FF00FF"
Case Is = "16776960"
Color = "#00FFFF"
Case Is = "16777215"
Color = "#FFFFFF"
End Select
Form1.txtCodeWin.SelText = "<font face=" & Chr(34) & FontDialog.FontName & Chr(34) & " style=" & Chr(34) & "font-size:" & FontDialog.FontSize & "pt" & Chr(34) & " color=" & Chr(34) & Color & Chr(34) & ">" & x & "</font>"
Form1.txtCodeWin.SelStart = Form1.txtCodeWin.SelStart - 7 - Len(x)
Form1.txtCodeWin.SelLength = Len(x)
Form1.txtCodeWin.Refresh
Form1.flgx.text = "False"
Call RefreshPage
If FontDialog.FontBold = True Then
Call Bold_Click
Else
End If
If FontDialog.FontItalic = True Then
Call Italic_Click
Else
End If
If FontDialog.FontStrikethru = True Then
Call Strike_Click
Else
End If
If FontDialog.FontUnderline = True Then
Call Underline_Click
Else
End If
ErrHandler:
Exit Sub
End Sub

Private Sub ColorCombo_Click()
Dim Color As String
Select Case ColorCombo.SelectedItem
Case "Black"
Color = "#000000"
Case "Maroon"
Color = "#CC0000"
Case "Green"
Color = "#00CC00"
Case "Olive"
Color = "#CCCC00"
Case "Navy"
Color = "#0000CC"
Case Is = "Purple"
Color = "#CC00CC"
Case Is = "Teal"
Color = "#00CCCC"
Case Is = "Gray"
Color = "#999999"
Case Is = "Silver"
Color = "#CCCCCC"
Case Is = "Red"
Color = "#FF0000"
Case Is = "Lime"
Color = "#00FF00"
Case Is = "Yellow"
Color = "#FFFF00"
Case Is = "Blue"
Color = "#0000FF"
Case Is = "Fuchsia"
Color = "#FF00FF"
Case Is = "Aqua"
Color = "#00FFFF"
Case Is = "White"
Color = "#FFFFFF"
End Select

Dim x As String
x = Form1.txtCodeWin.SelText
Form1.flgx.text = "True"
Form1.txtCodeWin.SelText = "<font color=" & Chr(34) & Color & Chr(34) & ">" & x & "</font>"
Form1.txtCodeWin.SelStart = Form1.txtCodeWin.SelStart - 7 - Len(x)
Form1.txtCodeWin.SelLength = Len(x)
Form1.txtCodeWin.Refresh
Form1.flgx.text = "False"
Call RefreshPage
End Sub

Private Sub CoolBar1_HeightChanged(ByVal NewHeight As Single)
Select Case NewHeight
Case 450
heightflag.text = "980"
Case 810
heightflag.text = "1340"
End Select
Call MDIForm_Resize
End Sub

Private Sub Copy_Click()
Clipboard.SetText Form1.txtCodeWin.SelText
End Sub

Private Sub Delete_Click()
Form1.flgx.text = "True"
Form1.txtCodeWin.SelText = ""
Form1.txtCodeWin.Refresh
Form1.flgx.text = "False"
Call RefreshPage
End Sub

Private Sub FontCombo_Click()
Dim x As String
x = Form1.txtCodeWin.SelText
Form1.flgx.text = "True"
Form1.txtCodeWin.SelText = "<font face=" & Chr(34) & FontCombo.SelectedItem & Chr(34) & ">" & x & "</font>"
Form1.txtCodeWin.SelStart = Form1.txtCodeWin.SelStart - 7 - Len(x)
Form1.txtCodeWin.SelLength = Len(x)
Form1.txtCodeWin.Refresh
Form1.flgx.text = "False"
Call RefreshPage
End Sub

Private Sub HelpContents_Click()
Shellx "open", "hh.exe", App.Path & "\help\index.chm"
End Sub
Private Sub HR_Click()
frmHR.Show 1
End Sub

Private Sub Image_Click()
frmImage.Show 1
End Sub
Private Sub Italic_Click()
Dim x As String
x = Form1.txtCodeWin.SelText
Form1.flgx.text = "True"
Form1.txtCodeWin.SelText = "<i>" & x & "</i>"
Form1.txtCodeWin.SelStart = Form1.txtCodeWin.SelStart - 4 - Len(x)
Form1.txtCodeWin.SelLength = Len(x)
Form1.txtCodeWin.Refresh
Form1.flgx.text = "False"
Call RefreshPage
End Sub

Private Sub Link_Click()
frmLink.txtText.text = Form1.txtCodeWin.SelText 'Get selected text and place it as Link Text
frmLink.Show 1
End Sub

Private Sub MDIForm_Load()
heightflag.text = "1360"
'Call mCoolMenu.Install(Me.hWnd, , ImageList1) Batteries not included.
'mCoolMenu.FontName Me.hWnd, "Comic Sans MS" Batteries not included.
'mCoolMenu.FontSize Me.hWnd, 8& Batteries not included.
MDIForm1.StatusBar1.SimpleText = App.Path 'store the current path
Dim curfile As String
curfile = "Untitled"
Form2.Show
Form1.Show
Form1.Caption = curfile
Form1.List1.Clear
Dim lngSave As Long
    Open App.Path & "\temp.html" For Output As #1
        For lngSave& = 0 To Form1.List1.ListCount - 1
            Print #1, Form1.List1.List(lngSave&)
        Next lngSave&
    Close #1
Form2.Browser1.Navigate (App.Path & "\temp.html")
If Not Command$ = "" Then
Call command_line
Else
End If

  Dim i As Long
  For i = 0 To Screen.FontCount - 1
  List1.AddItem CStr(Screen.Fonts(i))
  Next i
 
 Dim savei As Long
 
  For i = 0 To List1.ListCount - 1
  If List1.List(i) = "Times New Roman" Then
  'MsgBox List1.List(i)
  savei = i
  Else
  End If
  FontCombo.ComboItems.Add , , List1.List(i), 19
          Next i
        
        ColorCombo.ComboItems.Add , , "Black", 1
        ColorCombo.ComboItems.Add , , "Maroon", 2
        ColorCombo.ComboItems.Add , , "Green", 3
        ColorCombo.ComboItems.Add , , "Olive", 4
        ColorCombo.ComboItems.Add , , "Navy", 5
        ColorCombo.ComboItems.Add , , "Purple", 6
        ColorCombo.ComboItems.Add , , "Teal", 7
        ColorCombo.ComboItems.Add , , "Gray", 8
        ColorCombo.ComboItems.Add , , "Silver", 9
        ColorCombo.ComboItems.Add , , "Red", 10
        ColorCombo.ComboItems.Add , , "Lime", 11
        ColorCombo.ComboItems.Add , , "Yellow", 12
        ColorCombo.ComboItems.Add , , "Blue", 13
        ColorCombo.ComboItems.Add , , "Fuchsia", 14
       ColorCombo.ComboItems.Add , , "Aqua", 15
       ColorCombo.ComboItems.Add , , "White", 16
       
       SizeCombo.ComboItems.Add , , 8
       SizeCombo.ComboItems.Add , , 10
       SizeCombo.ComboItems.Add , , 12
       SizeCombo.ComboItems.Add , , 14
       SizeCombo.ComboItems.Add , , 18
       SizeCombo.ComboItems.Add , , 24
       SizeCombo.ComboItems.Add , , 36
                        
      FontCombo.SelectedItem = FontCombo.ComboItems(savei + 1)
      ColorCombo.SelectedItem = ColorCombo.ComboItems(1)
      FontCombo.Refresh
      ColorCombo.Refresh
      SizeCombo.Refresh
      Form1.SetFocus
       End Sub

Private Sub MDIForm_Resize()
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
Form2.Move 0, d, cp - 175, diff2 - Val(heightflag.text)
Exit Sub
subend:
Call move2
Exit Sub
End Sub
Private Sub move2()
Dim a, b, cp, dp As Integer
a = MDIForm1.Left
b = MDIForm1.Top
cp = MDIForm1.Width
dp = MDIForm1.Height
On Error GoTo subend2:
Form1.Move 0, 0, cp - 184, dp - 2100
Exit Sub
subend2:
Exit Sub
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
' Call mCoolMenu.Uninstall(Me.hWnd) Batteries not included.
Unload frmAbout
Unload frmHR
Unload frmImage
Unload frmLink
Unload frmNewPage
Unload Form1
Unload Form2
Unload MDIForm1
End Sub

Private Sub New_Click()
Form1.Top = 0
Form1.Left = 0
If Not Form1.txtCodeWin = vbNullString Then
Call newsaveoldfile
Else
MDIForm1.StatusBar1.SimpleText = App.Path
Form1.txtCodeWin.text = ""
Form1.Caption = "Untitled"
frmNewPage.Show 1
Call savelist1
End If
End Sub

Private Sub savelist1()
Dim curtext As String
Dim arrcurtext As Variant
Dim i As Integer
curtext = Form1.txtCodeWin.text
arrcurtext = Split(curtext, Chr(13) + Chr(10))
Form1.List1.Clear
For i = LBound(arrcurtext) To UBound(arrcurtext)
Form1.List1.AddItem arrcurtext(i)
Next i
Call ListSaveFile(Form1.List1, MDIForm1.StatusBar1.SimpleText & "\temp.html")
Form2.Browser1.Navigate MDIForm1.StatusBar1.SimpleText & "\temp.html"
End Sub
Private Sub newsaveoldfile()
Dim x As Integer
x = MsgBox("Would you like to save this file before creating a new one?", vbYesNo)
If x = 6 Then
Call SaveAs_Click
Else
MDIForm1.StatusBar1.SimpleText = App.Path
Form1.txtCodeWin.text = ""
Form1.Caption = "Untitled"
frmNewPage.Show 1
End If
End Sub
Private Sub Open_Click()
Dim filex As Variant
Dim mainstring, x As String
Dim i As Integer
If Not Form1.txtCodeWin.text = "" Then
Call asksave
Else
End If
OpenDialog.DialogTitle = "Select a HTML file to open"
OpenDialog.Filter = "HTML Files|*.html;*.htm|All Files|*.*"
OpenDialog.ShowOpen
filex = OpenDialog.FileName
If filex = "" Then
Exit Sub
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
MsgBox filex
MsgBox lpath(CStr(filex))
MDIForm1.StatusBar1.SimpleText = lpath(CStr(filex))
x = namex(CStr(filex))
Form1.Caption = x
Form1.txtCodeWin.text = mainstring
Dim lngSave As Long
    Open MDIForm1.StatusBar1.SimpleText & "\temp.html" For Output As #1
        For lngSave& = 0 To Form1.List1.ListCount - 1
            Print #1, Form1.List1.List(lngSave&)
        Next lngSave&
    Close #1
    Form2.Browser1.Navigate (MDIForm1.StatusBar1.SimpleText & "\temp.html")
End Sub

Private Sub asksave()
Dim x As Integer
x = MsgBox("Would you like to save this file before opening a new document?", vbYesNo)
If x = 6 Then
Call Saveit
Else
End If
End Sub

Private Sub Open2_Click()
Dim filex As Variant
Dim myfile As String
OpenDialog.DialogTitle = "Select a HTML file to open"
OpenDialog.Filter = "HTML Files|*.html;*.htm|All Files|*.*"
OpenDialog.ShowOpen
filex = OpenDialog.FileName
myfile = CStr(filex)
Shellx "Open", App.Path & "\she.exe", myfile
End Sub

Private Sub PageBreak_Click()
Dim x As String
x = Form1.txtCodeWin.SelText
Form1.flgx.text = "True"
Form1.txtCodeWin.SelText = "<p>" & x
Form1.txtCodeWin.SelStart = Form1.txtCodeWin.SelStart - Len(x)
Form1.txtCodeWin.SelLength = Len(x)
Form1.txtCodeWin.Refresh
Form1.flgx.text = "False"
Call RefreshPage
End Sub

Private Sub Paste_Click()
Form1.flgx.text = "True"
Form1.txtCodeWin.SelText = Clipboard.GetText()
Form1.flgx.text = "False"
Call RefreshPage
End Sub

Private Sub Right_Click()
Dim x As String
Form1.flgx.text = "True"
x = Form1.txtCodeWin.SelText
Form1.txtCodeWin.SelText = "<table width=" & Chr(34) & "100%" & Chr(34) & "><td align=" & Chr(34) & "right" & Chr(34) & ">" & Chr(13) + Chr(10) & x & Chr(13) + Chr(10) & "</td></table>"
Form1.txtCodeWin.SelStart = Form1.txtCodeWin.SelStart - 15 - Len(x)
Form1.txtCodeWin.SelLength = Len(x)
Form1.txtCodeWin.Refresh
Form1.flgx.text = "False"
Call RefreshPage
End Sub

Private Sub Save_Click()
If Form1.Caption = "Untitled" Then
Call SaveAs_Click
Else
End If
ListSaveFile Form1.List1, MDIForm1.StatusBar1.SimpleText & "\" & Form1.Caption
End Sub

Private Sub SaveAs_Click()
Dim filex, arrcurtext As Variant
Dim mainstring, x, curtext As String
Dim i As Integer
OpenDialog.DialogTitle = "Select a HTML file to open"
OpenDialog.Filter = "HTML Files|*.html;*.htm|All Files|*.*"
OpenDialog.ShowSave
filex = OpenDialog.FileName
If filex = "" Then
Exit Sub
Else
End If
curtext = Form1.txtCodeWin.text
arrcurtext = Split(curtext, Chr(13) + Chr(10))
Form1.List1.Clear
For i = LBound(arrcurtext) To UBound(arrcurtext)
Form1.List1.AddItem arrcurtext(i)
Next i
Call ListSaveFile(Form1.List1, CStr(filex))
MDIForm1.StatusBar1.SimpleText = lpath(CStr(filex))
x = namex(CStr(filex))
Form1.Caption = x
Dim lngSave As Long
    Open MDIForm1.StatusBar1.SimpleText & "\temp.html" For Output As #1
        For lngSave& = 0 To Form1.List1.ListCount - 1
            Print #1, Form1.List1.List(lngSave&)
        Next lngSave&
    Close #1
    Form2.Browser1.Navigate (MDIForm1.StatusBar1.SimpleText & "\temp.html")
End Sub

Private Sub SelectAll_Click()
Form1.txtCodeWin.SelStart = 0
Form1.txtCodeWin.SelLength = Len(Form1.txtCodeWin.text)
End Sub

Private Sub SizeCombo_Click()
Dim x As String
Dim convertsize As String
Select Case SizeCombo.SelectedItem
Case "8"
convertsize = "1"
Case "10"
convertsize = "2"
Case "12"
convertsize = "3"
Case "14"
convertsize = "4"
Case "18"
convertsize = "5"
Case "24"
convertsize = "6"
Case "36"
convertsize = "7"
End Select
x = Form1.txtCodeWin.SelText
Form1.flgx.text = "True"
Form1.txtCodeWin.SelText = "<font size=" & Chr(34) & convertsize & Chr(34) & ">" & x & "</font>"
Form1.txtCodeWin.SelStart = Form1.txtCodeWin.SelStart - 7 - Len(x)
Form1.txtCodeWin.SelLength = Len(x)
Form1.txtCodeWin.Refresh
Form1.flgx.text = "False"
Call RefreshPage
End Sub

Private Sub Tab_Click()
Dim x As String
Form1.flgx.text = "True"
x = Form1.txtCodeWin.SelText
Form1.txtCodeWin.SelText = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & x
Form1.txtCodeWin.SelStart = Form1.txtCodeWin.SelStart - Len(x)
Form1.txtCodeWin.SelLength = Len(x)
Form1.txtCodeWin.Refresh
Form1.flgx.text = "False"
Call RefreshPage
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "New"
Call New_Click
Case "Open"
Call Open_Click
Case "Save"
Call Save_Click
Case "Cut"
Call Cut_Click
Case "Copy"
Call Copy_Click
Case "Paste"
Call Paste_Click
Case "ChangeFont"
Call ChangeFont_Click
Case "Bold"
Call Bold_Click
Case "Italic"
Call Italic_Click
Case "Underline"
Call Underline_Click
Case "Center"
Call Center_Click
Case "Right"
Call Right_Click
Case "Bullets"
Call Bullets_Click
Case "Tab"
Call Tab_Click
Case "BreakLine"
Call BreakLine_Click
Case "HR"
Call HR_Click
Case "Link"
Call Link_Click
Case "Image"
Call Image_Click
Case "Help"
Call HelpContents_Click
Case "About"
Call About_Click
End Select
End Sub

Private Sub Cut_Click()
Form1.flgx.text = "True"
Clipboard.SetText Form1.txtCodeWin.SelText
Form1.txtCodeWin.SelText = ""
Form1.txtCodeWin.Refresh
Form1.flgx.text = "False"
Call RefreshPage
End Sub
Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "Bold"
Call Bold_Click
Case "Italic"
Call Italic_Click
Case "Underline"
Call Underline_Click
Case "Center"
Call Center_Click
Case "Right"
Call Right_Click
Case "Bullets"
Call Bullets_Click
End Select
End Sub

Private Sub Underline_Click()
Dim x As String
Form1.flgx.text = "True"
x = Form1.txtCodeWin.SelText
Form1.txtCodeWin.SelText = "<u>" & x & "</u>"
Form1.txtCodeWin.SelStart = Form1.txtCodeWin.SelStart - 4 - Len(x)
Form1.txtCodeWin.SelLength = Len(x)
Form1.txtCodeWin.Refresh
Form1.flgx.text = "False"
Call RefreshPage
End Sub

Private Sub Strike_Click()
Dim x As String
Form1.flgx.text = "True"
x = Form1.txtCodeWin.SelText
Form1.txtCodeWin.SelText = "<s>" & x & "</s>"
Form1.txtCodeWin.SelStart = Form1.txtCodeWin.SelStart - 4 - Len(x)
Form1.txtCodeWin.SelLength = Len(x)
Form1.txtCodeWin.Refresh
Form1.flgx.text = "False"
Call RefreshPage
End Sub

Private Sub command_line()
Dim mainstring, x As String
Dim filex As String
Dim i As Integer
filex = CStr(Command$)
If filex = "" Then
Exit Sub
Else
End If
If Left(filex, 1) = Chr(34) Then
filex = xsubstr(filex, Len(filex) - 1, Len(filex) - 2)
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
MDIForm1.StatusBar1.SimpleText = lpath(CStr(filex))
x = namex(CStr(filex))
Form1.Caption = x
Form1.txtCodeWin.text = mainstring
Dim lngSave As Long
    Open MDIForm1.StatusBar1.SimpleText & "\temp.html" For Output As #1
        For lngSave& = 0 To Form1.List1.ListCount - 1
            Print #1, Form1.List1.List(lngSave&)
        Next lngSave&
    Close #1
    Form2.Browser1.Navigate (MDIForm1.StatusBar1.SimpleText & "\temp.html")
End Sub

