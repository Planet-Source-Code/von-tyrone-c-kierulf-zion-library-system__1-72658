VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   Caption         =   "About"
   ClientHeight    =   4035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6885
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   Picture         =   "frmAbout.frx":0000
   ScaleHeight     =   4035
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   6240
      Picture         =   "frmAbout.frx":5A28
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   1
      ToolTipText     =   "Close"
      Top             =   3360
      Width           =   375
   End
   Begin VB.ListBox List1 
      Height          =   840
      ItemData        =   "frmAbout.frx":5D5D
      Left            =   480
      List            =   "frmAbout.frx":5D6D
      TabIndex        =   0
      Top             =   1800
      Width           =   5895
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Picture1_Click()
    Set frmAbout = Nothing
    Unload Me
End Sub




