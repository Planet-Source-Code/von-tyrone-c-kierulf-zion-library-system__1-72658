VERSION 5.00
Begin VB.Form frmSort 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sort"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   Icon            =   "frmSort.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmSort.frx":6852
      Left            =   120
      List            =   "frmSort.frx":6854
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1095
      Width           =   4935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Sort"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   1815
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   1815
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmSort.frx":6856
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Sort records by:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   855
      Width           =   1815
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   120
      X2              =   5040
      Y1              =   735
      Y2              =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Select a field to be sorted bellow and click 'Sort' button. Click 'Cancel' if you want to cancel sorting of records"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   165
      Width           =   4335
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      X1              =   120
      X2              =   5040
      Y1              =   1575
      Y2              =   1575
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   120
      X2              =   5040
      Y1              =   1575
      Y2              =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   120
      X2              =   5040
      Y1              =   735
      Y2              =   735
   End
End
Attribute VB_Name = "frmSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------
'Sort Form
'This entire form can be called from anywhere in the application to sort
'a given recordset according to a selected field.
'This code has been reused from the master forms and the Selection
'Dialog Box (frmSelectDg)
'-----------------------------------
Public SourceRs        As ADODB.RecordSet

Private Sub Command1_Click()

    'On Error GoTo Err
    SourceRs.Sort = Combo1.Text
    Unload Me

Exit Sub

Err:
    MsgBox "Please select a valid section from the list.", vbExclamation
    Combo1.SetFocus

End Sub

Private Sub Command2_Click()
    Set frmSort = Nothing
    Set SourceRs = Nothing
    Unload Me

End Sub

Private Sub Form_Activate()

    Combo1.SetFocus
    Combo1.ListIndex = 0

End Sub

Private Sub Form_Load()

    FillCombo Combo1, SourceRs, True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSort = Nothing
    Set SourceRs = Nothing

End Sub

