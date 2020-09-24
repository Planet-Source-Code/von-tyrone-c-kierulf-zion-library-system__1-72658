VERSION 5.00
Begin VB.Form frmFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Filter"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
   Icon            =   "frmFilter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   5355
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
      ItemData        =   "frmFilter.frx":6852
      Left            =   240
      List            =   "frmFilter.frx":6854
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1935
      Width           =   4935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Filter"
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
      Left            =   2880
      TabIndex        =   2
      Top             =   2655
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
      Left            =   4080
      TabIndex        =   1
      Top             =   2655
      Width           =   1095
   End
   Begin VB.TextBox Text1 
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
      Left            =   240
      TabIndex        =   0
      Top             =   1215
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmFilter.frx":6856
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   6
      Top             =   120
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmFilter.frx":68E0
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Look in:"
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
      Left            =   240
      TabIndex        =   5
      Top             =   1695
      Width           =   1815
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   240
      X2              =   5160
      Y1              =   855
      Y2              =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Look for:"
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
      Left            =   240
      TabIndex        =   4
      Top             =   975
      Width           =   1815
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      X1              =   240
      X2              =   5160
      Y1              =   2415
      Y2              =   2415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   240
      X2              =   5160
      Y1              =   855
      Y2              =   855
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   240
      X2              =   5160
      Y1              =   2415
      Y2              =   2415
   End
End
Attribute VB_Name = "frmFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--------------------------------------------------------------------
'Filter Form
'This form is similar to frmSearch. However, instead of returning the
'absolute position of the matched criteria, it will filter out all the
'data that matches the criteria. It can also handle different data
'types through the use of error trapping.
'
'This code has been reused from the master forms and the Selection
'Dialog Box (frmSelectDg)
'--------------------------------------------------------------------

Public SourceRs            As ADODB.RecordSet

Private Sub Command1_Click()

    On Error GoTo Err
    If Text1.Text = "" Then Text1.SetFocus: Exit Sub
    If Combo1.Text = "" Then Combo1.SetFocus: Exit Sub
    SourceRs.Filter = "[" & Combo1.Text & "] like *" & Text1.Text & "*"
    Unload Me

Exit Sub

Err:
    'If Err.Number = 3001 Then MsgBox "Please select a valid section from the list.", vbExclamation: Text1.Text = "": Combo1.SetFocus: Exit Sub
    'If Err.Number = -2147217825 Then Search_Number: Resume Next: Exit Sub
    Handler Err

End Sub

Private Sub Search_Number()

'For Number data type

    On Error GoTo Err
    SourceRs.Filter = Combo1.Text & " like " & Text1.Text & ""

Exit Sub

Err:
    Search_Date_Time

End Sub

Private Sub Search_Date_Time()

'For Date/Time data type

    On Error GoTo Err
    SourceRs.Filter = Combo1.Text & " like #" & Text1.Text & "#"

Exit Sub

Err:
    MsgBox "Please enter an appropriate value that correspand" & vbCrLf & "where to find it (ex.Search for 10/23/1985 and Look in Date).", vbExclamation

End Sub

Private Sub Command2_Click()
    Set SourceRs = Nothing
    Set frmFilter = Nothing

    Unload Me

End Sub

Private Sub Form_Load()

    Me.Icon = Image1.Picture
    FillCombo Combo1, SourceRs, False
    Combo1.ListIndex = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set SourceRs = Nothing
    Set frmFilter = Nothing

End Sub

Private Sub Image2_Click()

End Sub

