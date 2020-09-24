VERSION 5.00
Begin VB.Form frmIssue 
   Caption         =   "Issue"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5535
   ClipControls    =   0   'False
   Icon            =   "frmIssue.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   5535
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtBookAmount 
      BackColor       =   &H8000000E&
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2760
      Width           =   2895
   End
   Begin VB.CommandButton cmdIssue 
      Caption         =   "&Issue Book"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   9
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox txtStudName 
      BackColor       =   &H00F4FEFF&
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1680
      Width           =   2895
   End
   Begin VB.TextBox txtBookTitle 
      BackColor       =   &H00F4FEFF&
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2400
      Width           =   2895
   End
   Begin VB.TextBox txtDateIssued 
      BackColor       =   &H8000000E&
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   3120
      Width           =   2895
   End
   Begin VB.TextBox txtStudCode 
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox txtBookID 
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2040
      Width           =   2535
   End
   Begin VB.CommandButton cmdStudent 
      Height          =   315
      Left            =   4890
      Picture         =   "frmIssue.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Browse record."
      Top             =   1320
      Width           =   315
   End
   Begin VB.CommandButton cmdBook 
      Height          =   315
      Left            =   4890
      Picture         =   "frmIssue.frx":6D44
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Browse for record"
      Top             =   2040
      Width           =   315
   End
   Begin VB.TextBox txtDateRet 
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   3480
      Width           =   2895
   End
   Begin VB.Label Label7 
      Caption         =   "Book Amount:"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   0
      X1              =   240
      X2              =   5160
      Y1              =   975
      Y2              =   975
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   240
      X2              =   5160
      Y1              =   975
      Y2              =   975
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   240
      X2              =   5280
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label Label1 
      Caption         =   "Student Code:"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Book ID:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   16
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmIssue.frx":7236
      Stretch         =   -1  'True
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmIssue.frx":8A62
      Height          =   855
      Index           =   1
      Left            =   960
      TabIndex        =   15
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label3 
      Caption         =   "Student Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Book Title:"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Date Issued:"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Date to be returned:"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   1
      X1              =   240
      X2              =   5280
      Y1              =   3960
      Y2              =   3960
   End
End
Attribute VB_Name = "frmIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim datereturned As Date
Private Sub cmdBook_Click()
        '------------------------------------------------------------------------
        'Using ( WITH...END WITH ) structure to have a convenient syntax and to
        'make CODE easier to read and make it run faster...
        'refer to page 488 of visual basic 6,how to program by deitel
        'or to page 645 of the visual basic 6 complete
        'this part open the form frmselectdg just like windows style browsing
        'then gets using an SQL refering to tblbooks of books that is not borrowed
        '------------------------------------------------------------------------

    With frmSelectDg
        
        .CommandText = "Select * From tblBooks where Borrowed=False"
        .DataGrid1.Caption = "Books Table"
        .Show vbModal
        '-------------------------------------------------------------------------
        'uses vbmodal so that the user cant interact with any other forms while
        'the form frmSlectdg is not yet closed
        'used to haver it the hard way...coding 1 kilometer
        'setting other forms enable prop to false while the form is active
        '-------------------------------------------------------------------------
        If .OKPressed Then
            txtBookID.Text = .rRS1     ' variable that receives value from the field book ID in tblBooks
            txtBookTitle.Text = .rRS2  'variable that receives value from the field Title in tblbooks
            txtBookAmount.Text = CCur(.BookAmount)
        End If
    End With
End Sub

Private Sub cmdcancel_Click()
  Set frmIssue = Nothing
  Set RS = Nothing
  Unload Me
End Sub

Private Sub cmdIssue_Click()
    '-----------------------------------------------------------------
    'Record that the book was taken in two places. In tblTrans, and in
    'tblBooks which will set the Borrowed Boolean to True.
    '-----------------------------------------------------------------
Dim RS As ADODB.RecordSet

    If txtStudCode.Text = "" Then txtStudCode.SetFocus: Exit Sub
    If txtBookID.Text = "" Then txtBookID.SetFocus: Exit Sub
    On Error GoTo hell
    CN.BeginTrans
    Set RS = New ADODB.RecordSet
    With RS
        .Open "Select * from tblTrans", CN, adOpenDynamic, adLockOptimistic
        .AddNew
        .Fields(0) = txtBookID.Text
        .Fields(1) = txtStudCode.Text
        .Fields(2) = Time + Date
        .Fields(3) = frmMain.StatusBar1.Panels(2).Text
        .Fields(6) = CCur(txtBookAmount.Text)
        .Update
        .Close

        'this part checks the borrowed field from tblbooks
        .Open "Select [Borrowed] from tblBooks where [Book ID]='" & txtBookID.Text & "'", CN, adOpenDynamic, adLockOptimistic
        
        .MoveFirst
        .Fields(0) = True
        .Update
        .Close
        Set RS = Nothing
    End With
    CN.CommitTrans
    frmBorrowed.cmdRefresh_Click
    If MsgBox("The book " & txtBookID.Text & " has been issued to " & txtStudCode.Text & vbNewLine & vbNewLine & "Do you want to create a new issue instance?", vbInformation + vbYesNo) = vbYes Then
        cmdReset_Click
    Else
        Unload Me
    End If
    
Exit Sub

hell:
    Handler Err
    CN.RollbackTrans

End Sub

Private Sub cmdReset_Click()
    txtStudCode.Text = ""
    txtStudName.Text = ""
    txtBookID.Text = ""
    txtBookTitle.Text = ""
    txtBookAmount.Text = ""
    txtDateIssued.Text = FormatDateTime$(Date, vbLongDate)
   'txtDateIssued.Text = FormatDateTime$(Now(), vbLongDate)
    txtDateRet.Text = FormatDateTime$(Date + frmReturn.MaxDays, vbLongDate)

End Sub

Private Sub cmdStudent_Click()

Dim A As String, b As String, c As String

    With frmSelectDg
        .CommandText = "Select * From tblMembers"
        .DataGrid1.Caption = "Members Table"
        .Show vbModal
        If .OKPressed Then
            txtStudCode.Text = .rRS1
            A = .rRS2
            b = .rRS3
            c = .rRS4
            txtStudName.Text = A & " " & b & " " & c
        End If
    End With
End Sub

Private Sub Form_Load()
    'Image1.Picture = frmMain.Icon
     cmdReset_Click
    'Me.Icon = Image1.Picture

    'SendKeys "{enter}"
    'txtDateIssued.Text = FormatDateTime$(Date, vbLongDate)

    'datereturned = Date + 3
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set frmIssue = Nothing
  Set RS = Nothing
End Sub
