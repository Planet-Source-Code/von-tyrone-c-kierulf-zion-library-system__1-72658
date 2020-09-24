VERSION 5.00
Begin VB.Form frmReturn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Books Return Form"
   ClientHeight    =   5820
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5325
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReturn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtBookAmount 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   3120
      Width           =   3015
   End
   Begin VB.TextBox txtBookPayable 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "$0.00"
      Top             =   3480
      Width           =   3015
   End
   Begin VB.ComboBox cboBookStatus 
      Height          =   315
      ItemData        =   "frmReturn.frx":6852
      Left            =   1800
      List            =   "frmReturn.frx":6862
      TabIndex        =   22
      Text            =   "In Good Condition"
      Top             =   2760
      Width           =   3375
   End
   Begin VB.TextBox txtDateRet 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2040
      Width           =   3375
   End
   Begin VB.TextBox txtFines 
      Height          =   285
      Left            =   1800
      TabIndex        =   16
      Top             =   2400
      Width           =   3375
   End
   Begin VB.CommandButton Command4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4920
      Picture         =   "frmReturn.frx":68A3
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Use calculator."
      Top             =   3480
      Width           =   315
   End
   Begin VB.TextBox txtStudentCode 
      BackColor       =   &H00F4FEFF&
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1680
      Width           =   3375
   End
   Begin VB.TextBox txtBookID 
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1320
      Width           =   3015
   End
   Begin VB.CommandButton cmdCode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4920
      Picture         =   "frmReturn.frx":6CF5
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Browse record."
      Top             =   1320
      Width           =   315
   End
   Begin VB.Frame Frame1 
      Caption         =   "Info Panel"
      Height          =   1215
      Left            =   240
      TabIndex        =   4
      Top             =   3840
      Width           =   4935
      Begin VB.Label Label4 
         Caption         =   "Date borrowed:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Days late in returning the book:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "Total amount of fine accumulated:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label lblDate 
         Caption         =   "Select a book first"
         Height          =   255
         Left            =   3120
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblLate 
         Caption         =   "Select a book first"
         Height          =   255
         Left            =   3120
         TabIndex        =   6
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblFines 
         Caption         =   "Select a book first"
         Height          =   255
         Left            =   3120
         TabIndex        =   5
         Top             =   840
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "R&eturn Book"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   "Book Amount:"
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Book Payable:"
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Book Status:"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmReturn.frx":71E7
      Height          =   735
      Index           =   1
      Left            =   840
      TabIndex        =   20
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label7 
      Caption         =   "Date Returned:"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Fines collected:"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Student Code:"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   120
      X2              =   5280
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label2 
      Caption         =   "Book ID:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmReturn.frx":72A0
      Stretch         =   -1  'True
      Top             =   240
      Width           =   480
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   5280
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   240
      X2              =   5160
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   1
      X1              =   240
      X2              =   5160
      Y1              =   5160
      Y2              =   5160
   End
End
Attribute VB_Name = "frmReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'----------------------------------------------------------------------
'Books Return Form
'This form is used to return books that have been borrowed from the library
'and recorded in the database via the issue form.
'More details in cmdCode_Click and cmdReturn_Click procedures.
'-----------------------------------------------------------------------
Public MaxDays As Integer
Public FineAmnt As Currency
Private Sub cmdcancel_Click()
    Set frmReturn = Nothing

    Unload Me

End Sub

Private Sub cmdReset_Click()

    lblLate.Caption = "Select a book first"
    lblFines.Caption = "Select a book first"
    lblDate.Caption = "Select a book first"
    txtFines.Text = ""
    txtFines.Locked = True
    txtStudentCode.Text = ""
    txtBookID.Text = ""
    txtDateRet.Text = FormatDateTime$(Date, vbLongDate)

End Sub

Private Sub cmdReturn_Click()

    Dim RS As ADODB.RecordSet

    'What it does here is pretty easy. The return information is recorded
    'in two places. One in the Book Table where the book Borrowed is set to
    'False, and in the Transaction Table where the amount payed and book
    'returned is stored with the date the book is returned.

    If txtBookID.Text = "" And txtStudentCode.Text = "" And txtFines.Text = "" Then txtBookID.SetFocus
    '-----------
  
    '-----------
    On Error GoTo hell
    Set RS = New ADODB.RecordSet
    With RS
        CN.BeginTrans       'Begin a new transaction
        .Open "Select [Borrowed] from tblBooks where [Book ID]='" & txtBookID.Text & "'", CN, adOpenDynamic, adLockOptimistic
        'select borrowed field of tblbooks
        '( provided bookId matches with the value in txtbookID )
        
        .MoveFirst
        .Fields(0) = False 'unchecks the borrowed field in tblBooks
      
        .Update
        .Close

        .Open "Select [Date Returned],[Librarian Retrieved],[Book Status],[Book Payable],[Fines],[Returned] From tblTrans where [Book ID]='" & txtBookID.Text & "'" & "And [Returned] = False", CN, adOpenDynamic, adLockOptimistic
        .MoveFirst
        .Fields("Date Returned") = Time + Date
        .Fields("Librarian Retrieved") = frmMain.StatusBar1.Panels(2).Text
        '===================
        If cboBookStatus.Text = "Lost" Then
     
     
        .Fields("Book Payable") = CCur(txtBookAmount.Text)
        Else
        .Fields("Fines") = CCur(txtFines.Text)
        .Fields("Book Payable") = CCur(txtBookPayable.Text)
        End If
        '===================
        .Fields("Book Status") = cboBookStatus.Text
      
        .Fields("Returned") = True
        .Update
        .Close
        CN.CommitTrans      'If no error was raised then record info
        frmBorrowed.cmdRefresh_Click
    End With
    Set RS = Nothing

    'Show MsgBox if another book needs returning
    If MsgBox("The book " & txtBookID.Text & " has been returned from " & txtStudentCode.Text & vbNewLine & vbCrLf & "Do you want to create a new return book instance?", vbInformation + vbYesNo) = vbYes Then
        cmdReset_Click
        
    Else
        Unload Me
    End If

Exit Sub

hell:
    Handler Err

    On Error Resume Next    'If an error was raised then rollback
        CN.RollbackTrans        'any transaction so GIGO does not take place
                            'in the future.

End Sub

Private Sub cmdCode_Click()

Dim RS As ADODB.RecordSet, i As Integer

    'The first part of this event procedure will open the frmSelectDg form
    'and expect an input from the user. This will ease the selection part
    'from the users point-of-view and validation part from the devolopers
    'point-of-view.

    On Error Resume Next
        With frmSelectDg
            'First show the box
            .CommandText = "SELECT tblTrans.[Book ID], tblTrans.[Student ID], tblBooks.Title, [First Name] & ' ' & [Middle Initial] & ' ' & [Last Name] AS Borrower, tblTrans.[Date Borrowed] FROM tblMembers INNER JOIN (tblBooks INNER JOIN tblTrans ON tblBooks.[Book ID] = tblTrans.[Book ID]) ON tblMembers.[Student ID] = tblTrans.[Student ID] Where (((tblTrans.Returned) = False)) ORDER BY tblTrans.[Book ID];"
            .DataGrid1.Caption = "Borrowed Books Table"
            .Show vbModal

            'Now display the data
            If .OKPressed Then
                txtBookID.Text = .rRS1
                txtStudentCode.Text = .rRS2
                'txtBookAmount.Text = .BookAmount
                txtFines.Locked = False
            Else
                'If the user did not enter anything then skip the second
                'part of the procedure to skip errors that may arise because
                'there will be no data (in txtBookID and txtStudentCode) and as such
                'null errors or record not found errors.
                Exit Sub
            End If
        End With

        'The second part will calculate the number of days a book was taken out
        'of the library and print it in the txtFines text box.

        Set RS = New ADODB.RecordSet
        RS.Open "Select * from tblTrans Where [Book ID] ='" & txtBookID.Text & "'", CN, adOpenDynamic, adLockOptimistic
        lblDate.Caption = CDate(RS(2))      'Just for validation

        'Store the difference of the current date and the date returned
        'in a variable. It the variable is negative it means that the
        'book returned is within the time limit and Fines=i*FineAmnt
        'must be 0. So transform i into 0
        i = Date - CDate(lblDate.Caption) 'date the book is borrowed
        If i < 0 Then i = 0
        If MaxDays < i Then lblLate.Caption = i - MaxDays Else lblLate.Caption = "0"

        'Print fines due in a label and a text box
        lblFines.Caption = CStr(FormatCurrency$(FineAmnt * lblLate))

        'Also, use an editable text box so the correct amount a member
        'is payed is recorded. Sometimes the member may pay money not
        'exactly as required (payable $15 from $15.25 total fines)
        txtFines.Text = lblFines.Caption
        txtBookAmount.Text = CStr(FormatCurrency$(RS.Fields(6)))
        'txtBookPayable.Text = CStr(FormatCurrency$(txtBookPayable.Text))
        'txtTotalPayable.Text = CStr(FormatCurrency$(txtTotalPayable.Text))
        Set RS = Nothing

        'So, practically all the librarian did was just select a book id through
        'a GUI friendly interface and everything will be done by the system

End Sub

Private Sub Command4_Click()

    On Error GoTo hell
    'Shell "calc.exe", vbNormalFocus
    MathCalc.Show vbModal
    Me.Icon = Command4.Picture
    '----------------------------------
    ' I Intend to use my own calculator...so that when i compute the total fines.
    'i can assign the amount to txtFines for saving ( user dont need to enter an amount )
    '...user-friendly, is'nt it ?
    ' il try to do most of the things that i've learned in our class..
    'and many more things that i've learned outside of our class

    
Exit Sub

hell:
    'MsgBox "The operating system cannot find the system calculator." & vbNewLine & "Please check whether it is properly installed or not", vbCritical, "File not found"

End Sub

Private Sub Form_Load()
   
    
    cmdReset_Click
    

End Sub

Private Sub Text4_Keypress(KeyAscii As Integer)

    cmdCode_Click

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmReturn = Nothing
  
End Sub
