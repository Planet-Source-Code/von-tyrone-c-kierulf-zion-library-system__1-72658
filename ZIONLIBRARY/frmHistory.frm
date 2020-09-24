VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmHistory 
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11700
   Icon            =   "frmHistory.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7485
   ScaleWidth      =   11700
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdRefresh 
      Height          =   615
      Left            =   120
      Picture         =   "frmHistory.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7680
      Width           =   615
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   9600
      TabIndex        =   10
      Top             =   5640
      Width           =   855
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   120
         Picture         =   "frmHistory.frx":751C
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame frameNav 
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   5640
      Width           =   4095
      Begin VB.CommandButton cmdNavigate 
         Height          =   615
         Index           =   3
         Left            =   3000
         Picture         =   "frmHistory.frx":81E6
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Last"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdNavigate 
         Height          =   615
         Index           =   2
         Left            =   2040
         Picture         =   "frmHistory.frx":9C4A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Next"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdNavigate 
         Height          =   615
         Index           =   1
         Left            =   1080
         Picture         =   "frmHistory.frx":AA0E
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Previous"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdNavigate 
         Height          =   615
         Index           =   0
         Left            =   120
         Picture         =   "frmHistory.frx":B7D2
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "First"
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4815
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   11295
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4455
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   7858
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      Begin VB.Label Label9 
         Caption         =   "Borrowing History Details:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label11 
         Caption         =   "Information of  books borrowings  in the library are stored in this table."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   1
         Top             =   240
         Width           =   5655
      End
   End
End
Attribute VB_Name = "frmHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private RS As ADODB.RecordSet


Private Sub cmdClose_Click()
Set frmHistory = Nothing
Set RS = Nothing
Unload Me
End Sub

Public Sub cmdRefresh_Click() 'intentionally made this public to be accessible to othe forms
'like i have to use this when i issue or return a book to update the display of the table
'Refresh the recordset
    On Error Resume Next
    With RS
        .Filter = adFilterNone
        .Requery
    End With
    'DisplayRecords

End Sub


Private Sub Form_Load()
     
     
    On Error GoTo hell
    Set RS = New ADODB.RecordSet
    RS.CursorLocation = adUseClient
    'RS.Open "SELECT tblTrans.[Book ID], tblTrans.[Student ID], tblBooks.Title, [First Name] & ' ' & [Middle Initial] & ' ' & [Last Name] AS Borrower, tblTrans.[Date Borrowed] FROM tblMembers INNER JOIN (tblBooks INNER JOIN tblTrans ON tblBooks.[Book ID] = tblTrans.[Book ID]) ON tblMembers.[Student ID] = tblTrans.[Student ID] Where (((tblTrans.Returned) = False)) ORDER BY tblTrans.[Book ID];", CN, adOpenDynamic, adLockOptimistic
     RS.Open "SELECT tblTrans.[Book ID], tblTrans.[Student ID], tblBooks.Title, [First Name] & ' ' & [Middle Initial] & ' ' & [Last Name] AS Borrower, tblTrans.[Date Borrowed],tblTrans.[Librarian Issued],tblTrans.[Date Returned],tblTrans.[Librarian Retrieved],[Book Amount],[Book Status],[Book Payable],[Fines] FROM tblMembers INNER JOIN (tblBooks INNER JOIN tblTrans ON tblBooks.[Book ID] = tblTrans.[Book ID]) ON tblMembers.[Student ID] = tblTrans.[Student ID] Where (((tblTrans.Returned) = true)) ORDER BY tblTrans.[Date Borrowed];", CN, adOpenDynamic, adLockOptimistic
    'RS.Open "SELECT tblTrans.[Book ID], tblTrans.[Student ID], tblBooks.Title, [First Name] & ' ' & [Middle Initial] & ' ' & [Last Name] AS Borrower, tblTrans.[Date Borrowed],tblTrans.[Librarian Issued],tblTrans.[Date Returned],tblTrans.[Librarian Retrieved] FROM tblMembers INNER JOIN (tblBooks INNER JOIN tblTrans ON tblBooks.[Book ID] = tblTrans.[Book ID]) ON tblMembers.[Student ID] = tblTrans.[Student ID] Where (((tblTrans.Returned) = true)) ORDER BY tblTrans.[Book ID];", CN, adOpenDynamic, adLockOptimistic
    Set DataGrid1.DataSource = RS
    '-----------------------------------
     With frmMain
    .LastElement = .LastElement + 1
    .cmdCounter_Click
    End With
    '--------------------
    'Ref. pp. 782-784 of  Visual Basic 6, How to program, By Deitel & Deitel
    'the query must select bookid and studentid from tbltrans..
    'then title from tblbooks...
    'then create a temporary column Borrower...
    'the columns, of course, must be from left to right..
    'tbltrans,tblbooks and tblmembers column be merged on the condition that:
    'first, tblbooks.bookid and tbltrans.bookid matches...
    'second,tblmembes.studentid and tbltrans.studentid also matches...
    'most of all, tbltrans.returned must be false...
    'then of course display in the ORDER BY bookid...
    
    
Exit Sub

hell:
    Handler Err
    Resume Next
End Sub

Private Sub cmdNavigate_Click(Index As Integer)
 Navigate Index, RS

End Sub

Private Sub Form_Unload(Cancel As Integer)
     With frmMain
    .LastElement = .LastElement - 1
    .cmdCounter_Click
    End With
    Set frmHistory = Nothing
    Set RS = Nothing
End Sub

