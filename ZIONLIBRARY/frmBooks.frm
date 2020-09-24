VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmBooks 
   ClientHeight    =   7485
   ClientLeft      =   165
   ClientTop       =   -1725
   ClientWidth     =   11700
   Icon            =   "frmBooks.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmBooks.frx":6852
   ScaleHeight     =   7485
   ScaleWidth      =   11700
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdRefresh2 
      Height          =   615
      Left            =   720
      Picture         =   "frmBooks.frx":7616
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   10200
      Width           =   615
   End
   Begin VB.CommandButton cmdReports 
      Caption         =   "Reports"
      Height          =   615
      Left            =   1440
      TabIndex        =   49
      Top             =   10200
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   46
      Top             =   0
      Width           =   8775
      Begin VB.Label Label3 
         DataField       =   "Price"
         Height          =   255
         Left            =   6720
         TabIndex        =   61
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label11 
         Caption         =   "Information of all the books in the library are stored in this table."
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
         Left            =   1440
         TabIndex        =   48
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label Label9 
         Caption         =   "Book Details:"
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
         Left            =   240
         TabIndex        =   47
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdRefresh 
      Height          =   615
      Left            =   0
      Picture         =   "frmBooks.frx":82E0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   10200
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   4200
      TabIndex        =   45
      Top             =   5520
      Width           =   2775
      Begin VB.CommandButton cmdForm 
         Height          =   615
         Left            =   1440
         Picture         =   "frmBooks.frx":8FAA
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdGrid 
         Height          =   615
         Left            =   120
         Picture         =   "frmBooks.frx":9FF4
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame frameDisplay 
      Height          =   615
      Left            =   8880
      TabIndex        =   27
      Top             =   0
      Width           =   2535
      Begin VB.TextBox txtcount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "3333333"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblOF 
         Alignment       =   2  'Center
         Caption         =   "of"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   44
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblmax 
         Alignment       =   2  'Center
         Caption         =   "2222222"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   43
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame frameOperations 
      Height          =   975
      Left            =   6960
      TabIndex        =   26
      Top             =   5520
      Width           =   4455
      Begin VB.CommandButton cmdOperations 
         Height          =   615
         Index           =   2
         Left            =   2520
         Picture         =   "frmBooks.frx":B076
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Sort"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdOperations 
         Height          =   615
         Index           =   1
         Left            =   1920
         Picture         =   "frmBooks.frx":B940
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Filter"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdOperations 
         Height          =   615
         Index           =   0
         Left            =   1320
         Picture         =   "frmBooks.frx":C60A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Search"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdAMod 
         Height          =   615
         Index           =   1
         Left            =   120
         Picture         =   "frmBooks.frx":CED4
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Add "
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   615
         Left            =   3120
         Picture         =   "frmBooks.frx":DB9E
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Delete"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   3720
         Picture         =   "frmBooks.frx":E868
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdAMod 
         Height          =   615
         Index           =   0
         Left            =   720
         Picture         =   "frmBooks.frx":F532
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Edit"
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame frameNav 
      Height          =   975
      Left            =   120
      TabIndex        =   25
      Top             =   5520
      Width           =   4095
      Begin VB.CommandButton cmdNavigate 
         Height          =   615
         Index           =   3
         Left            =   3000
         Picture         =   "frmBooks.frx":101FC
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Last"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdNavigate 
         Height          =   615
         Index           =   2
         Left            =   2040
         Picture         =   "frmBooks.frx":11C60
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Next"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdNavigate 
         Height          =   615
         Index           =   1
         Left            =   1080
         Picture         =   "frmBooks.frx":12A24
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Previous"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdNavigate 
         Height          =   615
         Index           =   0
         Left            =   120
         Picture         =   "frmBooks.frx":137E8
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "First"
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame FrameFormView 
      Caption         =   "Form View"
      Height          =   4815
      Left            =   120
      TabIndex        =   28
      Top             =   600
      Width           =   11295
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   10080
         Top             =   840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdText 
         Caption         =   "Print"
         Height          =   495
         Index           =   4
         Left            =   9840
         TabIndex        =   59
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         Height          =   615
         Left            =   120
         TabIndex        =   56
         Top             =   3960
         Width           =   4815
         Begin VB.TextBox Text1 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
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
            Left            =   1560
            TabIndex        =   57
            Text            =   "Copy books record and save in a file."
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Copy and Save:"
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
            TabIndex        =   58
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.CommandButton cmdText 
         Caption         =   "Clear"
         Height          =   495
         Index           =   1
         Left            =   6240
         TabIndex        =   55
         Top             =   4080
         Width           =   1215
      End
      Begin VB.CommandButton cmdText 
         Caption         =   "Open"
         Height          =   495
         Index           =   2
         Left            =   7440
         TabIndex        =   54
         Top             =   4080
         Width           =   1215
      End
      Begin VB.CommandButton cmdText 
         Caption         =   "Save"
         Height          =   495
         Index           =   3
         Left            =   8640
         TabIndex        =   53
         Top             =   4080
         Width           =   1215
      End
      Begin VB.CommandButton cmdText 
         Caption         =   "Copy"
         Height          =   495
         Index           =   0
         Left            =   5040
         TabIndex        =   52
         Top             =   4080
         Width           =   1215
      End
      Begin VB.TextBox txtStream 
         Height          =   3255
         Left            =   5040
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   51
         Top             =   240
         Width           =   6015
      End
      Begin VB.TextBox txtDisp 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "Publiser Name"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   6
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   3240
         Width           =   3375
      End
      Begin VB.TextBox txtDisp 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "Publiser Name"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   5
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2880
         Width           =   3375
      End
      Begin VB.TextBox txtDisp 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "Publiser Name"
         Height          =   285
         Index           =   4
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2520
         Width           =   3375
      End
      Begin VB.TextBox txtDisp 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "Publiser Name"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   3
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1680
         Width           =   3375
      End
      Begin VB.TextBox txtDisp 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "Author"
         DataSource      =   "adodc1"
         Height          =   285
         Index           =   2
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1320
         Width           =   3375
      End
      Begin VB.TextBox txtDisp 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "Title"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   1
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox txtDisp 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "Book ID"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   0
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   600
         Width           =   3330
      End
      Begin MSComDlg.CommonDialog cdlg 
         Left            =   8520
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label2 
         Height          =   495
         Left            =   2520
         TabIndex        =   60
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Line lnBorder 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   120
         X2              =   11160
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line lnBorder 
         BorderColor     =   &H80000010&
         BorderWidth     =   2
         Index           =   0
         X1              =   120
         X2              =   11160
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line lnBorder 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   120
         X2              =   11160
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line lnBorder 
         BorderColor     =   &H80000010&
         BorderWidth     =   2
         Index           =   3
         X1              =   120
         X2              =   11160
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "ISBN:"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   35
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Price:"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   34
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Category:"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   33
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Publiser Name:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   32
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Author:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   31
         Top             =   1320
         Width           =   1140
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Title:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   30
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Book ID:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   29
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Frame FrameReport 
      Caption         =   "Reports"
      Height          =   4815
      Left            =   120
      TabIndex        =   37
      Top             =   600
      Width           =   11295
      Begin VB.CommandButton cmdReport 
         Height          =   615
         Index           =   2
         Left            =   600
         Picture         =   "frmBooks.frx":1524C
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   3240
         Width           =   615
      End
      Begin VB.CommandButton cmdReport 
         Height          =   615
         Index           =   0
         Left            =   600
         Picture         =   "frmBooks.frx":15F16
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton cmdReport 
         Height          =   615
         Index           =   1
         Left            =   600
         Picture         =   "frmBooks.frx":16BE0
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label19 
         Caption         =   "Create Complete Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   41
         Top             =   720
         Width           =   4215
      End
      Begin VB.Label Label20 
         Caption         =   "Create a complete report on all the books that are in the library. The Grid View will show the complete inventory."
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
         Left            =   1440
         TabIndex        =   40
         Top             =   960
         Width           =   4095
      End
      Begin VB.Label Label21 
         Caption         =   "Create Custom Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   39
         Top             =   2040
         Width           =   4215
      End
      Begin VB.Label Label22 
         Caption         =   $"frmBooks.frx":178AA
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1440
         TabIndex        =   38
         Top             =   2280
         Width           =   4095
      End
   End
   Begin VB.Frame FrameGridView 
      Caption         =   "Grid View"
      Height          =   4815
      Left            =   120
      TabIndex        =   36
      Top             =   600
      Width           =   11295
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4455
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   10815
         _ExtentX        =   19076
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
         Caption         =   "Book Details"
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
End
Attribute VB_Name = "frmBooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-------------------------------------------
'Book Table Master Form
'Manupulates the Books Table, providing different views and the ability
'to create dynamic reports. Adding / Editing is done by Books Add/Edit
'Form.
'------------------------------------------
Private blank As String
Private RS As ADODB.RecordSet

Private Sub cmdAMod_Click(Index As Integer)
    'Open the add/edit form. Display current record values in form if modifying.
   
    On Error Resume Next
    
    OldPosition = RS.AbsolutePosition
    
    With frmMain
    
    'if ............not admin
    If .StatusBar1.Panels(4).Text <> "Administrator" Then
    MsgBox "Sorry! only an administrator is allowed to update the books database.", vbExclamation, _
    "Database update denied!"
    
    Else 'if admin
    
        With frmBooksAE
            .AddState = Index
            .OldID = RS.Fields(0) 'bookID column
            If Index = 0 Then 'Edit Mode
                .msdID.Text = RS.Fields(0)
                .txtTitle.Text = RS.Fields(1)
                .txtAuthor.Text = RS.Fields(2)
                .txtPublisher.Text = RS.Fields(3)
                .cboCategory.Text = RS.Fields(4)
                .txtPrice.Text = RS.Fields(5)
                .msdISBN.Text = RS.Fields(6)
            End If
            .Show vbModal
        End With
        cmdRefresh_Click
        DisplayRecords
    End If
    End With
End Sub



Private Sub cmdClose_Click()
Set frmBooks = Nothing
Set RS = Nothing
Unload Me
End Sub
Private Sub cmdDelete_Click()
'Deletes a record, undeletable if a book is borrowed.

Dim ans As Integer, pos As Integer

    On Error GoTo hell
    
    
    With frmMain 'x1
    'if ............not admin
    
    If .StatusBar1.Panels(4).Text <> "Administrator" Then  '1
    MsgBox "Sorry! only an administrator is allowed to update the books database.", vbExclamation, _
    "Database update denied!"
    
    
    Else 'if admin
    
    With RS 'x2
        If .RecordCount > 1 Then 'last record deletion trap 2
        
        'Check whether book is borrowed 3
        If .Fields("Borrowed") = True Then MsgBox "You cannot delete this book record because it is borrowed by someone" & vbNewLine & "The book must be returned to the library before its record can be deleted.", vbInformation, "Book Borrowed"
    
        'confirm deletion of record
         ans = MsgBox("Are you sure you want to delete the selected record?", vbCritical + vbYesNo, "Confirm Record Deletion")
    
            
            If ans = vbYes Then 'actual deletion happens within this line ----------------------
            'Delete the record
            pos = .AbsolutePosition
            CN.BeginTrans
            .Delete
            .Requery
            CN.CommitTrans
                If pos > .RecordCount Then
                    If Not .EOF Or .BOF Then .MoveFirst
                Else
                .AbsolutePosition = pos
                End If
                cmdRefresh2_Click
                MsgBox "Record has been successfully deleted.", vbInformation, "Confirm"
            End If 'actual deletion happens within this line---------------------------
    
        Else
            'if no record to delete...
            'If .RecordCount < 1 Then MsgBox "No record to delete.", vbExclamation: Exit Sub
            
            'If .RecordCount = 1 Then
            
            'if record is the last
            MsgBox "The system does not support last record deletion.", vbInformation, "deletion denied" 'last record cannot be deleted...
          
            'End If
        End If 'pair of last rec deletion trap..2
      End With 'x2
    End If 'pair of admin..1
   
    End With 'pair of with frmmain..x1
Exit Sub

hell:
    On Error Resume Next
        Handler Err
        CN.RollbackTrans
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFilter_Click()

End Sub

Private Sub cmdForm_Click()
    FrameFormView.Visible = True
    FrameGridView.Visible = False
    FrameReport.Visible = False
    'font of the buttons set
    cmdGrid.FontBold = False
    cmdForm.FontBold = True
    cmdReports.FontBold = False
End Sub
Private Sub cmdGrid_Click()
    FrameFormView.Visible = False
    FrameGridView.Visible = True
    FrameReport.Visible = False
    'font of the buttons set
    cmdGrid.FontBold = True
    cmdForm.FontBold = False
    cmdReports.FontBold = False
End Sub

Private Sub cmdNavigate_Click(Index As Integer)
    Navigate Index, RS
    DisplayRecords
End Sub
Private Sub cmdOperations_Click(Index As Integer)
'Create new instances of Search/Sort/Filter forms and display them. Destroy when done with

Dim obj As Form

    If Index = 0 Then Set obj = frmSearch
    If Index = 1 Then Set obj = frmFilter
    If Index = 2 Then Set obj = frmSort

    With obj
        Set .SourceRs = RS
        .Show vbModal
    End With
    Set obj = Nothing

End Sub



Private Sub cmdRefresh2_Click()
   On Error Resume Next
    With RS
        .Filter = adFilterNone
        .Requery
    End With
    DisplayRecords
End Sub

Private Sub cmdReport_Click(Index As Integer)
'Create dynamic reports

    If Index = 0 Then 'cmdRefresh_Click
    Set rptBookList.DataSource = RS
    rptBookList.Show
    End If
End Sub

Private Sub cmdReports_Click()
    FrameFormView.Visible = False
    FrameGridView.Visible = False
    FrameReport.Visible = True
    'font of the buttons set
    cmdGrid.FontBold = False
    cmdForm.FontBold = False
    cmdReports.FontBold = True

End Sub

Private Sub cmdText_Click(Index As Integer)
     If Index = 0 Then
     Dim str As String
  
    
     '------------------------
     str = "Book ID:" & " " & txtDisp(0).Text & vbCrLf & "Title:" & " " & txtDisp(1).Text & vbCrLf & "Author:" & " " & txtDisp(2).Text & vbCrLf & "Publishers Name:" & " " & txtDisp(3).Text & vbCrLf & "Category:" & " " & txtDisp(4).Text & vbCrLf & "Price:" & " " & txtDisp(5).Text & vbCrLf & "ISBN:" & " " & txtDisp(6).Text
     
     txtStream.Text = txtStream.Text & vbCrLf & str & vbCrLf & "________________________________________________"
     '------------------------
    cmdText(4).Enabled = True
    cmdText(3).Enabled = True
    cmdText(1).Enabled = True
    cmdText(2).Enabled = False
    
    ElseIf Index = 1 Then
    txtStream.Text = ""
    cmdText(4).Enabled = False
    cmdText(3).Enabled = False
    cmdText(1).Enabled = False
    cmdText(2).Enabled = True
    
    ElseIf Index = 2 Then
   
        
        cdlg.Filename = ""
        cdlg.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
        cdlg.ShowOpen
        
        
        
        txtStream = LoadText(cdlg.FileTitle)
    
        '=========================
        'If txtStream.Text <> "" Then
        'cmdText(4).Enabled = True
        'cmdText(3).Enabled = True
        'cmdText(1).Enabled = True
        'cmdText(2).Enabled = False
        'Else
        'cmdText(4).Enabled = False
        'cmdText(3).Enabled = False
        'cmdText(1).Enabled = False
        'cmdText(2).Enabled = True
        'End If
        '==========================
    
    ElseIf Index = 3 Then
    cdlg.ShowSave
    SaveText txtStream, cdlg.FileTitle & ".txt"
    Else
    CommonDialog1.ShowPrinter
    End If
    'CommonDialog1

    
End Sub



Private Sub Form_Load()
    blank = txtStream.Text
    cmdGrid.FontBold = True
    FrameFormView.Visible = False
    FrameGridView.Visible = True
    FrameReport.Visible = False
    
    With frmMain
    .LastElement = .LastElement + 1
    .cmdCounter_Click
        
    End With
'---------------------------------
'disable 2 menu coammands
'With frmMain
 '   .mnucreateuser.Enabled = False
  '  .mnuLogoff.Enabled = False
'End With
cmdText(4).Enabled = False
cmdText(3).Enabled = False
cmdText(1).Enabled = False

'---------------------------------
'Create recordset and refresh. Link Report icons to ImageList

    On Error GoTo hell
    Set RS = New ADODB.RecordSet
    RS.CursorLocation = adUseClient
    RS.Open "SELECT * FROM tblBooks", CN, adOpenDynamic, adLockOptimistic
    
    Set DataGrid1.DataSource = RS
    DisplayRecords
    
Exit Sub

hell:
    Handler Err
    Resume Next

End Sub

Private Sub DisplayRecords()

'Display the current and total number of record

Dim i As Integer

    On Error Resume Next
        With RS
            If .RecordCount < 1 Then
                txtcount.Text = 0
            Else
                txtcount.Text = .AbsolutePosition
            End If
            lblmax.Caption = .RecordCount

            For i = 0 To 6
                txtDisp(i).Text = .Fields(i)
            Next i
        End With
        txtDisp(5).Text = FormatCurrency$(txtDisp(5).Text)

End Sub
Public Sub cmdRefresh_Click() 'intentionally made this public to be accessible to othe forms
'like i have to use this when i add a new record..to update the display of the recordcount
'Refresh the recordset
    On Error Resume Next
    With RS
        .Filter = adFilterNone
        .Requery
     '------------------------
        If AddSuccess = True Then
        .MoveLast
         'AddSuccess = False
        
        Else
        .AbsolutePosition = OldPosition
        
        End If
        '------------------------
    
    
    End With
    DisplayRecords

End Sub

Private Sub Form_Unload(Cancel As Integer)
    With frmMain
    .LastElement = .LastElement - 1
    .cmdCounter_Click
    End With
    Set frmBooks = Nothing
    Set RS = Nothing
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    DisplayRecords
End Sub

'Private Sub Sum()
 '   Dim RS As RecordSet
    
  '  Set RS = New ADODB.RecordSet
   ' RS.CursorLocation = adUseClient
    'RS.Open "SELECT SUM(Price) FROM tblBooks", CN, adOpenDynamic, adLockOptimistic
    
   'Set DataGrid2.DataSource = RS
   
    
'End Sub



