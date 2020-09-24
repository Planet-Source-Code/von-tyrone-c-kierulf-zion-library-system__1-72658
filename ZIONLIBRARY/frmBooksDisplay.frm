VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBooksDisplay 
   Caption         =   "Form1"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6570
   ScaleWidth      =   11700
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
      Height          =   495
      Left            =   120
      TabIndex        =   44
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton cmdPidlit 
      Height          =   495
      Left            =   1560
      TabIndex        =   43
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Frame frameOperations 
      Height          =   975
      Left            =   6240
      TabIndex        =   34
      Top             =   5520
      Width           =   5055
      Begin VB.CommandButton cmdEdit 
         Height          =   615
         Left            =   720
         Picture         =   "frmBooksDisplay.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   4320
         Picture         =   "frmBooksDisplay.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdRefresh 
         Height          =   615
         Left            =   3720
         Picture         =   "frmBooksDisplay.frx":1994
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   615
         Left            =   3120
         Picture         =   "frmBooksDisplay.frx":265E
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         Height          =   615
         Left            =   120
         Picture         =   "frmBooksDisplay.frx":3328
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdSearch 
         Height          =   615
         Left            =   1320
         Picture         =   "frmBooksDisplay.frx":3FF2
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdFilter 
         Height          =   615
         Left            =   1920
         Picture         =   "frmBooksDisplay.frx":48BC
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdSort 
         Height          =   615
         Left            =   2520
         Picture         =   "frmBooksDisplay.frx":5586
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame frameDisplay 
      Height          =   975
      Left            =   4200
      TabIndex        =   32
      Top             =   5520
      Width           =   2175
      Begin VB.Label lblDisplay 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame frameNav 
      Height          =   975
      Left            =   120
      TabIndex        =   27
      Top             =   5520
      Width           =   4095
      Begin VB.CommandButton cmdFirst 
         Height          =   615
         Left            =   120
         Picture         =   "frmBooksDisplay.frx":5E50
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   615
         Left            =   1080
         Picture         =   "frmBooksDisplay.frx":6D58
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdNext 
         Height          =   615
         Left            =   2040
         Picture         =   "frmBooksDisplay.frx":7C60
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdLast 
         Height          =   615
         Left            =   3000
         Picture         =   "frmBooksDisplay.frx":8B68
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdForm 
      Caption         =   "Form View"
      Height          =   375
      Left            =   1320
      TabIndex        =   17
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdGrid 
      Caption         =   "Grid View"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdReports 
      Caption         =   "Reports"
      Height          =   375
      Left            =   2520
      TabIndex        =   15
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame FrameFormView 
      Caption         =   "Form View"
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   11175
      Begin VB.TextBox txtBookID 
         DataField       =   "Book ID"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1485
         TabIndex        =   7
         Top             =   360
         Width           =   3330
      End
      Begin VB.TextBox txtBookTitle 
         DataField       =   "Title"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1485
         TabIndex        =   6
         Top             =   720
         Width           =   3375
      End
      Begin VB.TextBox txtAuthor 
         DataField       =   "Author"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1485
         TabIndex        =   5
         Top             =   1125
         Width           =   3375
      End
      Begin VB.TextBox txtPubliserName 
         DataField       =   "Publiser Name"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   0
         Left            =   1485
         TabIndex        =   4
         Top             =   1500
         Width           =   3375
      End
      Begin VB.TextBox txtCategory 
         DataField       =   "Category"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   0
         Left            =   1485
         TabIndex        =   3
         Top             =   1875
         Width           =   3375
      End
      Begin VB.TextBox txtPrice 
         DataField       =   "Price"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   0
         Left            =   1485
         TabIndex        =   2
         Top             =   2265
         Width           =   3360
      End
      Begin VB.TextBox txtISBN 
         DataField       =   "ISBN"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   0
         Left            =   1485
         TabIndex        =   1
         Top             =   2640
         Width           =   3345
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Book ID:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   405
         Width           =   1095
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Title:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   780
         Width           =   1095
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Author:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   1170
         Width           =   1140
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Publiser Name:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   11
         Top             =   1590
         Width           =   1095
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Category:"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   10
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Price:"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   9
         Top             =   2310
         Width           =   1095
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "ISBN:"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   8
         Top             =   2685
         Width           =   1095
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   0
      Top             =   7680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Program Files\ZIONLIBRARY\database\masterfile.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Program Files\ZIONLIBRARY\database\masterfile.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tblBooks"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame FrameGridView 
      Caption         =   "Grid View"
      Height          =   4815
      Left            =   120
      TabIndex        =   25
      Top             =   600
      Width           =   11175
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmBooksDisplay.frx":9A70
         Height          =   4455
         Left            =   120
         TabIndex        =   26
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
   Begin VB.Frame FrameReport 
      Caption         =   "Reports"
      Height          =   4815
      Left            =   120
      TabIndex        =   18
      Top             =   600
      Width           =   11175
      Begin VB.CommandButton cmdReport 
         Height          =   615
         Index           =   1
         Left            =   600
         Picture         =   "frmBooksDisplay.frx":9A85
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2640
         Width           =   615
      End
      Begin VB.CommandButton cmdReport 
         Height          =   615
         Index           =   2
         Left            =   600
         Picture         =   "frmBooksDisplay.frx":A74F
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label22 
         Caption         =   $"frmBooksDisplay.frx":B419
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
         TabIndex        =   24
         Top             =   2880
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
         TabIndex        =   23
         Top             =   2640
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
         TabIndex        =   22
         Top             =   1680
         Width           =   4095
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
         TabIndex        =   21
         Top             =   1440
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmBooksDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdAdd_Click()
frmBooksAE.Show
Unload Me
End Sub

Private Sub cmdAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblDisplay.Caption = "Add New"
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblDisplay.Caption = "Close"
End Sub



Private Sub cmdDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblDisplay.Caption = "Delete"
End Sub

Private Sub cmdDeleteXXX_Click()

End Sub

Private Sub cmdEdit_Click()
FrameFormView.Visible = True
FrameGridView.Visible = False
FrameReport.Visible = False
'font of the buttons set
cmdGrid.FontBold = False
cmdForm.FontBold = True
cmdReports.FontBold = False

'Adodc1.Recordset.Update
End Sub

Private Sub cmdEdit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblDisplay.Caption = "Edit Record"
End Sub



Private Sub cmdFilter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblDisplay.Caption = "Filter"
End Sub

Private Sub cmdExit_Click()
End Sub

Private Sub cmdFirst_Click()
Adodc1.Recordset.MoveFirst
End Sub

Private Sub cmdFirst_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblDisplay.Caption = "First"
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

Private Sub cmdLast_Click()
Adodc1.Recordset.MoveLast
End Sub

Private Sub cmdLast_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblDisplay.Caption = "Last"
End Sub

Private Sub cmdNext_Click()
Adodc1.Recordset.MoveNext
'XXXXXXXXXXXXXXX


If Adodc1.Recordset.EOF Then
    Adodc1.Recordset.MoveLast
End If


End Sub

Private Sub cmdNext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblDisplay.Caption = "Next"
End Sub

Private Sub cmdPidlit_Click()
Adodc1.Recordset.MoveLast
End Sub

Private Sub cmdPrevious_Click()
Adodc1.Recordset.MovePrevious
'XXXXXXXXXXXX
 
    If Adodc1.Recordset.BOF Then
    Adodc1.Recordset.MoveFirst
    
    End If
    
End Sub

Private Sub cmdPrevious_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblDisplay.Caption = "Previous"
End Sub

Private Sub cmdRefresh_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblDisplay.Caption = "Refresh"
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



Private Sub cmdSearch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblDisplay.Caption = "Search"
End Sub

Private Sub cmdSelect_Click()
frmIssue.Visible = True
'frmIssue.Show
frmIssue.txtBookID.Text = frmBooks.txtBookID.Text
frmIssue.txtBookTitle.Text = frmBooks.txtBookTitle.Text
Adodc1.Recordset.Update
Unload Me
End Sub

Private Sub cmdSort_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblDisplay.Caption = "Sort"
End Sub




Private Sub Form_Load()
cmdGrid.FontBold = True
FrameFormView.Visible = False
FrameGridView.Visible = True
FrameReport.Visible = False
End Sub

Private Sub frameDisplay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblDisplay.Caption = ""
End Sub


'Private Sub adodc1_validate(Action As Integer, Save As Integer)
 '   If cmdAdd.Caption = "Add" Then
    
    
  '      If Save = -1 Then
   '         If MsgBox("Save changes?", vbYesNo) = vbYes Then
    
    '         Save = -1
    
     '        Else
      '       Save = 0
       '      End If
       ' End If
    
    'End If
    
'End Sub

