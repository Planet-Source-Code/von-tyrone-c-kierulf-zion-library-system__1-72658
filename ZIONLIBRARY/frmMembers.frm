VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMembers 
   Caption         =   "members"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11700
   Icon            =   "frmMembers.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7710
   ScaleWidth      =   11700
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdRefresh2 
      Height          =   615
      Left            =   720
      Picture         =   "frmMembers.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   10320
      Width           =   615
   End
   Begin VB.Frame frameDisplay 
      Height          =   615
      Left            =   8880
      TabIndex        =   38
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
         TabIndex        =   39
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
         TabIndex        =   41
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblmax 
         Alignment       =   2  'Center
         Caption         =   "5555555"
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
         TabIndex        =   40
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   120
      TabIndex        =   35
      Top             =   0
      Width           =   8775
      Begin VB.Label Label11 
         Caption         =   "Information of all the members of the Library are stored here."
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
         Left            =   1560
         TabIndex        =   37
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label Label9 
         Caption         =   "Member Details:"
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
         TabIndex        =   36
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   4200
      TabIndex        =   32
      Top             =   5520
      Width           =   2775
      Begin VB.CommandButton cmdForm 
         Height          =   615
         Left            =   1440
         Picture         =   "frmMembers.frx":751C
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdGrid 
         Height          =   615
         Left            =   120
         Picture         =   "frmMembers.frx":8566
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRefresh 
      Height          =   615
      Left            =   0
      Picture         =   "frmMembers.frx":95E8
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   10320
      Width           =   615
   End
   Begin VB.Frame frameNav 
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   5520
      Width           =   4095
      Begin VB.CommandButton cmdNavigate 
         Height          =   615
         Index           =   0
         Left            =   120
         Picture         =   "frmMembers.frx":A2B2
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "First"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdNavigate 
         Height          =   615
         Index           =   1
         Left            =   1080
         Picture         =   "frmMembers.frx":BD16
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Previous"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdNavigate 
         Height          =   615
         Index           =   2
         Left            =   2040
         Picture         =   "frmMembers.frx":CADA
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Next"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdNavigate 
         Height          =   615
         Index           =   3
         Left            =   3000
         Picture         =   "frmMembers.frx":D89E
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Last"
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame frameOperations 
      Height          =   975
      Left            =   6960
      TabIndex        =   3
      Top             =   5520
      Width           =   4455
      Begin VB.CommandButton cmdAMod 
         Height          =   615
         Index           =   0
         Left            =   720
         Picture         =   "frmMembers.frx":F302
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Edit"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   3720
         Picture         =   "frmMembers.frx":FFCC
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   615
         Left            =   3120
         Picture         =   "frmMembers.frx":10C96
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Delete"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdAMod 
         Height          =   615
         Index           =   1
         Left            =   120
         Picture         =   "frmMembers.frx":11960
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Add"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdOperations 
         Height          =   615
         Index           =   0
         Left            =   1320
         Picture         =   "frmMembers.frx":1262A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Search"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdOperations 
         Height          =   615
         Index           =   1
         Left            =   1920
         Picture         =   "frmMembers.frx":12EF4
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Filter"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdOperations 
         Height          =   615
         Index           =   2
         Left            =   2520
         Picture         =   "frmMembers.frx":13BBE
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Sort"
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame FrameFormView 
      Caption         =   "Form"
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   11295
      Begin VB.PictureBox Picture1 
         Height          =   3135
         Left            =   6120
         ScaleHeight     =   3075
         ScaleWidth      =   2475
         TabIndex        =   42
         Top             =   480
         Width           =   2535
         Begin VB.Image Image1 
            DataField       =   "Picture"
            Height          =   3135
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   2535
         End
      End
      Begin VB.TextBox txtDisp 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   0
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtDisp 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   1
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox txtDisp 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   2
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtDisp 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   3
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1680
         Width           =   3015
      End
      Begin VB.TextBox txtDisp 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   4
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox txtDisp 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   5
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox txtDisp 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   6
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   3120
         Width           =   615
      End
      Begin VB.Line lnBorder 
         BorderColor     =   &H80000010&
         BorderWidth     =   2
         Index           =   3
         X1              =   120
         X2              =   9600
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line lnBorder 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   120
         X2              =   9600
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line lnBorder 
         BorderColor     =   &H80000010&
         BorderWidth     =   2
         Index           =   0
         X1              =   120
         X2              =   9600
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line lnBorder 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   120
         X2              =   9600
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label Label4 
         Caption         =   "Student Code:"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "First Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "Middle Initial:"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "Last Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "Class:"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label17 
         Caption         =   "Section:"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label18 
         Caption         =   "Roll:"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label Label19 
         Caption         =   "Picture:"
         Height          =   255
         Left            =   6120
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame FrameGridView 
      Caption         =   "Grid View"
      Height          =   4815
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   11295
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4335
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   7646
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
End
Attribute VB_Name = "frmMembers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------
'Members Master Form
'this form manupulates the Members Table, except the Adding/Editing which
'are done by Members Add/Edit Form
'------------------------------------------------------------------------
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Type OPENFILENAME
    Von As Long
  
End Type
'------------------------------------------------------------------------
Private RS As ADODB.RecordSet
Private Sub cmdAMod_Click(Index As Integer)

    On Error Resume Next
  
    OldPosition = RS.AbsolutePosition
    
    With frmMain
    
    'if ............not admin
    If .StatusBar1.Panels(4).Text <> "Administrator" Then
    MsgBox "Sorry! only an administrator is allowed to update the books database.", vbExclamation, _
    "Database update denied!"
    
    Else 'if admin
       
        With frmMembersAE
            .AddState = Index
            .OldID = RS.Fields(0)
            If Index = 0 Then
                .txtCode.Text = RS(0)
                .txtFirst.Text = RS(1)
                .txtM.Text = RS(2)
                .txtLast.Text = RS(3)
                .cboClass.Text = RS(4)
                .cboSection = RS(5)
                .txtRoll = RS(6)
            End If
            .Show vbModal
        End With
   End If
   End With
        cmdRefresh_Click
        DisplayRecords

End Sub

Private Sub cmdDelete_Click()
    'Deletes a record
    On Error GoTo hell
    With frmMain
    'if ............not admin
    If .StatusBar1.Panels(4).Text <> "Administrator" Then
    MsgBox "Sorry! only an administrator is allowed to update the books database.", vbExclamation, _
    "Database update denied!"
    
    Else 'if admin
    With RS
        '-Check if there is no record
        If .RecordCount < 1 Then MsgBox "No record to delete.", vbExclamation: Exit Sub
        '-Confirm deletion of record

        Dim ans As Integer, pos As Integer
        ans = MsgBox("Are you sure you want to delete the selected record?", vbCritical + vbYesNo, "Confirm Record Deletion")
        Screen.MousePointer = vbHourglass
        If ans = vbYes Then
            '-Delete the record
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
        End If
        Screen.MousePointer = vbDefault
    End With
    End If
    End With
Exit Sub

hell:
    Handler Err
    CN.RollbackTrans

End Sub

Private Sub cmdForm_Click()
    FrameFormView.Visible = True
    FrameGridView.Visible = False
   
    'font of the buttons set
    cmdGrid.FontBold = False
    cmdForm.FontBold = True
 
End Sub

Private Sub cmdGrid_Click()
    FrameFormView.Visible = False
    FrameGridView.Visible = True
    
    'font of the buttons set
    cmdGrid.FontBold = True
    cmdForm.FontBold = False
  
End Sub

Private Sub cmdNavigate_Click(Index As Integer)
    Navigate Index, RS
    DisplayRecords
    '----------------------
   
End Sub

Private Sub cmdOperations_Click(Index As Integer)
    'Shows the Search/Sort/Filter form by creating a new instance and destroying it once done
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

Public Sub cmdRefresh_Click()
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
End Sub

Private Sub cmdReport_Click(Index As Integer)

End Sub

Private Sub cmdReports_Click()
   
End Sub

Private Sub cmdSelect_Click()

End Sub


Private Sub cmdRefresh2_Click()
   On Error Resume Next
     With RS
        .Filter = adFilterNone
        .Requery
        '------------------------
        DisplayRecords
     End With
End Sub

Private Sub DataGrid1_GotFocus()
On Error Resume Next
End Sub

Private Sub DataGrid1_LostFocus()
On Error Resume Next
End Sub

Private Sub Form_Load()
    '----------------
    'cmdRetrive_Click
     With frmMain
     .LastElement = .LastElement + 1
     .cmdCounter_Click
    End With
    '----------------
    cmdGrid.FontBold = True
    FrameFormView.Visible = False
    FrameGridView.Visible = True
   
  
'------------------------------------------
'---------------------------------
'disable 2 menu coammands
'With frmMain
 '   .mnucreateuser.Enabled = False
  '  .mnuLogoff.Enabled = False
'End With
'---------------------------------
'Loads a form and initializes all variables
    On Error GoTo hell
    Set RS = New ADODB.RecordSet
    RS.CursorLocation = adUseClient
    RS.Open "SELECT * FROM tblMembers", CN, adOpenDynamic, adLockOptimistic
    Set DataGrid1.DataSource = RS
    Set Image1.DataSource = RS
    DisplayRecords
    
Exit Sub

hell:
    Handler Err
    Resume Next

End Sub
Private Sub cmdClose_Click()
 Set frmMembers = Nothing
 Set RS = Nothing
 Unload Me
End Sub
Private Sub DataGrid1_DblClick()
cmdSelect_Click
End Sub

Private Sub DisplayRecords()

'-Display the current and total number of record

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

End Sub


Private Sub Form_Unload(Cancel As Integer)
    With frmMain
    .LastElement = .LastElement - 1
    .cmdCounter_Click
    End With
'---------------------------------
    'Destroys variables to free memory
    Set RS = Nothing
    Set frmMembers = Nothing

End Sub
Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    DisplayRecords
End Sub




