VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10c.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   ClientHeight    =   7035
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8880
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   1058
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      _Version        =   393216
      Begin VB.CommandButton cmdCounter 
         Height          =   615
         Left            =   -720
         TabIndex        =   13
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton cmdBorrow 
         Enabled         =   0   'False
         Height          =   615
         Left            =   120
         Picture         =   "frmMain.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Issue"
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton cmdReturn 
         Enabled         =   0   'False
         Height          =   615
         Left            =   960
         Picture         =   "frmMain.frx":807E
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Return"
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton cmdBook 
         Enabled         =   0   'False
         Height          =   615
         Left            =   1800
         Picture         =   "frmMain.frx":98AA
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Book List"
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton cmdAdmin 
         Enabled         =   0   'False
         Height          =   615
         Left            =   2640
         Picture         =   "frmMain.frx":B0D6
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Members List"
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton cmdReport 
         Enabled         =   0   'False
         Height          =   615
         Left            =   3480
         Picture         =   "frmMain.frx":C902
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Reports"
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton cmdSetting 
         Enabled         =   0   'False
         Height          =   615
         Left            =   4320
         Picture         =   "frmMain.frx":E12E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Settings"
         Top             =   0
         Width           =   800
      End
      Begin VB.Frame Frame1 
         Height          =   580
         Left            =   6000
         TabIndex        =   10
         Top             =   -45
         Width           =   8505
         Begin VB.Label Label1 
            Height          =   255
            Left            =   600
            TabIndex        =   15
            Top             =   240
            Width           =   4215
         End
         Begin VB.Label control_num 
            Height          =   375
            Left            =   3480
            TabIndex        =   12
            Top             =   120
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblstatus 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   360
            TabIndex        =   11
            Top             =   240
            Width           =   45
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   14520
         TabIndex        =   9
         Top             =   60
         Width           =   660
      End
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
         Height          =   405
         Left            =   14565
         TabIndex        =   8
         ToolTipText     =   "Connected"
         Top             =   105
         Width           =   570
         _cx             =   1005
         _cy             =   714
         FlashVars       =   ""
         Movie           =   "D:\Video Rental System (VRS)\Images\LOADING.SWF"
         Src             =   "D:\Video Rental System (VRS)\Images\LOADING.SWF"
         WMode           =   "Window"
         Play            =   "-1"
         Loop            =   "-1"
         Quality         =   "High"
         SAlign          =   ""
         Menu            =   "-1"
         Base            =   ""
         AllowScriptAccess=   "always"
         Scale           =   "ShowAll"
         DeviceFont      =   "0"
         EmbedMovie      =   "0"
         BGColor         =   ""
         SWRemote        =   ""
         MovieData       =   ""
         SeamlessTabbing =   "1"
         Profile         =   "0"
         ProfileAddress  =   ""
         ProfilePort     =   "0"
         AllowNetworking =   "all"
         AllowFullScreen =   "false"
      End
      Begin VB.CommandButton cmdHelp 
         Enabled         =   0   'False
         Height          =   615
         Left            =   5160
         Picture         =   "frmMain.frx":F95A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Help"
         Top             =   0
         Width           =   800
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   14
      Top             =   6705
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   13
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   2408
            MinWidth        =   1235
            Text            =   "User Name"
            TextSave        =   "User Name"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2118
            MinWidth        =   2118
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   1746
            MinWidth        =   1587
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2824
            MinWidth        =   2824
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "Date"
            TextSave        =   "Date"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1931
            MinWidth        =   1941
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "Time Log-in"
            TextSave        =   "Time Log-in"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   18
            MinWidth        =   18
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Bevel           =   0
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel12 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Bevel           =   0
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel13 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Bevel           =   0
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "INS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSettings 
         Caption         =   "&Settings"
      End
      Begin VB.Menu mnucreateuser 
         Caption         =   "Create user account"
      End
      Begin VB.Menu mnulock 
         Caption         =   "Lock Application"
      End
      Begin VB.Menu mnuLogoff 
         Caption         =   "Log-off"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuTrans 
      Caption         =   "Tr&ansaction"
      Begin VB.Menu mnuIssue 
         Caption         =   "Issue"
      End
      Begin VB.Menu mnuReturn 
         Caption         =   "Return"
      End
   End
   Begin VB.Menu mnuRec 
      Caption         =   "&Records"
      Begin VB.Menu mnuBookRec 
         Caption         =   "Books Record"
      End
      Begin VB.Menu mnuMemRec 
         Caption         =   "Members Record"
      End
      Begin VB.Menu mnuBorrowedBooks 
         Caption         =   "Borrowed Books"
      End
      Begin VB.Menu mnuHistory 
         Caption         =   "Borrowing History"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "Reports"
      Begin VB.Menu mnuBookrep 
         Caption         =   "Books Report"
      End
      Begin VB.Menu mnuMemRep 
         Caption         =   "Members Report"
      End
      Begin VB.Menu mnuBorrowedRep 
         Caption         =   "Borrowed Books Report"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnuCas 
         Caption         =   "Cascade"
      End
      Begin VB.Menu mnuTileHor 
         Caption         =   "Tile Horizontally"
      End
      Begin VB.Menu mnuTileVer 
         Caption         =   "Tile Vertically"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuZionHelp 
         Caption         =   "Zion Library Help..."
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "Report"
      Visible         =   0   'False
      Begin VB.Menu mnuBookreport 
         Caption         =   "Books Report"
      End
      Begin VB.Menu mnuMembersreport 
         Caption         =   "Members Report"
      End
      Begin VB.Menu mnuBorrowed 
         Caption         =   "Borrowed Books Report"
      End
   End
   Begin VB.Menu mnuHelp2 
      Caption         =   "Help"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Reentry As Boolean
Public LastElement  As Integer
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
                    ByVal hWnd As Long, _
                    ByVal lpOperation As String, _
                    ByVal lpFile As String, _
                    ByVal lpParameters As String, _
                    ByVal lpDirectory As String, _
                    ByVal nShowCmd As Long) As Long

Private Const SW_HIDE As Long = 0
Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const SW_SHOWMINIMIZED As Long = 2

Private Sub cmdAdmin_Click()
   frmMembers.Show
End Sub

Private Sub cmdBook_Click()
    frmBooks.Show
End Sub

Private Sub cmdBorrow_Click()
    If StatusBar1.Panels(4).Text <> "Administrator" And StatusBar1.Panels(4).Text <> "Assistant" Then
    MsgBox "Only Administrator or Assistant is allowed to use this feature.", vbExclamation, "Access denied"
    Else
    frmIssue.Show vbModal
    End If
End Sub

Public Sub cmdCounter_Click()
    '********************************************
    '********************************************
    'Frame1.Caption = LastElement
    '*************************
    '---------------------------------
    'disable menu commands and enable back the commands when 'LastElement' variable
    'counts back to zero
    'made this command button public so that it can be clicked from other forms
    'somehow our exercise exer9A help us got the algorithm of this feature..
    'to count the open child forms,then enables back when the count is zero..
    If LastElement <> 0 Then
        mnucreateuser.Enabled = False
        mnuLogoff.Enabled = False
        mnuAbout.Enabled = False
      
  
   
        mnulock.Enabled = False
        Label1.Caption = "Please close tables to enable  menu commands."
    Else
        mnucreateuser.Enabled = True
        mnuLogoff.Enabled = True
        mnuAbout.Enabled = True
   
      
    
        mnulock.Enabled = True
        Label1.Caption = ""
    End If
    If LastElement = 1 Then
        StatusBar1.Panels(9).Text = "  1  table is open."
    ElseIf LastElement = 2 Then
        StatusBar1.Panels(9).Text = "  2  tables  are open."
    ElseIf LastElement = 3 Then
        StatusBar1.Panels(9).Text = "  3  tables  are open."
    ElseIf LastElement = 4 Then
        StatusBar1.Panels(9).Text = "  4  tables  are open."
    Else
        StatusBar1.Panels(9).Text = "  No  table is open."
    End If
'---------------------------------
End Sub

Private Sub cmdHelp_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    'reference visual basic 6,how to program, by deitel & deitel
    
     frmAbout.Show vbModal
End Sub

Private Sub cmdReport_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    'reference visual basic 6,how to program, by deitel & deitel
    
    If Button = vbLeftButton Then
        Call PopupMenu(mnuReports)
    End If
End Sub

Private Sub cmdReturn_Click()
    If StatusBar1.Panels(4).Text <> "Administrator" And StatusBar1.Panels(4).Text <> "Assistant" Then
    MsgBox "Only Administrator or Assistant is allowed to use this feature.", vbExclamation, "Access denied"
    Else
    frmReturn.Show vbModal
    End If
End Sub

Private Sub cmdSetting_Click()
    If StatusBar1.Panels(4).Text <> "Administrator" Then
    MsgBox "Sorry! only an administrator is allowed to use this feature.", vbExclamation, _
    "User creation denied!"
    Else 'if admin
    frmSettings.Show
    End If
End Sub

Private Sub MDIForm_Load()
  
    Me.Show
    Set CN = New ADODB.Connection
    CN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\MasterFile.mdb;Persist Security Info=False;Jet OLEDB:Database Password=MLEVDQ48L2"
    'CN.Close
    If CN.State <> adStateOpen Then MsgBox "Could not establish a connection with the database" & vbNewLine & vbNewLine & "The database should exist in ApplicationPath\MasterFile.mdb", vbExclamation, "Database not found!": 'Unload Me
    'if something goes wrong and connection could not be opened..
    'used msgbox icon constant vbexclamation ref. how to program by deitel pp. 440
    'if connection is not established, then unload the main form too
    'it will be lousy to see an open form with no connection
    'my datagrid and other controls would be blank...blah blah
    '---------------------------------------------------------------
    'loads the Windows registry value saved using function savesetting
    'as early as loading this form...
    'see SaveSetting details on frmreturn...
    frmReturn.FineAmnt = CCur(GetSetting(App.Title, "Settings", "Fine Amount", "5"))
    frmReturn.MaxDays = CInt(GetSetting(App.Title, "Settings", "Max Days", "3"))
    
    'frmReturn.FineAmnt = CCur(GetSetting(App.Title, "Settings", "Fine Amount"))
    'frmReturn.MaxDays = CInt(GetSetting(App.Title, "Settings", "Max Days"))
    
    Reentry = False
    
    LastElement = 0
    cmdCounter_Click
'---------------------------------
    
   
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If MsgBox("This will terminate the application. Proceed?", vbOKCancel + vbQuestion, "Terminate Application") = vbOK Then
   Set frmMain = Nothing
   CN.Close
  '--------------------------------------
  'set the screen reso back...
  SetOrigRes
  '--------------------------------------
   End
   Else
   Cancel = 1
   End If
End Sub



Private Sub mnuAbout_Click()
  frmAbout.Show vbModal

End Sub

Private Sub mnuBookRec_Click()
 
    
    
    With frmBooks
        .Show
       
    End With
    
End Sub

Private Sub mnuBookRep_Click()
    Dim RS As New ADODB.RecordSet

    'RS.Open "SELECT * FROM tblBooks Order by [Book ID]", CN, adOpenStatic, adLockReadOnly
    RS.Open "SELECT * FROM tblBooks Order by [Book ID]", CN, adOpenStatic, adLockReadOnly
    Set rptBookList.DataSource = RS
    rptBookList.Show
    Set RS = Nothing
  
End Sub

Private Sub mnuBorrowedRec_Click()
    With frmBorrowed
        .Show
       
    End With
       
End Sub

Private Sub mnuBorrowedBooks_Click()
    
    frmBorrowed.Show
End Sub

Private Sub mnuBorrowedRep_Click()
Dim RS As New ADODB.RecordSet

    RS.Open "SELECT tblTrans.[Book ID], tblTrans.[Student ID], tblBooks.Title, [First Name] & ' ' & [Middle Initial] & ' ' & [Last Name] AS Borrower, tblTrans.[Date Borrowed] FROM tblMembers INNER JOIN (tblBooks INNER JOIN tblTrans ON tblBooks.[Book ID] = tblTrans.[Book ID]) ON tblMembers.[Student ID] = tblTrans.[Student ID] Where (((tblTrans.Returned) = False)) ORDER BY tblTrans.[Book ID];", CN, adOpenStatic, adLockReadOnly
    Set rptTrans.DataSource = RS
    rptTrans.Show
    Set RS = Nothing
    '-----------------------------------------------------------------------
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
    

End Sub

Private Sub mnuCas_Click()
    frmMain.Arrange vbCascade
     'ref. exer9A & exer9B
End Sub

Private Sub mnucreateuser_Click()
    'if ............not admin
    If StatusBar1.Panels(4).Text <> "Administrator" Then
    MsgBox "Sorry! only an administrator is allowed to use this feature.", vbExclamation, _
    "User creation denied!"
    Else 'if admin
    
    frmadduser.Show vbModal
    
    End If
    
End Sub



Private Sub mnuExit_Click()
  If MsgBox("This will terminate the application. Proceed?", vbYesNo + vbQuestion, "Terminate Application") = vbYes Then
   Set frmMain = Nothing
   CN.Close
    '--------------------------------------
  'set the screen reso back...
   SetOrigRes
  '--------------------------------------
   End
  
   End If
End Sub

Private Sub mnuHistory_Click()
frmHistory.Show
End Sub

Private Sub mnuIssue_Click()
    If StatusBar1.Panels(4).Text <> "Administrator" And StatusBar1.Panels(4).Text <> "Assistant" Then
    MsgBox "Only Administrator or Assistant is allowed to use this feature.", vbExclamation, "Access denied"
    Else
    frmIssue.Show vbModal
    End If
End Sub

Private Sub mnulock_Click()
   StatusBar1.Panels(2).Text = "Waiting..."
   StatusBar1.Panels(4).Text = "Waiting..."
  frmLock.Show vbModal
  
End Sub

Private Sub mnuLogoff_Click()
    Reentry = True
    
    If Reentry = True Then
    StatusBar1.Panels(2).Text = "Waiting..."
    StatusBar1.Panels(4).Text = "Waiting..."
    StatusBar1.Panels(8).Text = Time
    StatusBar1.Panels(6).Text = Date
    '-------------------------------------
        frmLogin.cmdRefresh_Click
        'Command Buttons; setting to enabled = false
        'before Acess is granted to user...
        '-----------------------
        DisableCommandButtons
        '-----------------------
        
        'Menu Commands; setting to enabled = false
        'before acces is granted to user...
        
        DisableMenuCommands
                        
    '-------------------------------------
    End If
    frmLogin.Show vbModal
       
End Sub

Private Sub mnuMemRec_Click()
    
    frmMembers.Show
End Sub

Private Sub mnuMemRep_Click()
    Dim RS As New ADODB.RecordSet

    RS.Open "SELECT * FROM tblMembers Order by [Student ID]", CN, adOpenStatic, adLockReadOnly
    Set rptMembers.DataSource = RS
    rptMembers.Show
    Set RS = Nothing

End Sub

Private Sub mnuReturn_Click()
    If StatusBar1.Panels(4).Text <> "Administrator" And StatusBar1.Panels(4).Text <> "Assistant" Then
    MsgBox "Only Administrator or Assistant is allowed to use this feature.", vbExclamation, "Access denied"
    Else
    frmReturn.Show vbModal
    End If
End Sub

Private Sub mnuSettings_Click()
    If StatusBar1.Panels(4).Text <> "Administrator" Then
    MsgBox "Sorry! only an administrator is allowed to use this feature.", vbExclamation, _
    "User creation denied!"
    Else 'if admin
    frmSettings.Show vbModal
    End If
End Sub
Private Sub mnuBookreport_click()
    Dim RS As New ADODB.RecordSet

    RS.Open "SELECT * FROM tblBooks Order by [Book ID]", CN, adOpenStatic, adLockReadOnly
    Set rptBookList.DataSource = RS
    rptBookList.Show
    Set RS = Nothing
  '---------------------
End Sub
Private Sub mnuMembersreport_click()
    Dim RS As New ADODB.RecordSet

    RS.Open "SELECT * FROM tblMembers Order by [Student ID]", CN, adOpenStatic, adLockReadOnly
    Set rptMembers.DataSource = RS
    rptMembers.Show
    Set RS = Nothing

End Sub

Private Sub mnuBorrowed_click()
    Dim RS As New ADODB.RecordSet

    RS.Open "SELECT tblTrans.[Book ID], tblTrans.[Student ID], tblBooks.Title, [First Name] & ' ' & [Middle Initial] & ' ' & [Last Name] AS Borrower, tblTrans.[Date Borrowed] FROM tblMembers INNER JOIN (tblBooks INNER JOIN tblTrans ON tblBooks.[Book ID] = tblTrans.[Book ID]) ON tblMembers.[Student ID] = tblTrans.[Student ID] Where (((tblTrans.Returned) = False)) ORDER BY tblTrans.[Book ID];", CN, adOpenStatic, adLockReadOnly
    Set rptTrans.DataSource = RS
    rptTrans.Show
    Set RS = Nothing
    
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
    
End Sub


Private Sub mnuAbout2_click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuTileHor_Click()
    frmMain.Arrange vbTileHorizontal
    'ref. exer9A & exer9B
End Sub

Private Sub mnuTileVer_Click()
    frmMain.Arrange vbTileVertical
     'ref. exer9A & exer9B
End Sub
Private Sub mnuZionHelp_Click()
'Dim WordApp As Word.Application
 
 
'Set WordApp = CreateObject("word.Application")

'WordApp.Visible = True
'-----------
'WordApp.Documents.Open "D:\Program Files\ZIONLIBRARY\Von.DOC"
'-----------



'WordApp.Documents.Open App.Path & "\Von.DOC"
'-----------
 RunMe
    
End Sub

Private Sub RunMe()
     'ShellExecute Me.hWnd, "open", "C:\WINDOWS\Von.pdf", vbNullString, "C:\WINDOWS", SW_SHOWNORMAL
     ShellExecute Me.hWnd, "open", "C:\WINDOWS\UserManual.pdf", vbNullString, "C:\WINDOWS", SW_SHOWNORMAL

End Sub

