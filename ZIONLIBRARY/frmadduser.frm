VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmadduser 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "Create Account"
   ClientHeight    =   6495
   ClientLeft      =   2805
   ClientTop       =   1500
   ClientWidth     =   6405
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2655
      Left            =   120
      TabIndex        =   11
      Top             =   3600
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   4683
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00404040&
      Height          =   3855
      Left            =   5040
      TabIndex        =   20
      Top             =   2520
      Width           =   1215
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   3240
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   960
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   2640
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Caption         =   "Verify your password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   960
      Left            =   120
      TabIndex        =   13
      Top             =   2520
      Width           =   2760
      Begin VB.TextBox txtPass2 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   135
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   825
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Pls. Fillup the field"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1860
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   6120
      Begin VB.ComboBox cboLevel 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmadduser.frx":0000
         Left            =   3120
         List            =   "frmadduser.frx":000D
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   2400
      End
      Begin VB.TextBox txtUserID 
         Enabled         =   0   'False
         Height          =   345
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   0
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtPass 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   3120
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1320
         Width           =   2385
      End
      Begin VB.TextBox txtUser 
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User ID #"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   810
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status/Level"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   3120
         TabIndex        =   17
         Top             =   360
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   3120
         TabIndex        =   15
         Top             =   1080
         Width           =   825
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   855
      End
   End
   Begin TabDlg.SSTab stab 
      Height          =   555
      Left            =   120
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   240
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   979
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   758
      BackColor       =   -2147483636
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "User"
      TabPicture(0)   =   "frmadduser.frx":0032
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      BorderWidth     =   9
      Height          =   6495
      Left            =   0
      Top             =   0
      Width           =   6375
   End
End
Attribute VB_Name = "frmadduser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As ADODB.RecordSet
Dim AddState As Boolean

Private Sub cmdAdd_Click()
    
    On Error Resume Next
   
    
    If RS.RecordCount = 0 Then
    
    txtUserID.Text = "001"
    
    Else
        
        RS.MoveLast
        txtUserID.Text = Format(RS!UserId + 1, "000")
        '------------------------
        EnableSaveCancel 'enable save & cancel button; disable add button
   
        '------------------------
        Enable_Controls 'enable the controls
        '------------------------
        txtUser.SetFocus
    
    End If
    
End Sub

Private Sub cmdcancel_Click()
    On Error Resume Next
    EnableAdd 'enable add button ; disable Save & Cancel button
    '------------------
    Disable_Controls 'disable controls
    Clear_Controls 'clear fields
End Sub

Private Sub cmdDelete_Click()
    Dim pos As Integer
    On Error GoTo hell
    With RS
    If .RecordCount = 1 Then 'if last record,system does not delete
    MsgBox "Record deletion Denied!" & vbCrLf & "The system does not support last record deletion.", vbExclamation, "Deletion Denied"
    
    Else
        If MsgBox("Are you sure you want to delete the selected record?", vbYesNo) = vbYes Then
             
            'Delete the record
            
             pos = .AbsolutePosition
             CN.BeginTrans
             
            .Delete
            .Requery
             Set RS = Nothing
             CN.CommitTrans
            
            MsgBox "Record has been successfully deleted.", vbInformation, "Deletion confirmed"
       
            
          
        
            
        End If
        
    End If
    End With
hell:
    On Error Resume Next
        Handler Err
        CN.RollbackTrans

End Sub

Private Sub cmdExit_Click()
    Set frmadduser = Nothing
    Set RS = Nothing
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
    
    Set RS = New ADODB.RecordSet
    RS.CursorLocation = adUseClient
    RS.Open "SELECT UserID,User,UserLevel FROM tblAccess", CN, adOpenStatic, adLockOptimistic
    
    Set DataGrid1.DataSource = RS
    RS.Requery
    'Set RS = Nothing
    

    
End Sub

Private Sub cmdSave_Click()
    
    On Error GoTo hell
  
    If txtUserID = "" Or cboLevel = "" Or txtUser = "" Or txtPass = "" Then '1 blank fields
    MsgBox ("Missing Data, please complete the Fields!"), vbInformation, "Invalid"
    
    
    Else 'fields all field up..
        If txtPass.Text <> txtPass2.Text Then '2 password mismatch
        MsgBox "Password you entered does not match", vbInformation, "Invalid"
        txtPass.Text = ""
        txtPass2.Text = ""
        txtPass.SetFocus
        
        Else 'evrythins fine,proceed adding..
        Set RS = New ADODB.RecordSet
        RS.CursorLocation = adUseClient
        RS.Open "SELECT * FROM tblAccess", CN, adOpenStatic, adLockOptimistic
        CN.BeginTrans
        With RS
        .AddNew
        .Fields(0) = txtUserID.Text
        .Fields(1) = txtUser.Text
        .Fields(2) = txtPass.Text
        .Fields(3) = cboLevel.Text
        .Update
        .Close
        Set RS = Nothing
        CN.CommitTrans
        cmdRefresh_Click
        
        MsgBox txtUser.Text + " " + "is now a registered user.", vbInformation, "Registration confirmed"
        'cmdAdd.SetFocus
            If MsgBox("Do you wish to add another user?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
            EnableAdd
            cmdAdd_Click
            txtUser = ""
            txtPass = ""
            txtPass2 = ""
        
            Else
            'Clear_Controls
            'Disable_Controls
            'EnableAdd
            Unload Me
            frmadduser.Show vbModal
            End If
        'Clear_Controls
        'Disable_Controls
        'EnableAdd
        'Exit Sub
        End With
        
        End If '2
    End If '1
    
hell:
        
        On Error Resume Next
        Handler Err
        'CN.RollbackTrans
   

End Sub

Private Sub Form_Load()
    
    cmdRefresh_Click
End Sub

Sub Clear_Controls()
    txtUser.Text = ""
    txtPass.Text = ""
    txtPass2.Text = ""
    txtUserID.Text = ""
    cboLevel.Text = ""
End Sub


Sub Enable_Controls()
    txtUserID.Enabled = True
    txtUser.Enabled = True
    txtPass.Enabled = True
    txtPass2.Enabled = True
    cboLevel.Enabled = True
End Sub

Sub Disable_Controls()
    txtUserID.Enabled = False
    txtUser.Enabled = False
    txtPass.Enabled = False
    txtPass2.Enabled = False
    cboLevel.Enabled = False

End Sub

Sub EnableAdd()
    cmdAdd.Enabled = True
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
End Sub

Sub EnableSaveCancel()
    cmdAdd.Enabled = False
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
End Sub


