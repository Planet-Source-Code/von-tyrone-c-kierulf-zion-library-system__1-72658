VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   2535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5550
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2535
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   4200
      Width           =   1575
   End
   Begin MCI.MMControl MMControl1 
      Height          =   735
      Left            =   600
      TabIndex        =   6
      Top             =   2880
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   1296
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   1785
      Width           =   1215
   End
   Begin VB.TextBox txtpass 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox txtuser 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   510
      TabIndex        =   5
      Top             =   945
      Width           =   1350
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   510
      TabIndex        =   4
      Top             =   1305
      Width           =   1245
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   50
      Height          =   2655
      Left            =   600
      Shape           =   1  'Square
      Top             =   8040
      Width           =   9735
   End
   Begin VB.Image Image2 
      Height          =   765
      Left            =   840
      Top             =   7920
      Width           =   5250
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   50
      Height          =   3255
      Left            =   3120
      Shape           =   3  'Circle
      Top             =   7680
      Width           =   4695
   End
   Begin VB.Image Image3 
      Height          =   2550
      Left            =   0
      Picture         =   "frmLogin.frx":0000
      Top             =   0
      Width           =   5550
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public LoginSuccess As Boolean
Dim RS As ADODB.RecordSet
Private Sub cmdcancel_Click()
  On Error Resume Next
  With frmMain
    If .Reentry = True Then  'reentry
                            
    '------

    .StatusBar1.Panels(2).Text = "Waiting..."                   'RS!User
    .StatusBar1.Panels(4).Text = "Waiting..."                   'RS!UserLevel
    .StatusBar1.Panels(8).Text = Time
    .StatusBar1.Panels(6).Text = Date
    '-------------------------------------
        
    
        'Command Buttons; setting to enabled = true
        'when cancel button of log-off is pressed..
        '-------------
         'EnableCommandButtons
        '-------------
        '-------------
        'Menu Commands; setting to enabled = true
         'when cancel button of log-off is pressed..
         'EnableMenuCommands
         Relogin
         Unload Me
    '-------------------------------------
    Else   'first time entry
      SetOrigRes
      Set frmLogin = Nothing
   
      End
    End If
  End With
End Sub

Private Sub cmdok_Click()
    Set RS = New ADODB.RecordSet
    RS.CursorLocation = adUseClient
    RS.Open "SELECT * FROM tblAccess WHERE User = '" + txtUser + "' and Pass = '" + txtPass + "' ", CN, adOpenDynamic, adLockOptimistic
    
   
        If (txtPass.Text = "") And (txtUser.Text = "") Then '1
                MsgBox "Please Type Username and Password          ", vbOKOnly + vbInformation, " Access Denied"

                txtPass.Text = ""
                txtUser.Text = ""
                txtUser.SetFocus
                

        ElseIf (txtUser.Text = "") Then
                MsgBox "Please Type Username          ", vbOKOnly + vbInformation, " Access Denied"
                
                txtUser.Text = ""
                txtUser.SetFocus
        ElseIf (txtPass.Text = "") Then
        
                MsgBox "Please Type Password          ", vbOKOnly + vbInformation, " Access Denied"
                
                txtPass.Text = ""
                txtPass.SetFocus
                
                'Exit Sub
       'SendKeys "{Home}+{End}"
' ----------------------------------------------------------------------------------
            
        Else   'This else part signifies that txtuser & txtpas has been filled up...
        With RS
            
             
           
             If RS.RecordCount = 1 Then '2
                    frmMain.cmdCounter_Click
                    cmdRefresh_Click
                    Unload Me
                    'cmdRefresh_Click
                    MsgBox "Access Code Accepted!          ", vbInformation, "Access Granted"
                    'Call AccessLog 'audio prompt "identfication confirmed"
                    With frmMain
                    'Command Buttons; setting to enabled = true
          
                    'meaning access to the main form is granted
                    '---------
                    EnableCommandButtons
                    '---------
                    
                    'Menu Commands; setting to enabled = true
           
                    'meaning  acces is granted to user...
                    '---------
                    
                    EnableMenuCommands
                    '---------
                    .StatusBar1.Panels(2).Text = RS!User
                    .StatusBar1.Panels(4).Text = RS!UserLevel
                    .StatusBar1.Panels(8).Text = Time
                    .StatusBar1.Panels(6).Text = Date
                    
                    '-------------------------------------
                    'mnulogoff...
                    
                    '.mnuLogoff.Caption = "Log-off" + " " + RS!User + "..."
                     .mnuLogoff.Caption = "Log-off" + " " + .StatusBar1.Panels(2).Text + "..."
                    
                    End With
                    Exit Sub
                  
            Else
            'login denied
            MsgBox "Password and/or Username  Mismatch!            ", vbInformation, " Access Denied"
            'Call DeniedLog 'audio prompt "access denied"
            txtPass.Text = ""
            txtUser.Text = ""
            txtUser.SetFocus
            'SendKeys "{Home}+{End}"
                
            End If '2
        
        
        End With
        
        End If '1
  '----------------- 'Below recycle line' ( might still use them later )
 
  
  'LoginSuccess = True
  'If LoginSuccess = True Then '3
 'End If '3
End Sub

Public Sub cmdRefresh_Click()
    cmdCancel.Caption = "Cancel"
End Sub

Private Sub Form_Load()
    'LoginSuccess = False
         
       
        'Command Buttons; setting to enabled = false
        'before Acess is granted to user...
        '---------------------
        DisableCommandButtons
        '---------------------
        'Menu Commands; setting to enabled = false
        'before acces is granted to user...
        '----------------------
        DisableMenuCommands
        '----------------------
    

End Sub

Public Sub AccessLog()
'This code play's a wave file called confirmed.wav using the MMcontrol control
        
        MMControl1.Command = "Close"
        MMControl1.Filename = App.Path & "\sounds\confirmed.wav"
        MMControl1.Command = "Open"
        MMControl1.Command = "Play"
End Sub

Public Sub DeniedLog()
        'This code play's a wave file called access_denied.wav using the MMcontrol control
        
        MMControl1.Command = "Close"
        MMControl1.Filename = App.Path & "\sounds\access_denied.wav"
        MMControl1.Command = "Open"
        MMControl1.Command = "Play"
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmLogin = Nothing

End Sub

Private Sub txtuser_GotFocus()
  SendKeys hl
End Sub

Private Sub txtuser_KeyPress(KeyAscii As Integer)
  'KeyAscii = Asc(UCase(Chr(KeyAscii)))
 
  If KeyAscii = 13 Then
     txtPass.SetFocus
     'SendKeys hl
  End If
End Sub

Private Sub txtpass_GotFocus()
   'SendKeys hl
End Sub

Private Sub txtpass_KeyPress(KeyAscii As Integer)
  'KeyAscii = Asc(UCase(Chr(KeyAscii)))
  If KeyAscii = 13 Then
    Call cmdok_Click
    
  End If
End Sub

