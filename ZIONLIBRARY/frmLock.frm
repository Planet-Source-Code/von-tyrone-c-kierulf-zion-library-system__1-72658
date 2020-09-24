VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10c.ocx"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmLock 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   ClientHeight    =   6135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6435
   Icon            =   "frmLock.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   6360
      Width           =   1215
   End
   Begin VB.TextBox txtPass 
      BackColor       =   &H80000006&
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
      ForeColor       =   &H8000000E&
      Height          =   255
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   5160
      Width           =   2415
   End
   Begin VB.CommandButton cmdVisualControl 
      Height          =   615
      Index           =   0
      Left            =   2160
      Picture         =   "frmLock.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton cmdVisualControl 
      Height          =   615
      Index           =   1
      Left            =   2880
      Picture         =   "frmLock.frx":6D2C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton cmdVisualControl 
      Height          =   615
      Index           =   2
      Left            =   3600
      Picture         =   "frmLock.frx":71E6
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   3495
      Left            =   1200
      TabIndex        =   0
      Top             =   1320
      Width           =   3975
      _cx             =   7011
      _cy             =   6165
      FlashVars       =   ""
      Movie           =   "D:\Program Files\ZIONLIBRARY\IMAGES\visual.swf"
      Src             =   "D:\Program Files\ZIONLIBRARY\IMAGES\visual.swf"
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
   Begin MCI.MMControl MMControl1 
      Height          =   330
      Left            =   120
      TabIndex        =   6
      Top             =   8160
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      PrevEnabled     =   -1  'True
      NextEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      PauseEnabled    =   -1  'True
      BackEnabled     =   -1  'True
      StepEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      RecordEnabled   =   -1  'True
      EjectEnabled    =   -1  'True
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   5160
      Width           =   1095
   End
End
Attribute VB_Name = "frmLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As ADODB.RecordSet

Private Sub cmdVisualControl_Click(Index As Integer)
   If Index = 0 Then
        
        MMControl1.Command = "Close"
        MMControl1.Filename = App.Path & "\Sounds\rain.wav"
        MMControl1.Command = "Open"
        MMControl1.Command = "Play"
        
        ElseIf Index = 1 Then
        MMControl1.Command = "Pause"
        Else
        MMControl1.Command = "Close"
        
        
        End If
        txtPass.SetFocus
End Sub

Private Sub Form_Load()

  LockApplication
  '---------------

   
 
  '---------------
  ShockwaveFlash1.Movie = App.Path & "\IMAGES\visual.swf"
  
  
  
  MsgBox "The System is Locked!    ", vbInformation, "System Locked"
  
  'MMControl1.Command = "Close"
  'MMControl1.Filename = App.Path & "\Sounds\lock.wav"
  'MMControl1.Command = "Open"
  'MMControl1.Command = "Play"
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmLock = Nothing
Set RS = Nothing
End Sub

Private Sub ShockwaveFlash1_GotFocus()
  txtPass.SetFocus
End Sub



Private Sub cmdok_Click()
    Set RS = New ADODB.RecordSet
    RS.CursorLocation = adUseClient
    RS.Open "SELECT * FROM tblAccess WHERE Pass = '" + txtPass + "' ", CN, adOpenDynamic, adLockOptimistic
    
   
       
' ----------------------------------------------------------------------------------
            
   
        With RS
            
             
           
             If .RecordCount = 1 Then     '2
                    'frmMain.cmdCounter_Click
                    'cmdRefresh_Click
                    'Call Access 'audio prompt "identfication confirmed"
                    'MsgBox " "
                    'cmdRefresh_Click
                    MsgBox "The System is Unlocked!          ", vbInformation, " System Unlocked"
                  
                    With frmMain
                    'Command Buttons; setting to enabled = true
                    
                    'meaning access to the main form is granted
                    '---------
                    'EnableCommandButtons
                    '---------
                    ' .mnuFile.Enabled = True
                    'Menu Commands; setting to enabled = true
                     UnlockApplication
                    'meaning  acces is granted to user...
                    '---------
                    
                    'EnableMenuCommands
                    '---------
                    .StatusBar1.Panels(2).Text = RS!User
                    .StatusBar1.Panels(4).Text = RS!UserLevel
                    .StatusBar1.Panels(8).Text = Time
                    .StatusBar1.Panels(6).Text = Date
                    
                    '-------------------------------------
                    'mnulogoff...
                    
                    '.mnuLogoff.Caption = "Log-off" + " " + RS!User + "..."
                     .mnuLogoff.Caption = "Log-off" + " " + .StatusBar1.Panels(2).Text + "..."
                    
                    Set frmLock = Nothing
                    Set RS = Nothing
                    Unload Me
                   
                   
                    
                    End With
                    'Exit Sub
                  
            Else
            'login denied
              If txtPass.Text = "" Then
              MsgBox "Please enter Access Code!          ", vbInformation, "Access denied"
              
              Else
              MsgBox "Password Mismatch!          ", vbInformation, " Access Denied"
              'Call Denied 'audio prompt "access denied"
              txtPass.Text = ""
            
              txtPass.SetFocus
              'SendKeys "{Home}+{End}"
              End If
                
            End If '2
        
        End With
  
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

Private Sub Access()
'This code play's a wave file called confirmed.wav using the MMcontrol control
        
        MMControl1.Command = "Close"
        MMControl1.Filename = App.Path & "\sounds\unlock.wav"
        MMControl1.Command = "Open"
        MMControl1.Command = "Play"
End Sub

Private Sub Denied()
        'This code play's a wave file called access_denied.wav using the MMcontrol control
        
        MMControl1.Command = "Close"
        MMControl1.Filename = App.Path & "\sounds\notallowed.wav"
        MMControl1.Command = "Open"
        MMControl1.Command = "Play"
        
End Sub


