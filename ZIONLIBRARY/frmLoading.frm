VERSION 5.00
Begin VB.Form frmLoading 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   ClientHeight    =   4005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6900
   LinkTopic       =   "Form1"
   Picture         =   "frmLoading.frx":0000
   ScaleHeight     =   4005
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMaximumDays 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Text            =   "3"
      Top             =   4800
      Width           =   3975
   End
   Begin VB.TextBox txtFinePerDay 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Text            =   "5"
      Top             =   6000
      Width           =   3975
   End
End
Attribute VB_Name = "frmLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public Bin1 As Variant
'Public Bin2 As Variant
'Public Bin3 As Variant
'Public Bin4 As Variant
'Public Bin5 As Variant
'Public RegName As String
'Public OrgName As String

Private Sub Form_Click()
Unload Me
frmLogin.Show vbModal
End Sub

Private Sub Form_Deactivate()
Unload Me
frmLogin.Show vbModal
End Sub

Private Sub Form_Initialize()
    Main 'set trap for a single instance of the app
    '--------------------------
    'change screen resolution of the system...
    'used to make sure...our program will be displayed
    'using the appropriate screen resolution...
    ChangeRes 800, 600, 32, 85
    GetCurrentRes
   
    
End Sub

Private Sub Form_Load()


      With frmMain
         .StatusBar1.Panels(9).Text = ""
      End With
    DisableMenuCommands   'disables frmMain menu
    '------------
    DisableCommandButtons 'disables frmMain command buttons
    '------------
    'frmReturn.FineAmnt = CCur(GetSetting(App.Title, "Settings", "Fine Amount"))
    'frmReturn.MaxDays = CInt(GetSetting(App.Title, "Settings", "Max Days"))
    
    
    '------------
   
    'SaveSetting App.Title, "Settings", "Fine Amount", CStr(CCur(txtFinePerDay.Text))
    'SaveSetting App.Title, "Settings", "Max Days", CStr(CCur(txtMaximumDays.Text))
End Sub


Private Sub Form_Unload(Cancel As Integer)
Set frmLoading = Nothing
Set RS = Nothing
End Sub
