VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4170
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Save and close"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox txtFinePerDay 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "5"
      Top             =   1920
      Width           =   3975
   End
   Begin VB.TextBox txtMaximumDays 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "3"
      Top             =   720
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "What is the ammount of fines per day enforced if a book is not returned on time?"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "What is the maximum number of days a book can be kept before the fines are generataed?"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------------------
'This form will record Maximum number of days a book can be borrowed
'before the fine is placed and the amount of fine imposed per day the
'book is late in the registry. Used Registry function SaveSetting and GetSetting here...
'-----------------------------------------------------------------

Private Sub Command1_Click()

    On Error GoTo hell
    
    'Using SaveSetting Registry function here..we will save values in the registry...
    'values of entered and saved in this form will be saved in the windows registry..
    'Ref. pp. 844 of Visual Basic Complete..

    If txtMaximumDays.Text = "" Or IsNumeric(txtMaximumDays.Text) = False Or txtMaximumDays.Text < 0 Or txtFinePerDay.Text = "" Or IsNumeric(txtFinePerDay.Text) = False Or txtFinePerDay.Text < 0 Then
        GoTo hell
        Exit Sub '---> Bottom
    Else 'NOT txtMaximumDays.TEXT...
        SaveSetting App.Title, "Settings", "Fine Amount", CStr(CCur(txtFinePerDay.Text))
        SaveSetting App.Title, "Settings", "Max Days", CStr(CCur(txtMaximumDays.Text))
        Unload Me
        Set frmSettings = Nothing
    End If

Exit Sub

hell:
    MsgBox "You have entered an invalid charecter or no charecters at all in the textboxes" & vbNewLine & "therefore you cannot save the settings" & vbNewLine & "You can enter only numeric data in the boxes", vbExclamation

End Sub

Private Sub Form_Load()
    'Using GetSetting Registry function here..we will get values saved in the Windows registry...
    'values saved in the windows registry will be retrieved here...
    'Ref. pp. 844 of Visual Basic Complete..

    'txtFinePerDay.Text = GetSetting(App.Title, "Settings", "Fine Amount", "5")
    'txtMaximumDays.Text = GetSetting(App.Title, "Settings", "Max Days", "3")
    txtFinePerDay.Text = GetSetting(App.Title, "Settings", "Fine Amount")
    txtMaximumDays.Text = GetSetting(App.Title, "Settings", "Max Days")


End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmSettings = Nothing
End Sub
