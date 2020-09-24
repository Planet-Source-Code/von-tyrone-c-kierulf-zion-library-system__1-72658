VERSION 5.00
Begin VB.Form MathCalc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Calculator   &   Clock       "
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3165
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MathCalc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   3165
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   2520
      Width           =   2895
   End
   Begin VB.Timer tmrTimer 
      Interval        =   500
      Left            =   720
      Top             =   1080
   End
   Begin VB.CommandButton DotBttn 
      Appearance      =   0  'Flat
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton ClearBttn 
      Appearance      =   0  'Flat
      Caption         =   "C"
      Height          =   375
      Index           =   10
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Equals 
      Appearance      =   0  'Flat
      Caption         =   "="
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Div 
      Appearance      =   0  'Flat
      Caption         =   "/"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Times 
      Appearance      =   0  'Flat
      Caption         =   "*"
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Minus 
      Appearance      =   0  'Flat
      Caption         =   "-"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton Plus 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton Digits 
      Appearance      =   0  'Flat
      Caption         =   "0"
      Height          =   375
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Digits 
      Appearance      =   0  'Flat
      Caption         =   "9"
      Height          =   375
      Index           =   9
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton Digits 
      Appearance      =   0  'Flat
      Caption         =   "8"
      Height          =   375
      Index           =   8
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton Digits 
      Appearance      =   0  'Flat
      Caption         =   "7"
      Height          =   375
      Index           =   7
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton Digits 
      Appearance      =   0  'Flat
      Caption         =   "6"
      Height          =   375
      Index           =   6
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton Digits 
      Appearance      =   0  'Flat
      Caption         =   "5"
      Height          =   375
      Index           =   5
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton Digits 
      Appearance      =   0  'Flat
      Caption         =   "4"
      Height          =   375
      Index           =   4
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton Digits 
      Appearance      =   0  'Flat
      Caption         =   "3"
      Height          =   375
      Index           =   3
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Digits 
      Appearance      =   0  'Flat
      Caption         =   "2"
      Height          =   375
      Index           =   2
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Digits 
      Appearance      =   0  'Flat
      Caption         =   "1"
      Height          =   375
      Index           =   1
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblVon 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VON"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   1920
      TabIndex        =   18
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Display 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "MathCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private OldCaption As String
Dim Operand1 As Double, Operand2 As Double
Dim Operator As String
Dim ClearDisplay As Boolean
Public CalcDisp As Currency


Private Sub ClearBttn_Click(Index As Integer)
Display.Caption = ""
End Sub

Private Sub cmdExport_Click()
    
    If Display = "" Then
    MsgBox "You cannot export a null value,please type in an amount", vbInformation, "Export Denied"
    
    Else
    CalcDisp = Display.Caption
    frmReturn.txtBookPayable = CalcDisp
    frmReturn.txtBookPayable = CStr(FormatCurrency$(frmReturn.txtBookPayable))
    Set MathCalc = Nothing
    Unload Me
    End If
End Sub

Private Sub Digits_Click(Index As Integer)
If ClearDisplay Then
Display.Caption = ""
ClearDisplay = False
End If
Display.Caption = Display.Caption + Digits(Index).Caption
End Sub

Private Sub Div_Click()
Operand1 = Val(Display.Caption)
Operator = "/"
Display.Caption = ""
End Sub

Private Sub DotBttn_Click(Index As Integer)
If InStr(Display.Caption, ".") Then
Exit Sub
Else
Display.Caption = Display.Caption + "."
End If
End Sub

Private Sub Equals_Click()
Dim Result As Double
Operand2 = Val(Display.Caption)
If Operator = "+" Then Result = Operand1 + Operand2
If Operator = "-" Then Result = Operand1 - Operand2
If Operator = "*" Then Result = Operand1 * Operand2
If Operator = "/" And Operand2 <> "0" Then _
Result = Operand1 / Operand2
Display.Caption = Result

End Sub



Private Sub Form_Unload(Cancel As Integer)
Set MathCalc = Nothing
End Sub

Private Sub Minus_Click()
Operand1 = Val(Display.Caption)
Operator = "-"
Display.Caption = ""
End Sub

Private Sub Plus_Click()
Operand1 = Val(Display.Caption)
Operator = "+"
Display.Caption = ""
End Sub

Private Sub Times_Click()
Operand1 = Val(Display.Caption)
Operator = "*"
Display.Caption = ""
End Sub

Private Sub tmrTimer_Timer()
Dim msg As String
msg = OldCaption & ": " & Now()

If msg <> Caption Then
Caption = msg
End If
End Sub


