VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMembersAE 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Record"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   Icon            =   "frmMembersAE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   27
      Top             =   6720
      Width           =   375
   End
   Begin VB.CommandButton cmdPicInsert 
      Height          =   375
      Left            =   3960
      Picture         =   "frmMembersAE.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Browse for a picture to store in the database..."
      Top             =   5520
      Width           =   375
   End
   Begin VB.CommandButton cmdPicSave 
      Height          =   375
      Left            =   2040
      Picture         =   "frmMembersAE.frx":6DBC
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Save Picture in a file..."
      Top             =   5520
      Width           =   375
   End
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   120
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".jpg"
      DialogTitle     =   "Save File from database as..."
      Filter          =   "Picture Files (*.jpg,*.bmp,*.wmf,*.emf)|*.jpg;*.bmp;*.wmf;*.emf|All files (*.*)|*.*"
   End
   Begin VB.TextBox txtCode 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox txtRoll 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1800
      MaxLength       =   2
      TabIndex        =   5
      Top             =   3360
      Width           =   975
   End
   Begin VB.ComboBox cboClass 
      Height          =   315
      ItemData        =   "frmMembersAE.frx":720E
      Left            =   1800
      List            =   "frmMembersAE.frx":7221
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2640
      Width           =   2895
   End
   Begin VB.ComboBox cboSection 
      Height          =   315
      ItemData        =   "frmMembersAE.frx":7238
      Left            =   1800
      List            =   "frmMembersAE.frx":7248
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3000
      Width           =   2895
   End
   Begin VB.TextBox txtLast 
      Height          =   285
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   2
      Top             =   2280
      Width           =   2895
   End
   Begin VB.TextBox txtM 
      Height          =   285
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   1
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox txtFirst 
      Height          =   285
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1560
      Width           =   2895
   End
   Begin VB.CommandButton cmdAddSave 
      Caption         =   "Update"
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   6075
      Width           =   1095
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   6075
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   8
      Top             =   6075
      Width           =   1095
   End
   Begin ZionLibrary.Photo Photo1 
      Height          =   2175
      Left            =   2040
      TabIndex        =   25
      Top             =   3720
      Width           =   2295
      _extentx        =   5530
      _extenty        =   2990
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMembersAE.frx":725C
      Height          =   855
      Left            =   720
      TabIndex        =   24
      Top             =   0
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   120
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   5
      Left            =   2880
      TabIndex        =   23
      Top             =   3360
      Width           =   105
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   4
      Left            =   4800
      TabIndex        =   22
      Top             =   3000
      Width           =   105
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   3
      Left            =   4800
      TabIndex        =   21
      Top             =   2640
      Width           =   105
   End
   Begin VB.Label Label10 
      Caption         =   "Roll:"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Section:"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Year:"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   2
      Left            =   4800
      TabIndex        =   17
      Top             =   1560
      Width           =   105
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   1
      Left            =   4800
      TabIndex        =   16
      Top             =   2280
      Width           =   105
   End
   Begin VB.Label Label6 
      Caption         =   "Picture:"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Last Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Middle Initial:"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "First Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Student ID:"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   120
      X2              =   5040
      Y1              =   975
      Y2              =   975
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      X1              =   120
      X2              =   5040
      Y1              =   5955
      Y2              =   5955
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   120
      X2              =   5040
      Y1              =   975
      Y2              =   975
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   120
      X2              =   5040
      Y1              =   5040
      Y2              =   5040
   End
End
Attribute VB_Name = "frmMembersAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
'-------------------------------------------------------------------
'Members Add/Edit Form
'With this form it is possible to add/modify the members table.
'
'AddState is the bool value that decides whether the form will update
'or modify a record. When AddState = True then Add
'                    When AddState = False then Modify
'-------------------------------------------------------------------
Private RS As ADODB.RecordSet
Public OldID As String, AddState As Boolean
Private str0 As String
Private str1 As String
Private str2 As String
Private str3 As String
Private str4 As String
Private str5 As String
Private str6 As String
Private str7 As Variant
Private str7x As Variant
Private N0 As String
Private N1 As String
Private N2 As String
Private N3 As String
Private N4 As String
Private N5 As String
Private N6 As String

Private Sub Form_Load()
   
    On Error GoTo Err
    
   
  
    AddSuccess = False ' this is to set back to default so that if user does not add up a
                       'new record, when the form unloads, places back the cursor to
                       'the old position
   
    Set RS = New ADODB.RecordSet
    If AddState Then
      
        
        RS.Open "SELECT * FROM tblMembers", CN, adOpenStatic, adLockOptimistic
        Me.Caption = "Add Record"
        Image1.Picture = frmMembers.cmdAMod(1).Picture
        cmdAddSave.Caption = "Save"
      
    Else 'NOT AddState...
       
        Me.Caption = "Modify Record"
          Image1.Picture = frmMembers.cmdAMod(0).Picture
        cmdAddSave.Caption = "Update"
        RS.Open "SELECT * FROM tblMembers WHERE [Student ID] = '" & OldID & "'", CN, adOpenStatic, adLockOptimistic
        If Len(RS!Picture) > 0 Then
        Photo1.LoadPhoto RS!Picture
        Von
        Tyrone
        End If
    End If

Exit Sub

Err:
    If Err.Number = 94 Or Err.Number = 3265 Then
        Resume Next 'If a null value is encountered
    Else
        Handler Err 'Unexpected error
    End If

End Sub

Private Sub cmdAddSave_Click()

  On Error GoTo hell

    If AddState Then '1 if addstate...
                           'If txtCode.Text = ""
                           If txtFirst.Text = "" Or txtM.Text = "" _
                           Or txtLast.Text = "" Or cboClass.Text = "" _
                           Or cboSection.Text = "" Or txtRoll.Text = "" Then   '2
                            'if a field is left blank..except the picture
        
                    MsgBox "Missing Data, please complete the Fields!"
                    txtFirst.SetFocus
        
       
       
        
        Else 'all fields filled up...then proceed...
            '----------------------------------------------------------------------
            'place record checking trap here...
              If RecordExists("tblMembers", "Student ID", txtCode.Text, txtCode) = True Then
            '3 if record exist..no duplication, do not add
             txtFirst.SetFocus
            
        
            
            '----------------------------------------------------------------------
            
              Else 'if record does not exist then..then proceed adding...
            
               
              
                'If IsNumeric(txtRoll.Text) <> True Then MsgBox "Roll Numbers must be numeric and between 1 and 99", vbExclamation, "Type Mismatch": HighLight txtRoll: Exit Sub
                '-------------------
                
                '--------------------
                CN.BeginTrans
                With RS
           
                
                .AddNew
         
                .Fields(0) = txtCode.Text
                .Fields(1) = txtFirst.Text
                .Fields(2) = txtM.Text
                .Fields(3) = txtLast.Text
                .Fields(4) = cboClass.Text
                .Fields(5) = cboSection.Text
                .Fields(6) = Int(txtRoll.Text)
                Photo1.SavePhoto .Fields("Picture")
               

                .Update
                End With
                CN.CommitTrans
                FindRecord RS, RS.Fields(0).Name, True, txtCode.Text, 0
                AddSuccess = True 'sets a marker here so that when user adds up a new record
                                  'when form unloads, cursor points to the last record,
                                  'the newly added record
                frmMembers.cmdRefresh_Click 'the pointing of the cursor is done here...
               
                MsgBox "New record has been successfully added", vbInformation
                
'----------------------------------------------------------------------
                If MsgBox("Do you want to add a new record?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                  'AddedNew = True 'sets back again so that cursor would just point to the last rec
                'Unload Me
                'frmMembersAE.Show
                cmdReset_Click
                Else
                Unload Me
                End If
              
            End If '3
        End If '2
        

    Else 'not addstate...( edit state )
        '--------------
        
        
        '--------------
        'If txtCode.Text = ""
        If txtFirst.Text = "" Or txtM.Text = "" _
        Or txtLast.Text = "" Or cboClass.Text = "" _
        Or cboSection.Text = "" Or txtRoll.Text = "" Then  'z
        'if a field is left blank..do not update
        MsgBox "Missing Data, please complete the Fields!"
        txtFirst.SetFocus
                               
        Else 'no blank fields...except pic ...then update
        CN.BeginTrans
        With RS
    
              
                .Fields(0) = txtCode.Text
                .Fields(1) = txtFirst.Text
                .Fields(2) = txtM.Text
                .Fields(3) = txtLast.Text
                .Fields(4) = cboClass.Text
                .Fields(5) = cboSection.Text
                .Fields(6) = Int(txtRoll.Text)
                 Photo1.SavePhoto .Fields("Picture")
               
                .Update
        
        End With
        CN.CommitTrans
        frmMembers.cmdRefresh_Click
      
        FindRecord RS, RS.Fields(0).Name, True, txtCode.Text, 0
                  
                    'If Len(RS!Picture) <= 0 Then
                     '   If str0 = txtCode.Text And _
                        str1 = txtFirst.Text And _
                        str2 = txtM.Text And _
                        str3 = txtLast.Text And _
                        str4 = cboClass.Text And _
                        str5 = cboSection.Text And _
                        str6 = txtRoll.Text Then
                        
                        
                      '  MsgBox "No changes made to this record.", vbInformation
                       ' Unload Me
                    
                        'Else
                        'MsgBox "Changes in record has been successfully saved", vbInformation
                        'Unload Me
                      
                        'End If
                    'End If
                    
                    
                    '-----------------------------------------------------
                   If Len(RS!Picture) > 0 Then 'X not null value in field picture
                        
                        
                        str7x = CStr(RS.Fields("Picture")) 'assign the newly saved pic to a string  variable
                        
                        If str0 = txtCode.Text And _
                        str1 = txtFirst.Text And _
                        str2 = txtM.Text And _
                        str3 = txtLast.Text And _
                        str4 = cboClass.Text And _
                        str5 = cboSection.Text And _
                        str6 = txtRoll.Text And _
                        str7 = str7x Then
                    

                        MsgBox "No changes made to this record.", vbInformation
                        Unload Me
                      
                        Else
                        
                        MsgBox "Changes in record has been successfully saved", vbInformation
                        Unload Me
                      
                        End If
                        'Exit Sub
                  
                  
                  Else  'null value in field picture
                        'str7x = CStr(RS.Fields("Picture"))
                         Softwise
                         'If N0 = txtCode.Text And _
                         N1 = txtFirst.Text And _
                         N2 = txtM.Text And _
                         N3 = txtLast.Text And _
                         N4 = cboClass.Text And _
                         N5 = cboSection.Text And _
                         N6 = txtRoll.Text Then
                         'MsgBox "hello"
                        
                         'MsgBox "No changes made to this record.", vbInformation
                         'Unload Me
                         
                         'Else
                        
                         'MsgBox "Changes in record has been successfully saved", vbInformation
                     
                         'Unload Me
                      
                         'End If
                         
                         
                  End If 'X
                
                
            End If 'z
   
    End If '1
    
hell:
      'If Err.Number = 94 Or Err.Number = 3265 Then
       Resume Next '-If encounter a null value
      'Else
    
        Handler Err
        CN.RollbackTrans
        
      'End If


End Sub

Private Sub cboClass_Click()

    MakeCode

End Sub

Private Sub cboSection_Click()

    MakeCode

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmMembersAE = Nothing
    Set RS = Nothing
End Sub

Private Sub txtRoll_KeyPress(KeyAscii As Integer)
     KeyAscii = Number_Keypress(KeyAscii)
End Sub

Private Sub txtRoll_LostFocus()

    MakeCode

End Sub

Private Sub cmdcancel_Click()
    Set frmMembersAE = Nothing
    Set RS = Nothing
    Unload Me

End Sub

Private Sub cmdPicInsert_Click()

'Open photo from disk

    Photo1.OpenPhotoFile

End Sub

Private Sub cmdPicSave_Click()

'Save photo to disk

    On Error GoTo hell
    cmdlg.ShowSave
    If cmdlg.Filename <> "" Then
        SavePicture Photo1.Picture, cmdlg.Filename
    End If
hell:

End Sub

Private Sub cmdPicShow_Click()

'Open photo from a temp file

    On Error Resume Next
        Kill "tmp.jpg"
        SavePicture Photo1.Picture, "tmp.jpg"
        ShellEx "tmp.jpg"

End Sub

Private Sub cmdReset_Click()

    txtCode.Text = ""
    txtFirst.Text = ""
    txtLast.Text = ""
    txtM.Text = ""
    txtRoll.Text = ""
    cboClass.ListIndex = 0
    cboSection.ListIndex = 0
    txtFirst.SetFocus
    
    
End Sub

Private Sub MakeCode()

'This sub automatically generates Student Code

Dim A As String, b As String

    If cboSection.Text <> "" And cboClass.Text <> "" And txtRoll.Text <> "" Then
        Select Case cboSection.ListIndex
        Case 0: A = "RD"
        Case 1: A = "GR"
        Case 2: A = "BL"
        Case 3: A = "WH"
        End Select
        
        
        b = cboClass.ListIndex + 1
        If b = 10 Then b = "X"
        If txtRoll.Text < 10 Then txtRoll.Text = "0" & txtRoll.Text
        txtCode.Text = A & b & txtRoll.Text
    End If

End Sub

Private Sub Von()
       On Error Resume Next
       If AddState = False Then
            With RS
    
            str0 = .Fields(0)
            str1 = .Fields(1)
            str2 = .Fields(2)
            str3 = .Fields(3)
            str4 = .Fields(4)
            str5 = .Fields(5)
            str6 = .Fields(6)
            str7 = CStr(.Fields("Picture"))
            End With
        End If

    
End Sub
Private Sub Tyrone()
  On Error Resume Next
  If AddState = False Then
    With RS
    N0 = .Fields(0)
    N1 = .Fields(1)
    N2 = .Fields(2)
    N3 = .Fields(3)
    N4 = .Fields(4)
    N5 = .Fields(5)
    N6 = .Fields(6)
    
    
    End With
  End If
End Sub
Private Sub Softwise()
    If N0 = txtCode.Text And _
       N1 = txtFirst.Text And _
       N2 = txtM.Text And _
       N3 = txtLast.Text And _
       N4 = cboClass.Text And _
       N5 = cboSection.Text And _
       N6 = txtRoll.Text Then
                         
                        
      MsgBox "No changes made to this record.", vbInformation
      Unload Me
                         
      Else
                        
      MsgBox "Changes in record has been successfully saved", vbInformation
                     
      Unload Me
                      
      End If
                         
End Sub





