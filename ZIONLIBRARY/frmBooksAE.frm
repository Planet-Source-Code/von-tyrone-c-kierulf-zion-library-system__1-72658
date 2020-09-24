VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmBooksAE 
   Caption         =   "Add Book Record"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5310
   Icon            =   "frmBooksAE.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPublisher 
      DataField       =   "Publiser Name"
      Height          =   285
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   3
      Top             =   2280
      Width           =   2895
   End
   Begin VB.TextBox txtAuthor 
      DataField       =   "Author"
      Height          =   285
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1920
      Width           =   2895
   End
   Begin VB.TextBox txtTitle 
      DataField       =   "Title"
      Height          =   285
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1560
      Width           =   2895
   End
   Begin VB.CommandButton cmdAddSave 
      Caption         =   "Update"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   4875
      Width           =   1095
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   4875
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   9
      Top             =   4875
      Width           =   1095
   End
   Begin VB.ComboBox cboCategory 
      DataField       =   "Category"
      Height          =   315
      ItemData        =   "frmBooksAE.frx":6852
      Left            =   1800
      List            =   "frmBooksAE.frx":6874
      Sorted          =   -1  'True
      TabIndex        =   4
      Text            =   "N/A"
      Top             =   2640
      Width           =   2895
   End
   Begin MSMask.MaskEdBox txtPrice 
      DataField       =   "Price"
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Top             =   3000
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   9
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox msdISBN 
      DataField       =   "ISBN"
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Top             =   3360
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      _Version        =   393216
      AutoTab         =   -1  'True
      MaxLength       =   13
      Mask            =   "#-###-#####-C"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox msdID 
      DataField       =   "Book ID"
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   1200
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      _Version        =   393216
      AutoTab         =   -1  'True
      MaxLength       =   10
      Mask            =   "B#########"
      PromptChar      =   "_"
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   120
      X2              =   5160
      Y1              =   4755
      Y2              =   4755
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   120
      X2              =   5040
      Y1              =   975
      Y2              =   975
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
      TabIndex        =   21
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
      Index           =   0
      Left            =   4800
      TabIndex        =   20
      Top             =   1200
      Width           =   105
   End
   Begin VB.Label Label6 
      Caption         =   "Category:"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Publisher:"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Author:"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Title:"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Book ID:"
      Height          =   255
      Left            =   240
      TabIndex        =   15
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
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmBooksAE.frx":693D
      Stretch         =   -1  'True
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmBooksAE.frx":7701
      Height          =   855
      Left            =   840
      TabIndex        =   14
      Top             =   120
      Width           =   4335
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      X1              =   120
      X2              =   5160
      Y1              =   4755
      Y2              =   4755
   End
   Begin VB.Label Label8 
      Caption         =   "Price:"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "ISBN:"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Borrowed:"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   $"frmBooksAE.frx":77D8
      Height          =   855
      Left            =   1800
      TabIndex        =   10
      Top             =   3720
      Width           =   2895
   End
End
Attribute VB_Name = "frmBooksAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-------------------------------------------------------------------
'Books Add/Edit Form
'This form it is possible to add/modify the books table.
'
'AddState is the bool value that decides whether the form will update
'or modify a record. When AddState = True then Add
'                    When AddState = False then Modify
'-------------------------------------------------------------------
Public AddState As Boolean, OldID As String
Private RS As ADODB.RecordSet

Private str0 As String
Private str1 As String
Private str2 As String
Private str3 As String
Private str4 As String
Private str5 As Currency
Private str6 As String

Private Sub Form_Load()

'Prepare form to add/edit. This is done because we are reusing the form.
    
    On Error GoTo Err
    AddSuccess = False ' this is to set back to default so that if user does not add up a
                       'new record, when the form unloads, places back the cursor to
                       'the old position
   
    Set RS = New ADODB.RecordSet
    If AddState Then
        SendKeys "{RIGHT}" 'to Tab and to move the cursor to the char after B...
        Image1.Picture = frmBooks.cmdAMod(1).Picture
        RS.Open "SELECT * FROM tblBooks", CN, adOpenStatic, adLockOptimistic
        Me.Caption = "Add Record"
        cmdAddSave.Caption = "Save"
    Else 'NOT AddState...
        
        Image1.Picture = frmBooks.cmdAMod(0).Picture
        Me.Caption = "Modify Record"
        cmdAddSave.Caption = "Update"
        RS.Open "SELECT * FROM tblBooks WHERE [Book ID] = '" & OldID & "'", CN, adOpenStatic, adLockOptimistic
        Von
        'WHERE [Book ID] = '" & OldID & "'"
        'stores values to string variables and a currency var
            
    End If
    cmdReset_Click
Exit Sub

Err:
    If Err.Number = 94 Or Err.Number = 3265 Then
        Resume Next '-If encounter a null value
    Else
        Handler Err '-Unexpected error
    End If

End Sub

Private Sub cmdAddSave_Click()
      On Error GoTo hell

    If AddState Then '1 if addstate...
        
        If msdID.Text = "" Or txtTitle.Text = "" Or txtAuthor.Text = "" _
                           Or txtPublisher.Text = "" Or txtPrice.Text = "" Then '2
                            'if a field is left blank..
        
                    MsgBox "Missing Data, please complete the Fields!"
                    msdID.SetFocus
        
        
       
        
        Else 'all fields filled up...then proceed...
            '----------------------------------------------------------------------
            'place record checking trap here...
             If RecordExists("tblBooks", "Book ID", msdID.Text, msdID) = True Then
            '3 if record exist..no duplication, do not add
             msdID.SetFocus
        
            
            '----------------------------------------------------------------------
            
            Else 'if record does not exist then..then proceed adding...
            
               
                
                '4 if characters inputted is less than 9..
                'msdID.Text = Len(msdID.Text)
                'If Len(Trim(msdID.Text)) < 10 Then
                'MsgBox "cannot...""": Exit Sub
                If IsNumeric(Right$(msdID.Text, 9)) = False Then MsgBox "Book ID must start with B followed by 9 digits" & vbCrLf & "Please refer to the last record(Book ID) of this table.", vbExclamation:   msdID.SetFocus: Exit Sub
                '-------------------
                
                '--------------------
                CN.BeginTrans
                With RS
           
                
                 RS.AddNew
                .Fields(0) = msdID.Text
                .Fields(1) = txtTitle.Text
                .Fields(2) = txtAuthor.Text
                .Fields(3) = txtPublisher.Text
                .Fields(4) = cboCategory.Text
                .Fields(5) = CCur(txtPrice.Text)
                .Fields(6) = msdISBN.Text

                RS.Update
                End With
                CN.CommitTrans
                FindRecord RS, RS.Fields(0).Name, True, msdID.Text, 0
                '--------------------
                'frmBooks.cmdRefresh_Click
                'AddSuccess = True
                '--------------------
                AddSuccess = True 'sets a marker here so that when user adds up a new record
                                  'when form unloads, cursor points to the last record,
                                  'the newly added record
                frmBooks.cmdRefresh_Click 'the pointing of the cursor is done here...
               
                '--------------------
                MsgBox "New record has been successfully added", vbInformation
'----------------------------------------------------------------------
                If MsgBox("Do you want to add a new record?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
            
                cmdReset_Click
                
                'Unload Me
                'frmBooksAE.Show
              
                Else
                Unload Me
                End If
              
            End If '3
        End If '2
        

    Else 'not addstate...( edit state )
        
        '--------------
        If msdID.Text = "" _
        Or txtTitle.Text = "" Or txtAuthor.Text = "" _
        Or txtPublisher.Text = "" Or txtPrice.Text = "" Then 'z
        'if a field is left blank..do not update
        MsgBox "Missing Data, please complete the Fields!"
        msdID.SetFocus
                               
        Else 'no blank fields...then update
        CN.BeginTrans
        With RS
    
            .Fields(0) = msdID.Text
            .Fields(1) = txtTitle.Text
            .Fields(2) = txtAuthor.Text
            .Fields(3) = txtPublisher.Text
            .Fields(4) = cboCategory.Text
            .Fields(5) = CCur(txtPrice.Text)
            .Fields(6) = msdISBN.Text
    
   
        
        RS.Update
        
        End With
        CN.CommitTrans
        FindRecord RS, RS.Fields(0).Name, True, msdID.Text, 0
        frmBooks.cmdRefresh_Click
               If msdID.Text = str0 And txtTitle.Text = str1 And txtAuthor.Text = str2 And _
           txtPublisher.Text = str3 And cboCategory.Text = str4 And txtPrice.Text = str5 And _
               msdISBN.Text = str6 Then
                'if no changes made
                MsgBox "No changes made to this record.", vbInformation
                Unload Me
                Else 'if changes have been made
                MsgBox "Changes in record has been successfully saved", vbInformation
                Unload Me
                End If
            End If 'z
   
    End If '1
    
hell:
    On Error Resume Next
        Handler Err
        CN.RollbackTrans


End Sub

Private Sub cmdcancel_Click()
    Set frmBooksAE = Nothing
    Set RS = Nothing
    Unload Me
    
End Sub

Private Sub cmdReset_Click()

  On Error Resume Next

'Reset all values to nothing/null/empty/nullstring/0
    'msdID.Mask = "B#########"
    'msdISBN.Mask = "#-###-#####-C"
   '---------------------------------edit below
     If AddState = False Then
     SendKeys "{TAB}" 'when the form loads..reset key is pressed
                              'and the tabs the focus to txtTitle txtbox
                              'no need to set the focus in the first textbox..
                              'cannot edit the BookID anyway...
     End If
     txtTitle.Text = ""
     txtAuthor.Text = ""
     txtPublisher.Text = ""
     txtPrice.Text = ""
     cboCategory.Text = "N/A"
    
    '--------------------
   
     '-------------------
     msdID.Text = ""
    '-------------------
    '----------------------
     msdISBN.Text = ""
    '----------------------
    '-----------------------
    'cboCategory.ListIndex = 0
    '-----------------------
     msdID.SetFocus
    '-------------------------------------
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmBooksAE = Nothing
Set RS = Nothing
End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)
    KeyAscii = Number_Keypress(KeyAscii)
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
            End With
        End If

    
End Sub

