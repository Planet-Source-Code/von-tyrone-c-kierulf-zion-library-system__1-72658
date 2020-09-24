Attribute VB_Name = "ModGeneral"
Option Explicit
Public CN As ADODB.Connection

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public AddSuccess As Boolean
Public OldPosition As Long
'CN is declared as as public to be accesible to all forms
'note: The keyword Public is equivalent to the dim statement ref. pp 766 of visual basic 6 complete
'Public Declare Sub Navigate Lib "Von.dll" (index As Integer, RS As Recordset) tried using a dll to "house this sub..
Public Sub Navigate(Index As Integer, RecordSet As ADODB.RecordSet)
' I have set this reusable code as public to provide navigation accessible to every form
    On Local Error Resume Next 'sets a local error trapping resume next for run time errors
        With RecordSet
            'Using ( WITH...END WITH ) structure to have a convenient syntax and to
            'make CODE easier to read and make it run faster...
            'refer to page 488 of visual basic 6,how to program by deitel
            'or to page 645 of the visual basic 6 complete
            Select Case Index
            Case 0
                If Not .RecordCount <= 1 Then
                    .MoveFirst
                     
                 End If
            Case 3
                If Not .RecordCount <= 1 Then
                    .MoveLast
                   
                End If
            Case 2
                If Not .AbsolutePosition >= .RecordCount Or .RecordCount <= 1 Then
                    .MoveNext
                  
                End If
            Case 1
                If Not .AbsolutePosition <= 1 Then
                    .MovePrevious
                   
                End If
             End Select
        End With
End Sub
Public Sub Handler(Error As ErrObject)

'Shows msgbox for unhandled errors only when error has truly occured,
'i.e. err<>0
    
    If Error.Number <> 0 Then
        MsgBox "Error Number: " & Error.Number & vbNewLine & Error.Description, vbExclamation, "Unexpected Error"
    
    End If
      
End Sub


Public Sub HighLight(ByRef sObj As Object)

'Procedure highlights text in a textbox
'SelStart and SelLength property...of a text box
'reference pp. 603,Vb 6, how to program by deitel & deitel

    With sObj
        .SelStart = 1
        .SelLength = Len(sObj.Text)
    End With

End Sub

Public Function RecordExists(ByVal sTable As String, ByVal sField As String, ByVal sStr As String, ByRef sEntryField As Object) As Boolean

Dim rs As New ADODB.RecordSet

    rs.Open "Select * From " & sTable & " Where [" & sField & "] = '" & sStr & "'", CN, adOpenStatic, adLockReadOnly
    If rs.RecordCount < 1 Then
        RecordExists = False
    Else
        MsgBox "The adding of new entry cannot be done because " & sStr & " already" & vbCrLf & "exists in the recordset. Please check and change it." & vbCrLf & vbCrLf & "Note: Duplication of entries is not allowed in this application.", vbExclamation
        HighLight sEntryField 'highlights if record is found in the database...
        RecordExists = True
    End If
    Set rs = Nothing
    
    'variable explanation...
  
    'stable holds the variable tables of either books or members
    'sField is the var that holds the field variable
      'sStr is the var that holds booksID
    'sEntryField holds the variable of the textbox or masked...
    
    'If RecordExists("tblBooks", "Book ID", msdID.Text, msdID)
End Function

Public Sub FindRecord(ByRef sRS As ADODB.RecordSet, ByVal sField As String, ByVal isString As Boolean, ByVal sStr As String, ByVal sNum As Long)

'This procedure finds a record in the selected recordset
'and sets its absolute position with the found record.

    On Local Error Resume Next
        With sRS
            
            .Filter = adFilterNone
            .Requery
            .MoveFirst
            
            If isString Then
                .Find sField & " = '" & sStr & "'"
            Else 'NOT ISSTRING...
                .Find sField & " = " & sNum
            End If
        End With
        '========================================================
    'FindRecord RS, RS.Fields(0).Name, True, msdID.Text, 0
    'sRS , sField, isString, sStr, sNum

      
End Sub

Public Sub CenterObj(ByRef ChildObj As Variant, ByVal ParentObj As Variant)

'This procedure centers an object over another object

    ChildObj.Move (ParentObj.Width - ChildObj.Width) / 2 + ChildObj.Left, (ParentObj.Height - ChildObj.Height) / 2 + ParentObj.Top

End Sub

Public Sub FillCombo(ByRef sCombo As ComboBox, ByVal sRS As ADODB.RecordSet, Sort As Boolean)

'This procedure fills a combo box with field name from a given recordset
'used in the combo boxes for Searching/Filtering/Sorting records

Dim x As Long

    With sCombo
        For x = 0 To sRS.Fields.Count - 1
            If sRS.Fields.Item(x).Name = "Picture" Then GoTo Von
            If Sort Then
                .AddItem "[" & sRS.Fields.Item(x).Name & "] Asc"
                .AddItem "[" & sRS.Fields.Item(x).Name & "] Desc"
            Else 'NOT SORT...
                .AddItem sRS.Fields.Item(x).Name
            End If
Von:
        Next x
    End With

End Sub

 '--------------------------
   'Special observation on cursorlocation, cursortype and locktype Properties:
   
   'For most uses, adUseClient/adOpenStatic is your best choice,
   'with adLockReadOnly as your lock type for any read-only operations
   '(export to a file, load rows to a listview, combobox, etc.)
   'and adLockOptimistic as your lock type for any read/write operations.
   'adOpenDynamic and adLockPessimistic are best suited for high-concurrency
   'situations where you need to ensure that multiple users do not corrupt each other's data.
   'While these offer the most current views of data and the most restrictive locking,
   'they do so at a severe price as far as performance is concerned.
'-------------------------------------------------------------------------
'three types of input trapping..to trap user input...

Public Function TextNum_Keypress(KeyAscii As Integer) As Integer 'text and number
If (Not ((KeyAscii >= Asc("a") And KeyAscii <= Asc("z")) Or (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z"))) And (KeyAscii <> 13 And KeyAscii <> 8 And KeyAscii <> Asc(" ") And KeyAscii <> 13 And KeyAscii <> Asc("-") And KeyAscii <> Asc("&") And KeyAscii <> Asc(".") And KeyAscii <> Asc("ñ") And KeyAscii <> Asc("Ñ") And KeyAscii <> Asc("'") And KeyAscii <> Asc("(") And KeyAscii <> Asc(",") And KeyAscii <> Asc("*") And KeyAscii <> Asc(")"))) And ((KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> Asc("-") And KeyAscii <> Asc("(") And KeyAscii <> Asc(")") And KeyAscii <> Asc(".") And KeyAscii <> Asc(":")) Then
    TextNum_Keypress = KeyAscii = 0
Else
    TextNum_Keypress = KeyAscii
End If

End Function

Public Function Text_Keypress(KeyAscii As Integer) As Integer 'text only
If (Not ((KeyAscii >= Asc("a") And KeyAscii <= Asc("z")) Or (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z"))) And (KeyAscii <> 13 And KeyAscii <> 8 And KeyAscii <> Asc(" ") And KeyAscii <> 13 And KeyAscii <> Asc("-") And KeyAscii <> Asc(".") And KeyAscii <> Asc("ñ") And KeyAscii <> Asc("Ñ") And KeyAscii <> Asc("'") And KeyAscii <> Asc("(") And KeyAscii <> Asc(",") And KeyAscii <> Asc("*") And KeyAscii <> Asc("&") And KeyAscii <> Asc(")"))) Then
    Text_Keypress = KeyAscii = 0
Else
    Text_Keypress = KeyAscii
End If
End Function

Public Function Number_Keypress(KeyAscii As Integer) As Integer 'number only
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> Asc("-") And KeyAscii <> Asc("(") And KeyAscii <> Asc(")") And KeyAscii <> Asc(".") And KeyAscii <> Asc(":") And Not (KeyAscii = 13 Or KeyAscii = 8) Then
        Number_Keypress = KeyAscii = 0
    Else
        Number_Keypress = KeyAscii
    End If
End Function

Public Sub ShellEx(PathName As String)
'Sub used to open a non-excutable file
    If ShellExecute(&O0, "Open", PathName, vbNullString, vbNullString, 1) < 33 Then
        Handler Err
    End If

End Sub

Public Sub Main()

'Provides entry point of the application

    'On Error Resume Next
        If App.PrevInstance Then
            
            MsgBox "An instance of " & App.Title & " is already running!" & vbNewLine & "You cannot run two instances of this application at the same time.", vbCritical, "Application already running"
            End
        Else 'NOT APP.PREVINSTANCE...
            frmLoading.Show
        End If

End Sub

Public Sub EnableMenuCommands()
  On Error Resume Next
    With frmMain
        'Menu Commands; setting to enabled = true
        .mnucreateuser.Enabled = True
        .mnuSettings.Enabled = True
        .mnuTrans.Enabled = True
        .mnuIssue.Enabled = True
        .mnuReturn.Enabled = True
        .mnuRec.Enabled = True
        .mnuBookRec.Enabled = True
        .mnuMemRec.Enabled = True
        .mnuReport.Enabled = True
        .mnuBookrep.Enabled = True
        .mnuMemRep.Enabled = True
        .mnuBorrowedRep.Enabled = True
        .mnuWindow.Enabled = True
        .mnuCas.Enabled = True
        .mnuTileHor.Enabled = True
        .mnuTileVer.Enabled = True
        .mnuHelp.Enabled = True
        .mnuAbout.Enabled = True
        .mnulock.Enabled = True
  
    End With
End Sub

Public Sub DisableMenuCommands()
  On Error Resume Next
    With frmMain
        .mnucreateuser.Enabled = False
        .mnuSettings.Enabled = False
        .mnuTrans.Enabled = False
        .mnuIssue.Enabled = False
        .mnuReturn.Enabled = False
        .mnuRec.Enabled = False
        .mnuBookRec.Enabled = False
        .mnuMemRec.Enabled = False
        .mnuReport.Enabled = False
        .mnuBookrep.Enabled = False
        .mnuMemRep.Enabled = False
        .mnuBorrowedRep.Enabled = False
        .mnuWindow.Enabled = False
        .mnuCas.Enabled = False
        .mnuTileHor.Enabled = False
        .mnuTileVer.Enabled = False
        .mnuHelp.Enabled = False
        .mnuAbout.Enabled = False
        .mnulock.Enabled = False
    End With
End Sub

Public Sub EnableCommandButtons()
  On Error Resume Next
    With frmMain
      .cmdBorrow.Enabled = True
      .cmdReturn.Enabled = True
      .cmdBook.Enabled = True
      .cmdAdmin.Enabled = True
      .cmdReport.Enabled = True
      .cmdSetting.Enabled = True
      .cmdHelp.Enabled = True
  
  End With
End Sub

Public Sub DisableCommandButtons()
  On Error Resume Next
    With frmMain
      .cmdBorrow.Enabled = False
      .cmdReturn.Enabled = False
      .cmdBook.Enabled = False
      .cmdAdmin.Enabled = False
      .cmdReport.Enabled = False
      .cmdSetting.Enabled = False
      .cmdHelp.Enabled = False
  
    End With
End Sub
Public Sub Relogin()
  With frmMain
    .mnuLogoff.Caption = "Relogin user..."
  End With
End Sub

Public Sub UnlockApplication()
  EnableMenuCommands
  EnableCommandButtons
  With frmMain
    .mnuFile.Enabled = True
  End With
End Sub

Public Sub LockApplication()

  DisableMenuCommands
  DisableCommandButtons
   With frmMain
    .mnuFile.Enabled = False
  End With

End Sub












 

