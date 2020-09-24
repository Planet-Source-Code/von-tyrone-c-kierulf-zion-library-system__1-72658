VERSION 5.00
Begin VB.UserControl Photo 
   BackStyle       =   0  'Transparent
   ClientHeight    =   1275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1290
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   1275
   ScaleWidth      =   1290
   Begin VB.Image Def 
      Height          =   240
      Left            =   855
      Top             =   855
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Photo 
      BorderStyle     =   1  'Fixed Single
      Height          =   1185
      Left            =   120
      OLEDropMode     =   1  'Manual
      Stretch         =   -1  'True
      Top             =   45
      Width           =   1095
   End
End
Attribute VB_Name = "Photo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------
'Photo OCX Ver 1.0
'Support ADO and DAO.
'
'Has only be tested on an access database.!!!!!!!
'
'Designed by Von kierulf Softcode solutions


'------------------------------------------------------------------

Enum Connect
    useAdo = 1
    useDao = 2
End Enum



Dim DataFile As Integer, FileLength As Long, Chunks As Integer
Dim SmallChunks As Integer, Chunk() As Byte, i As Integer
Const ChunkSize As Integer = 1024
Public PhotoFileName As String
Public Event OnPhotoSaving(Succeded As Boolean, Filename As String)
Public Event OnPhotoLoading(IsPicture As Boolean, ErrorDescription As String)
'Public Event Click()
Const m_def_ConnectionType = 1
Dim m_ConnectionType As Connect
'Event Declarations:
Event Click() 'MappingInfo=Photo,Photo,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."


Public Sub Reset()
    '---------------------------------------------
    'Clear the Photo picture box
    '---------------------------------------------
    Photo.Picture = LoadPicture("")
End Sub

Public Sub Refresh()
    '---------------------------------------------
    'Load the current imagefile into the picture box
    '---------------------------------------------
    If Len(PhotoFileName) > 0 Then Photo.Picture = LoadPicture(PhotoFileName)
End Sub

Public Function OpenPhotoFile() As String
Dim Filter As String
Dim Filename As String
On Error GoTo Out
    '---------------------------------------------
    'Open a common dialog whitout ocx to browse
    'for an image file
    '---------------------------------------------

    Filter = "Pictures(*.bmp;*.ico;*.gif;*.jpg)|*.bmp;*.ico;*.gif;*.jpg|All Files (*.*)|*.*"
    PhotoFileName = OpenFile(Filter, "Select Photo Image", App.Path)
    OpenPhotoFile = PhotoFileName
    Photo.Picture = LoadPicture(PhotoFileName)
Exit Function
Out:
    MsgBox Err.Description
End Function

Public Sub SavePhoto(Fieldname As Field)
Dim RS As RecordSet
On Error GoTo Out

'---------------------------------------------
' If there is no image file exits
'---------------------------------------------
If Len(PhotoFileName) = 0 Then Exit Sub
DataFile = 1

'---------------------------------------------
'Open the image file
'---------------------------------------------
Open PhotoFileName For Binary Access Read As DataFile
    FileLength = LOF(DataFile)    ' Length of data in file
    '---------------------------------------------
    'If the imagefile is empty exits
    '---------------------------------------------
    If FileLength = 0 Then
  
        Close DataFile
        Exit Sub
    End If
    '---------------------------------------------
    'Calculate the bytes(Chunks)pakages to write
    '---------------------------------------------
    Chunks = FileLength \ ChunkSize
    SmallChunks = FileLength Mod ChunkSize
    '---------------------------------------------
    'Resize the chunck array to adjust the firts bytes package
    'To be copied
    '---------------------------------------------
    
    ReDim Chunk(SmallChunks)
    Get DataFile, , Chunk()
    '---------------------------------------------
    'Write the bytes to the given database fieldname
    '---------------------------------------------
    Fieldname.AppendChunk Chunk()
    '---------------------------------------------
    'Adjust the chunck array for the rest bytes
    'packages to be copied
    '---------------------------------------------
    ReDim Chunk(ChunkSize)
    For i = 1 To Chunks
        Get DataFile, , Chunk()
        Fieldname.AppendChunk Chunk()
    Next i
Close DataFile
RaiseEvent OnPhotoSaving(True, PhotoFileName)
Exit Sub
Out:
RaiseEvent OnPhotoSaving(False, PhotoFileName)
End Sub


Public Function LoadPhoto(Fieldname As Field) As String

Dim lngOffset As Long
Dim lngTotalSize As Long
Dim strChunk As String


On Error GoTo Out

DataFile = 1

Open App.Path & "\RscPic.tmp" For Binary Access Write As DataFile
   '============================================
   'Support ado and Dao
   'Choose the right command according to user connection type
   '============================================
   If m_ConnectionType = useAdo Then
        lngTotalSize = Fieldname.ActualSize
    Else
        lngTotalSize = DaoFieldSize(Fieldname)
    End If
    
    Chunks = lngTotalSize \ ChunkSize
    SmallChunks = lngTotalSize Mod ChunkSize
        
        ReDim Chunk(ChunkSize)
            '============================================
            'Support ado and Dao
            'Choose the right command according to user connection type
            '============================================
            
        If m_ConnectionType = useDao Then
            Chunk() = GetDaoChunk(Fieldname)
        Else
            Chunk() = Fieldname.GetChunk(ChunkSize)
        End If
        
        Put DataFile, , Chunk()
        lngOffset = lngOffset + ChunkSize
        
        Do While lngOffset < lngTotalSize
            '============================================
            'Support ado and Dao
            'Choose the right command according to user connection type
            '============================================
            
            If m_ConnectionType = useAdo Then
                 Chunk() = Fieldname.GetChunk(ChunkSize)
            Else
                 Chunk() = GetDaoChunk(Fieldname)
            End If
            Put DataFile, , Chunk()
            lngOffset = lngOffset + ChunkSize
        Loop
Close DataFile
'============================================
' Pass the image file location to the function
'============================================
LoadPhoto = App.Path & "\RscPic.tmp"

'============================================
'Load the picture into the image box
'============================================

Photo.Picture = LoadPicture(App.Path & "\RscPic.tmp")
RaiseEvent OnPhotoLoading(True, "")

Exit Function

Out:
Err.Clear
RaiseEvent OnPhotoLoading(False, Err.Description)

End Function

'The fallowing function retrieve the fieldsize when
'Using a dao connection
Private Function DaoFieldSize(Fieldname As DAO.Field) As Long
Dim lngOffset As Long
    DaoFieldSize = Fieldname.FieldSize
End Function

'The fallowing function retrieve the Chunk data when
'Using a dao connection
Private Function GetDaoChunk(Fieldname As DAO.Field)
Dim lngOffset As Long
    GetDaoChunk = Fieldname.GetChunk(lngOffset, ChunkSize)
End Function
'
'Private Sub Photo_Click()
'RaiseEvent Click
'End Sub

'The fallowing Sub  set the frame and resize it
'To the user size
Private Sub UserControl_Resize()
Photo.Move 20, 20, UserControl.Width - 20, UserControl.Height - 20
sHwnd = UserControl.hWnd
End Sub

Private Sub UserControl_InitProperties()
    m_ConnectionType = m_def_ConnectionType
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_ConnectionType = PropBag.ReadProperty("ConnectionType", m_def_ConnectionType)
    Photo.Stretch = PropBag.ReadProperty("Stretch", True)
'    Photo.BorderStyle = PropBag.ReadProperty("BackStyle", 0)
    Photo.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Stretch", Photo.Stretch, True)
    Call PropBag.WriteProperty("ConnectionType", m_ConnectionType, m_def_ConnectionType)
'    Call PropBag.WriteProperty("BackStyle", Photo.BorderStyle, 0)
    Call PropBag.WriteProperty("BorderStyle", Photo.BorderStyle, 1)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
End Sub

Public Property Get ConnectionType() As Connect
Attribute ConnectionType.VB_Description = "Return which connection type is used ADO or Dao"
    ConnectionType = m_ConnectionType
End Property

Public Property Let ConnectionType(ByVal New_ConnectionType As Connect)
    m_ConnectionType = New_ConnectionType
    PropertyChanged "ConnectionType"
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get Stretch() As Boolean
Attribute Stretch.VB_Description = "Returns/sets a value that determines whether a graphic resizes to fit the size of an Image control."
    Stretch = Photo.Stretch
End Property

Public Property Let Stretch(ByVal New_Stretch As Boolean)
    Photo.Stretch() = New_Stretch
    PropertyChanged "Stretch"
End Property
''
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Photo,Photo,-1,BorderStyle
'Public Property Get BackStyle() As Integer
'    BackStyle = Photo.BorderStyle
'End Property
'
'Public Property Let BackStyle(ByVal New_BackStyle As Integer)
'    Photo.BorderStyle() = New_BackStyle
'    PropertyChanged "BackStyle"
'End Property
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Photo,Photo,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = Photo.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    Photo.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Private Sub Photo_Click()
    RaiseEvent Click
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Photo,Photo,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = Photo.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set Photo.Picture = New_Picture
    PropertyChanged "Picture"
End Property

