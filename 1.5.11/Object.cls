VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Common Properties
Public SceneryIndex As Integer, ObjectIndex As Integer, _
       Name As String, Locked As Boolean, _
       PositionX As Single, PositionY As Single, _
       Rotation As Single, Complexity As Byte
       
' Standard Add function to add the object (index = NewIndex) at X, Y
' Returns False if user clicked cancel
Public Function Add(ByVal NewX As Single, ByVal NewY As Single) As Boolean

End Function

' Standard Property that returns the Tooltips caption
Public Property Get Caption(Optional ByVal Default As Boolean = False) As String

End Property

' Standard Compile function to write the SCASM code
Public Sub Compile(ByVal FileNum As Integer)

End Sub

' Copy data elements
Public Sub CopyTo(Destination As clsObject)

End Sub

' Standard DrawBottom function to draw the background
Public Sub DrawBottom(Optional ByVal DrawX As Single = -999999, Optional ByVal DrawY As Single)

End Sub

' Standard DrawTop function to draw lines
Public Sub DrawTop(Optional ByVal DrawX As Single = -999999, Optional ByVal DrawY As Single)

End Sub

' Standard EditProperties function to display the dialog box
' Returns false if user pressed cancel
Public Function EditProperties() As Boolean

End Function

' Standard LoadAPTData function to load Airport data
Public Sub LoadAPTData(FileData As clsAPTData)

End Sub

' Standard LoadBinaryData function to load binary data
Public Sub LoadBinaryData(ByVal FileNum As Integer, ByVal Version As Integer, ByVal UserData As Long)

End Sub

' Standard LoadData function to load the data
Public Sub LoadData(FileData As clsFileData, ByVal Version As Integer)

End Sub

' Standard Property that returns the object type
Public Property Get ObjectType() As ObjectTypeEnum

End Property

' Standard SaveData function to save the data
Public Sub SaveData(FileData As clsFileData)

End Sub

' Get a numeric value indicating the sort order priority
' of the object
Public Property Get SortIndex() As Long

End Property

' Something in the object has changed from an external
' command, update the drawing list
Public Sub UpdateObject()

End Sub

' The FS version has changed, update properties to
' ensure compatibility
Public Sub VersionChanged(ByVal NewVersion As Integer)

End Sub
