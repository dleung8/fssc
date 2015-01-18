VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsShapeEnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements clsObject

Private Function clsObject_Add(ByVal NewX As Single, ByVal NewY As Single) As Boolean

End Function

Private Property Get clsObject_Caption(Optional ByVal Default As Boolean = False) As String

End Property

Private Sub clsObject_Compile(ByVal FileNum As Integer)
  With Scenery
    If Scenery.PolygonLayer < 255 Then
      Print #FileNum, .PolygonBuffer;
      .PolygonBuffer = ""
      If .PolygonLayer > 0 Then
        Print #FileNum, "    Return"
      End If
      Print #FileNum, "EndA" & vbCrLf
      Print #FileNum, .RunwayBuffer;
      .RunwayBuffer = ""
      .PolygonLayer = 255
    End If
  End With
End Sub

Private Property Get clsObject_Complexity() As Byte

End Property

Private Property Let clsObject_Complexity(ByVal RHS As Byte)

End Property

Private Sub clsObject_CopyTo(Destination As clsObject)

End Sub

Private Sub clsObject_DrawBottom(Optional ByVal DrawX As Single = -999999#, Optional ByVal DrawY As Single)

End Sub

Private Sub clsObject_DrawTop(Optional ByVal DrawX As Single = -999999#, Optional ByVal DrawY As Single)

End Sub

Private Function clsObject_EditProperties() As Boolean

End Function

Private Sub clsObject_LoadAPTData(FileData As clsAPTDat)

End Sub

Private Sub clsObject_LoadBinaryData(ByVal FileNum As Integer, ByVal Version As Integer, ByVal UserData As Long)

End Sub

Private Sub clsObject_LoadData(FileData As clsFileData, ByVal Version As Integer)

End Sub

Private Property Get clsObject_Locked() As Boolean

End Property

Private Property Let clsObject_Locked(ByVal RHS As Boolean)

End Property

Private Property Get clsObject_Name() As String

End Property

Private Property Let clsObject_Name(ByVal RHS As String)

End Property

Private Property Get clsObject_ObjectIndex() As Integer

End Property

Private Property Let clsObject_ObjectIndex(ByVal RHS As Integer)

End Property

Private Property Get clsObject_ObjectType() As ObjectTypeEnum

End Property

Private Property Get clsObject_PositionX() As Single

End Property

Private Property Let clsObject_PositionX(ByVal RHS As Single)

End Property

Private Property Get clsObject_PositionY() As Single

End Property

Private Property Let clsObject_PositionY(ByVal RHS As Single)

End Property

Private Property Get clsObject_Rotation() As Single

End Property

Private Property Let clsObject_Rotation(ByVal RHS As Single)

End Property

Private Sub clsObject_SaveData(FileData As clsFileData)

End Sub

Private Property Get clsObject_SceneryIndex() As Integer

End Property

Private Property Let clsObject_SceneryIndex(ByVal RHS As Integer)

End Property

Private Property Get clsObject_SortIndex() As Long
  clsObject_SortIndex = &H8000& * &H41&
End Property

Private Sub clsObject_UpdateObject()

End Sub

Private Sub clsObject_VersionChanged(ByVal NewVersion As Integer)

End Sub

