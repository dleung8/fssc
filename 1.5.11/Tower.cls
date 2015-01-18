VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTower"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Tower Class
' Provides functions for placing a tower in the program.

Option Explicit
Implements clsObject

' Common properties
Public SceneryIndex As Integer, _
       ObjectIndex As Integer, _
       Name As String, _
       Locked As Boolean, _
       X As Single, _
       Y As Single, _
       Height As Single

Private Frequencies(11) As Single

Public mDrawable As Boolean

Public Property Get COMFrequency(ByVal Index As Integer) As Single
  COMFrequency = Frequencies(Index)
End Property

Public Property Let COMFrequency(ByVal Index As Integer, ByVal RHS As Single)
  Frequencies(Index) = RHS
End Property

Public Property Get Caption(Optional ByVal Default As Boolean = False)
  Caption = clsObject_Caption(Default)
End Property

Private Sub Class_Terminate()
'  If Closing Then Error 1
End Sub

Private Function clsObject_Add(ByVal NewX As Single, ByVal NewY As Single) As Boolean
  Dim I As Integer
  ' No defaults since there is only one tower
'  If Defaults(OT_TowerLocation) Is Nothing Or Not Options.Remember Then
'
'  Else
'    Defaults(OT_TowerLocation).CopyTo Me
'  End If
  X = NewX
  Y = NewY
  Height = 20
  mDrawable = True
  clsObject_Add = clsObject_EditProperties()
End Function

Private Property Get clsObject_Caption(Optional ByVal Default As Boolean = False) As String
  If Name = "" Or Default Then
    clsObject_Caption = Lang.GetString(RES_Obj_Tower) & " " & ObjectIndex
  Else
    clsObject_Caption = Name
  End If
End Property

Private Sub clsObject_Compile(ByVal FileNum As Integer)
  ' Nothing
End Sub

Private Property Get clsObject_Complexity() As Byte

End Property

Private Property Let clsObject_Complexity(ByVal RHS As Byte)

End Property

Private Sub clsObject_CopyTo(Destination As clsObject)
  Dim Temp As clsTower, I As Integer
  
  If TypeOf Destination Is clsTower Then
    Set Temp = Destination
    With Temp
      .Name = Name
      .Locked = Locked
      .X = X
      .Y = Y
      .Height = Height
      For I = 0 To 11
        .COMFrequency(I) = COMFrequency(I)
      Next I
    End With
  End If
End Sub

Private Sub clsObject_DrawBottom(Optional ByVal DrawX As Single = -999999#, Optional ByVal DrawY As Single)

End Sub

Private Sub clsObject_DrawTop(Optional ByVal DrawX As Single = -999999#, Optional ByVal DrawY As Single)
  Dim myColor As Long, Buffer As Single
  
  If Not mDrawable Or Not Options.ObjectVisible(OT_TowerLocation) Then Exit Sub

  glPushMatrix
  If DrawX = -999999 Then
    ' No drag
    If Scenery.Focus(SceneryIndex) Then
      myColor = COLOR_FocusColor
    Else
      myColor = OT_TowerLocation
    End If
    glTranslatef X, Y, 0
  Else
    ' Drag
    myColor = COLOR_DragColor
    glTranslatef DrawX, DrawY, 0
  End If
  glRotatef Scenery.Header.Rotation, 0, 0, 1
  
  Buffer = picEditor.ScaleX(2)
  
  Options.SetGLColor myColor
  glDrawRect -Buffer, -Buffer, Buffer, Buffer
  
  If DrawX = -999999 Then
    If Options.FocusCircle Then
      Options.SetGLColor myColor
      glDrawCircle X, Y, FocusWidth, Options.ObjectColor(myColor)
    End If
  End If
  glPopMatrix
End Sub

Private Function clsObject_EditProperties() As Boolean
  If frmTower.EditData(Me) Then
    clsObject_EditProperties = True
    clsObject_UpdateObject
  End If
End Function

Private Sub clsObject_LoadAPTData(FileData As clsAPTData)
  With FileData
    Frequencies(0) = ValEx(.GetData("ATISFREQ"))
    Frequencies(1) = ValEx(.GetData("MULTICOMFREQ"))
    Frequencies(2) = ValEx(.GetData("UNICOMFREQ"))
    Frequencies(3) = ValEx(.GetData("CTAFFREQ"))
    Frequencies(4) = ValEx(.GetData("GROUNDFREQ"))
    Frequencies(5) = ValEx(.GetData("TOWERFREQ"))
    Frequencies(6) = ValEx(.GetData("CLRDELFREQ"))
    Frequencies(7) = ValEx(.GetData("APPCNTLFREQ"))
    Frequencies(8) = ValEx(.GetData("DEPCNTLFREQ"))
    Frequencies(9) = ValEx(.GetData("CENTERFREQ"))
    Frequencies(10) = ValEx(.GetData("FSSFREQ"))
    Frequencies(11) = ValEx(.GetData("AWOSFREQ"))
  End With
  mDrawable = True
  clsObject_UpdateObject
End Sub

Private Sub clsObject_LoadBinaryData(ByVal FileNum As Integer, ByVal Version As Integer, ByVal UserData As Long)

End Sub

Private Sub clsObject_LoadData(FileData As clsFileData, ByVal Version As Integer)
  With FileData
    .GetData "Name", Name, ""
    .GetData "Locked", Locked, False
    .GetData "X", X, 0
    .GetData "Y", Y, 0
    .GetData "Height", Height, 0
    .GetData "ATIS", Frequencies(0), 0
    .GetData "MULTICOM", Frequencies(1), 0
    .GetData "UNICOM", Frequencies(2), 0
    .GetData "CTAF", Frequencies(3), 0
    .GetData "GNDCON", Frequencies(4), 0
    .GetData "CT", Frequencies(5), 0
    .GetData "CDEL", Frequencies(6), 0
    .GetData "APCON", Frequencies(7), 0
    .GetData "DEPCON", Frequencies(8), 0
    .GetData "CENTER", Frequencies(9), 0
    .GetData "FSS", Frequencies(10), 0
    .GetData "AWOS", Frequencies(11), 0
  End With
  mDrawable = True
  clsObject_UpdateObject
End Sub

Private Property Get clsObject_Locked() As Boolean
  clsObject_Locked = Locked
End Property

Private Property Let clsObject_Locked(ByVal RHS As Boolean)
  Locked = RHS
End Property

Private Property Get clsObject_Name() As String
  clsObject_Name = Name
End Property

Private Property Let clsObject_Name(ByVal RHS As String)
  Name = RHS
End Property

Private Property Get clsObject_ObjectIndex() As Integer
  clsObject_ObjectIndex = ObjectIndex
End Property

Private Property Let clsObject_ObjectIndex(ByVal RHS As Integer)
  ObjectIndex = RHS
End Property

Private Property Get clsObject_ObjectType() As ObjectTypeEnum
  clsObject_ObjectType = OT_TowerLocation
End Property

Private Property Get clsObject_PositionX() As Single
  clsObject_PositionX = X
End Property

Private Property Let clsObject_PositionX(ByVal RHS As Single)
  X = RHS
End Property

Private Property Get clsObject_PositionY() As Single
  clsObject_PositionY = Y
End Property

Private Property Let clsObject_PositionY(ByVal RHS As Single)
  Y = RHS
End Property

Private Property Get clsObject_Rotation() As Single

End Property

Private Property Let clsObject_Rotation(ByVal RHS As Single)

End Property

Private Sub clsObject_SaveData(FileData As clsFileData)
  With FileData
    .InitializeForWrite ObjectNames(OT_TowerLocation), Caption
    .SaveData "Name", Name, ""
    .SaveData "Locked", Locked, False
    .SaveData "X", X, 0
    .SaveData "Y", Y, 0
    .SaveData "Height", Height, 0
    .SaveData "ATIS", Frequencies(0), 0
    .SaveData "MULTICOM", Frequencies(1), 0
    .SaveData "UNICOM", Frequencies(2), 0
    .SaveData "CTAF", Frequencies(3), 0
    .SaveData "GNDCON", Frequencies(4), 0
    .SaveData "CT", Frequencies(5), 0
    .SaveData "CDEL", Frequencies(6), 0
    .SaveData "APCON", Frequencies(7), 0
    .SaveData "DEPCON", Frequencies(8), 0
    .SaveData "CENTER", Frequencies(9), 0
    .SaveData "FSS", Frequencies(10), 0
    .SaveData "AWOS", Frequencies(11), 0
    .WriteSection
  End With
End Sub

Private Property Get clsObject_SceneryIndex() As Integer
  clsObject_SceneryIndex = SceneryIndex
End Property

Private Property Let clsObject_SceneryIndex(ByVal RHS As Integer)
  SceneryIndex = RHS
End Property

Private Property Get clsObject_SortIndex() As Long
  clsObject_SortIndex = &H8000& * &HFF& + SceneryIndex
End Property

Private Sub clsObject_UpdateObject()
  mDrawable = True
End Sub

Private Sub clsObject_VersionChanged(ByVal NewVersion As Integer)

End Sub



