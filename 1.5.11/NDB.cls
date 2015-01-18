VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' NDB Class
' Provides functions for using NDB in the program

Option Explicit
Implements clsObject

' Common properties
Public SceneryIndex As Integer, _
       ObjectIndex As Integer, _
       Name As String, _
       Locked As Boolean, _
       X As Single, _
       Y As Single

' General properties
Public ID As String, _
       Frequency As Single, _
       Range As Byte, _
       AFDEntry As Boolean

Public mDrawable As Boolean

Public Property Get Caption(Optional ByVal Default As Boolean = False)
  Caption = clsObject_Caption(Default)
End Property

Private Sub Class_Terminate()
'  If Closing Then Error 1
End Sub

Private Function clsObject_Add(ByVal NewX As Single, ByVal NewY As Single) As Boolean
  Dim I As Integer
  If Defaults(OT_NDB) Is Nothing Or Not Options.Remember Then
    ' Default values:
    Frequency = 619
    Range = 100
    AFDEntry = 1
  Else
    Defaults(OT_NDB).CopyTo Me
  End If
  Name = Caption(True)
  ID = Replace(Name, " ", "")
  X = NewX
  Y = NewY
  mDrawable = True
  clsObject_Add = clsObject_EditProperties()
End Function

Private Property Get clsObject_Caption(Optional ByVal Default As Boolean = False) As String
  If Name = "" Or Default Then
    clsObject_Caption = Lang.GetString(RES_Obj_NDB) & " " & ObjectIndex
  Else
    clsObject_Caption = Name
  End If
End Property

Private Sub clsObject_Compile(ByVal FileNum As Integer)
  Dim TempStr As String, myLatLon As clsLatLon

  ' SCASM, FreeSC incompatibility
  If Scenery.QuoteStrings Then
    TempStr = QuoteString(Name)
  Else
    TempStr = ReplaceParens(Replace(Name, " ", "_"))
  End If
  
  Set myLatLon = ReturnPoint(X, Y)
  
  Print #FileNum, "; " & clsObject_Caption
  Print #FileNum, "NDB( " & FloatFormat(Frequency, "000.0") & " " & CLng(Range * NmToM / 2048) & " " & ID & " " & TempStr & " " & myLatLon.LatLon & " " & CInt(Scenery.Header.Altitude) & " )" & vbCrLf
  
  If ((AFDEntry And 1) > 0) Then
    With Scenery
      .AFDBuffer = .AFDBuffer & "Container( NDB " & SceneryIndex & " )" & vbCrLf & _
                 "  Navaid( " & myLatLon.LatLon & " " & FloatFormat(Scenery.Header.Altitude, "####0.0#") & " 0 " & FloatFormat(Frequency / 1000, "0.0000") & " NDB HIGHALT 0 " & ID & " )" & vbCrLf
      .AFDBuffer = .AFDBuffer & "EndC" & vbCrLf & vbCrLf
      .NameListBuffer = .NameListBuffer & "    NameEntry( NDB " & SceneryIndex & " " & QuoteString(Name) & " )" & vbCrLf
    End With
  End If
  Set myLatLon = Nothing
End Sub

Private Property Get clsObject_Complexity() As Byte

End Property

Private Property Let clsObject_Complexity(ByVal RHS As Byte)

End Property

Private Sub clsObject_CopyTo(Destination As clsObject)
  Dim Temp As clsNDB
  
  If TypeOf Destination Is clsNDB Then
    Set Temp = Destination
    With Temp
      .Name = Name
      .Locked = Locked
      .X = X
      .Y = Y
      .ID = ID
      .Frequency = Frequency
      .Range = Range
      .AFDEntry = AFDEntry
    End With
  End If
End Sub

Private Sub clsObject_DrawBottom(Optional ByVal DrawX As Single = -999999#, Optional ByVal DrawY As Single)

End Sub

Private Sub clsObject_DrawTop(Optional ByVal DrawX As Single = -999999#, Optional ByVal DrawY As Single)
  Dim myColor As Long, Buffer As Single, _
    Pt As PointType
  
  If Not mDrawable Or Not Options.ObjectVisible(OT_NDB) Then Exit Sub

  glPushMatrix
  If DrawX = -999999 Then
    ' No drag
    If Scenery.Focus(SceneryIndex) Then
      myColor = COLOR_FocusColor
    Else
      myColor = OT_NDB
    End If
    glTranslatef X, Y, 0
    DrawX = X
    DrawY = Y
  Else
    ' Drag
    myColor = COLOR_DragColor
    glTranslatef DrawX, DrawY, 0
  End If
  
  Options.SetGLColor myColor
  
  glDrawCircle DrawX, DrawY, picEditor.ScaleX(8), Options.ObjectColor(myColor)
  glDrawCircle DrawX, DrawY, picEditor.ScaleX(12), Options.ObjectColor(myColor)
  
  glPopMatrix
End Sub

Private Function clsObject_EditProperties() As Boolean
  If frmBeacon.EditDataNDB(Me) Then
    clsObject_EditProperties = True
    clsObject_UpdateObject
    If Defaults(OT_NDB) Is Nothing Then Set Defaults(OT_NDB) = New clsNDB
    clsObject_CopyTo Defaults(OT_NDB)
  End If
End Function

Private Sub clsObject_LoadAPTData(FileData As clsAPTData)
  Dim Temp As String
  
  With FileData
    X = .GetData("PosX")
    Y = .GetData("PosY")
    Frequency = ValEx(.GetData("Frequence"))
    Range = Nautical(.GetData("Portee"))
    Temp = .GetData("Nom")
    ID = ReadNext(Temp, " ")
    Name = Replace(Temp, "_", " ")
  End With
  mDrawable = True
  clsObject_UpdateObject
End Sub

Private Sub clsObject_LoadBinaryData(ByVal FileNum As Integer, ByVal Version As Integer, ByVal UserData As Long)
  Dim DummyByte1 As Byte, DummyByte2 As Byte, _
    DummyStr As String, _
    DummyStr1 As String, DummyStr2 As String, _
    DummyStr3 As String, DummyStr4 As String
  Select Case Version
    Case Is >= 8
      GetBinaryData FileNum, Name, X, Y, Locked, _
        Frequency, Range, DummyStr, ID, DummyStr1, _
        DummyStr2, DummyStr3, DummyStr4, _
        DummyByte1, DummyByte2, AFDEntry
    Case Else
      GetBinaryData FileNum, X, Y, Frequency, _
        Range, DummyStr, ID, DummyStr1, _
        DummyStr2, DummyStr3, DummyStr4, _
        DummyByte1, DummyByte2
  End Select
  
  If AFDEntry Then Scenery.AFDRefresh = True
  
  mDrawable = True
  clsObject_UpdateObject
End Sub

Private Sub clsObject_LoadData(FileData As clsFileData, ByVal Version As Integer)
  With FileData
    .GetData "Name", Name, ""
    .GetData "Locked", Locked, False
    .GetData "X", X, 0
    .GetData "Y", Y, 0
    .GetData "ID", ID, ""
    .GetData "Frequency", Frequency, 619
    .GetData "Range", Range, 100
    .GetData "afdentry", AFDEntry, 1
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
  clsObject_ObjectType = OT_NDB
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
    .InitializeForWrite ObjectNames(OT_NDB), Caption
    .SaveData "Name", Name, ""
    .SaveData "Locked", Locked, False
    .SaveData "X", X, 0
    .SaveData "Y", Y, 0
    .SaveData "ID", ID, ""
    .SaveData "Frequency", Frequency, 619
    .SaveData "Range", Range, 100
    .SaveData "afdentry", AFDEntry, 1
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
