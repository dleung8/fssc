VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' VOR Class
' Provides functions for using VOR in the program

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
       Flags As Byte

Public mDrawable As Boolean

Public Property Get Caption(Optional ByVal Default As Boolean = False)
  Caption = clsObject_Caption(Default)
End Property

Private Sub Class_Terminate()
'  If Closing Then Error 1
End Sub

Private Function clsObject_Add(ByVal NewX As Single, ByVal NewY As Single) As Boolean
  Dim I As Integer
  If Defaults(OT_VOR) Is Nothing Or Not Options.Remember Then
    ' Default values:
    Frequency = 110.5
    Range = 100
    Flags = 128 Or 3
  Else
    Defaults(OT_VOR).CopyTo Me
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
    clsObject_Caption = Lang.GetString(RES_Obj_VOR) & " " & ObjectIndex
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
  Print #FileNum, "VOR( " & FloatFormat(Frequency, "000.00") & " " & CLng(Range * NmToM / 2048) & " " & (Flags And 15) & " " & FloatFormat(myLatLon.MagVar(), "##0.0#") & " " & ID & " " & TempStr & " " & myLatLon.LatLon & " " & CInt(Scenery.Header.Altitude) & " 0 )" & vbCrLf

  If ((Flags And 128) > 0) Then
    With Scenery
      .AFDBuffer = .AFDBuffer & "Container( VOR " & SceneryIndex & " )" & vbCrLf & _
                 "  Navaid( " & myLatLon.LatLon & " " & FloatFormat(Scenery.Header.Altitude, "####0.0#") & " " & FloatFormat(myLatLon.MagVar, "##0.0#") & " " & FloatFormat(Frequency, "000.00") & " VOR HIGHALT 1 " & ID & " )" & vbCrLf
      If (Flags And 1) > 0 Then .AFDBuffer = .AFDBuffer & "  Navaid( " & myLatLon.LatLon & " " & FloatFormat(Scenery.Header.Altitude, "####0.0#") & " " & FloatFormat(myLatLon.MagVar, "##0.0#") & " " & FloatFormat(Frequency, "000.00") & " DME HIGHALT 1 " & ID & " )" & vbCrLf
      .AFDBuffer = .AFDBuffer & "EndC" & vbCrLf & vbCrLf
      .NameListBuffer = .NameListBuffer & "    NameEntry( VOR " & SceneryIndex & " " & QuoteString(Name) & " )" & vbCrLf
    End With
  End If
  Set myLatLon = Nothing
End Sub

Private Property Get clsObject_Complexity() As Byte

End Property

Private Property Let clsObject_Complexity(ByVal RHS As Byte)

End Property

Private Sub clsObject_CopyTo(Destination As clsObject)
  Dim Temp As clsVOR
  
  If TypeOf Destination Is clsVOR Then
    Set Temp = Destination
    With Temp
      .Name = Name
      .Locked = Locked
      .X = X
      .Y = Y
      .ID = ID
      .Frequency = Frequency
      .Range = Range
      .Flags = Flags
    End With
  End If
End Sub

Private Sub clsObject_DrawBottom(Optional ByVal DrawX As Single = -999999#, Optional ByVal DrawY As Single)

End Sub

Private Sub clsObject_DrawTop(Optional ByVal DrawX As Single = -999999#, Optional ByVal DrawY As Single)
  Dim myColor As Long, Buffer As Single, _
    Pt As PointType
  
  If Not mDrawable Or Not Options.ObjectVisible(OT_VOR) Then Exit Sub

  glPushMatrix
  If DrawX = -999999 Then
    ' No drag
    If Scenery.Focus(SceneryIndex) Then
      myColor = COLOR_FocusColor
    Else
      myColor = OT_VOR
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
  
  Pt.Y = picEditor.ScaleX(16)
  
  glDrawCircle DrawX, DrawY, Pt.Y, Options.ObjectColor(myColor)
  
  Rotate Pt, -ReturnPoint(X, Y).MagVar() - Scenery.Header.Rotation
  glBeginLines
  glDrawLine 0, 0, Pt.X, Pt.Y
  glEnd
  
  glPopMatrix
End Sub

Private Function clsObject_EditProperties() As Boolean
  If frmBeacon.EditDataVOR(Me) Then
    clsObject_EditProperties = True
    clsObject_UpdateObject
    If Defaults(OT_VOR) Is Nothing Then Set Defaults(OT_VOR) = New clsVOR
    clsObject_CopyTo Defaults(OT_VOR)
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
    Temp = .GetData("TACAN")
    If Temp = "O" Then
      Flags = 2 Or 128
    ElseIf Temp = "" Then
      Flags = ValEx(.GetData("Flags")) Or 128
    End If
  End With
  mDrawable = True
  clsObject_UpdateObject
  mDrawable = True
  clsObject_UpdateObject
End Sub

Private Sub clsObject_LoadBinaryData(ByVal FileNum As Integer, ByVal Version As Integer, ByVal UserData As Long)
  Dim ByteData As Byte, BoolData As Boolean, _
    DummyByte2 As Byte, DummyStr As String, _
    DummyStr1 As String, DummyStr2 As String, _
    DummyStr3 As String, DummyStr4 As String
  Select Case Version
    Case Is >= 8
      GetBinaryData FileNum, DummyStr, X, Y, Locked, Frequency, _
        Range, Name, ID, DummyStr1, DummyStr2, DummyStr3, DummyStr4, _
        ByteData, DummyByte2, BoolData
    Case Else
      GetBinaryData FileNum, X, Y, Frequency, _
        Range, Name, ID, DummyStr1, DummyStr2, DummyStr3, DummyStr4, _
        ByteData, DummyByte2
  End Select
  
  Flags = -BoolData * 128 + _
          0 * 8 + _
          ByteData * 1 + _
          ByteData * 2

  If ((Flags And 128) > 0) Then Scenery.AFDRefresh = True

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
    .GetData "Frequency", Frequency, 111.5
    .GetData "Range", Range, 100
    .GetData "Flags", Flags, 0
    If Version <= 16 Then
      If (Flags And 8) > 0 Then
        Flags = Flags And (Not 8)
      Else
        Flags = Flags Or 8
      End If
    End If
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
  clsObject_ObjectType = OT_VOR
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
    .InitializeForWrite ObjectNames(OT_VOR), Caption
    .SaveData "Name", Name, ""
    .SaveData "Locked", Locked, False
    .SaveData "X", X, 0
    .SaveData "Y", Y, 0
    .SaveData "ID", ID, ""
    .SaveData "Frequency", Frequency, 111.5
    .SaveData "Range", Range, 100
    .SaveData "Flags", Flags, 0
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
