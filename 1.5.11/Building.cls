VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBuilding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Building Class
' Provides functions for using buildings in the program

Option Explicit
Implements clsObject

' Common properties
Public SceneryIndex As Integer, _
       ObjectIndex As Integer, _
       Name As String, _
       Locked As Boolean, _
       X As Single, _
       Y As Single, _
       Complexity As Byte

' General properties
Public Rotation As Single, _
       Length As Integer, _
       Width As Integer, _
       Altitude As Single, _
       BuildType As Byte, _
       RoofLight As Byte, _
       Synchronize As Boolean, _
       RoofLength As Integer, _
       RoofWidth As Integer

Private myLevels(3) As clsLevel

Public mDrawable As Boolean

' Drawing cache
Private Pts() As PointType, TexCoords() As PointType
Private glBitmapID As Integer

Public Property Get Caption(Optional ByVal Default As Boolean = False)
  Caption = clsObject_Caption(Default)
End Property

Public Property Get Level(ByVal Index As Integer) As clsLevel
  Set Level = myLevels(Index)
End Property

Private Sub Class_Initialize()
  Dim I As Integer
  For I = 0 To 3
    Set myLevels(I) = New clsLevel
  Next I
  glBitmapID = -1
End Sub

Private Sub Class_Terminate()
  Dim I As Integer
  For I = 0 To 3
    Set myLevels(I) = Nothing
  Next I
  If glBitmapID > -1 Then picEditor.DeleteBitmap glBitmapID
'  If Closing Then Error 1
End Sub

Private Function clsObject_Add(ByVal NewX As Single, ByVal NewY As Single) As Boolean
  Dim I As Integer
  If Defaults(OT_Building) Is Nothing Or Not Options.Remember Then
    ' Default values:
    Length = 50
    Width = 50
    If Options.FSVersion >= Version_FS2K Then
      For I = 0 To 3
        myLevels(I).SetDefault
      Next I
      myLevels(0).TexID = 8
      myLevels(1).TexID = 5
      myLevels(2).TexID = 8
      myLevels(3).TexID = 4
      myLevels(1).X = 2.5
      myLevels(1).Y = 2.5
      BuildType = 4 ' Advanced, Rectangular
    Else
      Set myLevels(0) = New clsLevel
    End If
    myLevels(0).Height = 20
    Synchronize = True
  Else
    Defaults(OT_Building).CopyTo Me
  End If
  X = NewX
  Y = NewY
  mDrawable = True
  clsObject_Add = clsObject_EditProperties()
End Function

Private Property Get clsObject_Caption(Optional ByVal Default As Boolean = False) As String
  If Name = "" Or Default Then
    clsObject_Caption = Lang.GetString(RES_Obj_Building) & " " & ObjectIndex
  Else
    clsObject_Caption = Name
  End If
End Property

Private Sub clsObject_Compile(ByVal FileNum As Integer)
  Dim Pos As String, Height As Integer, _
    V1 As Long, V2 As Long
    
  If BuildType > 3 Then
    ' Advanced Building
    Height = Level(0).Height + Level(1).Height + Level(2).Height + Level(3).Height + 1
  Else
    ' Old Building
    Height = Level(0).Height + 1
  End If
  V2 = Sqr(CLng(Width) * Width + CLng(Length) * Length + CLng(Height) * Height)
  V1 = V2 * 15
  If V1 < 6000 Then V1 = 6000
  Pos = ReturnPoint(X, Y).LatLon
  
  Print #FileNum, "; " & clsObject_Caption
  Print #FileNum, "Area( 5 " & Pos & " 12 )"
  If Complexity > 0 Then Print #FileNum, "    IfVarRange( : 346 " & Complexity & " 32767 )"
  Print #FileNum, "    PerspectiveCall( :lblPerspective )"
  Print #FileNum, "    ShadowCall( :lblDraw )"
  Print #FileNum, "    Jump( : )"
  Print #FileNum, ":lblPerspective"
  Print #FileNum, "    Perspective"
  Print #FileNum, ":lblDraw"
  Print #FileNum, "    RefPoint( rel :lblReturn 1.00 " & Pos & " V1= " & V1 & " V2= " & V2 & " )"
  
  If Rotation > 0 Then
    Print #FileNum, "    RotatedCall( :lblRotated 0 0 " & FloatFormat(Rotation, "##0.0#") & " )"
    Print #FileNum, "    Return"
    Print #FileNum, ":lblRotated"
  End If
  
  If RoofLight > 0 Then
    Print #FileNum, "    VectorJump( :lblBelow m 0 32767 0 " & Height & " )"
    Print #FileNum, "    Call( :lblBuilding )"
    Print #FileNum, "    Call( :lblRoofLight )"
    Print #FileNum, "    Return"
    Print #FileNum, ":lblBelow"
    Print #FileNum, "    Call( :lblRoofLight )"
    Print #FileNum, "    Call( :lblBuilding )"
    Print #FileNum, "    Return"
    Print #FileNum, ":lblBuilding"
  End If

  Select Case BuildType
    Case 0, 1, 2, 3
      ' Old Building
      Print #FileNum, "    Building( 0 " & CInt(Altitude) & " 0 " & (CInt(Level(0).Height) \ 4) & " " & Width & " " & Length & " " & Hex$(Level(0).TexID + BuildType * 8) & " FFFF )"
    Case 4
      ' NORMAL
      Print #FileNum, "    AdvBldg( NORMAL " & Width & "    " & Length
      If Level(0).Height > 0 Then _
        Print #FileNum, "             LEVEL1 0x" & Hex$(Level(0).TexID) & vbTab & Level(0).Height & vbTab & Format$(Level(0).Y * 256, "000") & "     " & Format$(Level(0).X * 256, "000")
      If Level(1).Height > 0 Then _
        Print #FileNum, "             LEVEL2 0x" & Hex$(Level(1).TexID) & vbTab & Level(1).Height & vbTab & Format$(Level(1).Y * 256, "000") & " " & Format$(Level(1).Z * 256, "000") & " " & Format$(Level(1).X * 256, "000")
      If Level(2).Height > 0 Then _
        Print #FileNum, "             LEVEL3 0x" & Hex$(Level(2).TexID) & vbTab & Level(2).Height & vbTab & Format$(Level(2).Y * 256, "000") & "     " & Format$(Level(2).X * 256, "000")
      Print #FileNum, "             ROOF   0x" & Hex$(Level(3).TexID) & vbTab & vbTab & Format$(Level(3).Y * 256, "000") & "     " & Format$(Level(3).X * 256, "000")
      Print #FileNum, "    )"
    Case 5
      ' PEAKED
      Print #FileNum, "    AdvBldg( PEAKED " & Width & "    " & Length
      If Level(0).Height > 0 Then _
        Print #FileNum, "             LEVEL1 0x" & Hex$(Level(0).TexID) & vbTab & Level(0).Height & vbTab & Format$(Level(0).Y * 256, "000") & "     " & Format$(Level(0).X * 256, "000")
      If Level(1).Height > 0 Then _
        Print #FileNum, "             LEVEL2 0x" & Hex$(Level(1).TexID) & vbTab & Level(1).Height & vbTab & Format$(Level(1).Y * 256, "000") & " " & Format$(Level(1).Z * 256, "000") & " " & Format$(Level(1).X * 256, "000")
      If Level(2).Height > 0 Then _
        Print #FileNum, "             LEVEL3 0x" & Hex$(Level(2).TexID) & vbTab & Level(2).Height & vbTab & Format$(Level(2).Y * 256, "000") & "     " & Format$(Level(2).X * 256, "000")
      If Level(3).Height > 0 Then _
        Print #FileNum, "             ROOF   0x" & Hex$(Level(3).TexID) & vbTab & vbTab & Format$(Level(3).Y * 256, "000") & "     " & Format$(Level(3).X * 256, "000") & vbTab & Level(3).Height & vbTab & Format$(Level(3).Z * 256, "000")
      Print #FileNum, "    )"
    Case 6
      ' SLOPED
      Print #FileNum, "    AdvBldg( SLOPED " & Width & "    " & Length & "    " & RoofWidth & "    " & RoofLength
      If Level(0).Height > 0 Then _
        Print #FileNum, "             LEVEL1 0x" & Hex$(Level(0).TexID) & vbTab & Level(0).Height & vbTab & Format$(Level(0).Y * 256, "000") & "     " & Format$(Level(0).X * 256, "000")
      If Level(1).Height > 0 Then _
        Print #FileNum, "             LEVEL2 0x" & Hex$(Level(1).TexID) & vbTab & Level(1).Height & vbTab & Format$(Level(1).Y * 256, "000") & " " & Format$(Level(1).Z * 256, "000") & " " & Format$(Level(1).X * 256, "000")
      If Level(2).Height > 0 Then _
        Print #FileNum, "             LEVEL3 0x" & Hex$(Level(2).TexID) & vbTab & Level(2).Height & vbTab & Format$(Level(2).Y * 256, "000") & "     " & Format$(Level(2).X * 256, "000")
      Print #FileNum, "             ROOF   0x" & Hex$(Level(3).TexID) & vbTab & vbTab & Format$(Level(3).Y * 256, "000") & "     " & Format$(Level(3).X * 256, "000")
      Print #FileNum, "    )"
    Case Else
      ' POLY
      Print #FileNum, "    AdvBldg( POLY " & (BuildType - 4) & " " & Width & "    " & Length
      If Level(0).Height > 0 Then _
        Print #FileNum, "             LEVEL1 0x" & Hex$(Level(0).TexID) & vbTab & Level(0).Height & vbTab & Format$(Level(0).Y * 256, "000")
      If Level(1).Height > 0 Then _
        Print #FileNum, "             LEVEL2 0x" & Hex$(Level(1).TexID) & vbTab & Level(1).Height & vbTab & Format$(Level(1).Y * 256, "000") & " " & Format$(Level(1).X * 256, "000")
      If Level(2).Height > 0 Then _
        Print #FileNum, "             LEVEL3 0x" & Hex$(Level(2).TexID) & vbTab & Level(2).Height & vbTab & Format$(Level(2).Y * 256, "000")
      Print #FileNum, "             ROOF   0x" & Hex$(Level(3).TexID) & vbTab & Level(3).Height & vbTab & Format$(Level(3).Y * 256, "000") & " " & Format$(Level(3).X * 256, "000")
      Print #FileNum, "    )"
  End Select

  Select Case RoofLight
    Case 0
      ' None
    Case 1
      ' Steady
      Print #FileNum, "    Return"
      Print #FileNum, ":lblRoofLight"
      Print #FileNum, "    LineColor( 0F F0 )"
      Print #FileNum, "    Dot( 0 " & Height & " 0 )"
    Case 2
      ' Blinking
      Print #FileNum, "    Return"
      Print #FileNum, ":lblRoofLight"
      Print #FileNum, "    IfVarAnd( :lblReturn 30A 14 )"
      Print #FileNum, "    LineColor( 0F F0 )"
      Print #FileNum, "    Dot( 0 " & Height & " 0 )"
  End Select

  Print #FileNum, ":lblReturn"
  Print #FileNum, "    Return"
  Print #FileNum, "EndA" & vbCrLf
End Sub

Private Property Get clsObject_Complexity() As Byte
  clsObject_Complexity = Complexity
End Property

Private Property Let clsObject_Complexity(ByVal RHS As Byte)
  Complexity = RHS
End Property

Private Sub clsObject_CopyTo(Destination As clsObject)
  Dim Temp As clsBuilding, I As Integer
  
  If TypeOf Destination Is clsBuilding Then
    Set Temp = Destination
    With Temp
      .Name = Name
      .Locked = Locked
      .X = X
      .Y = Y
      .Complexity = Complexity
      .Length = Length
      .Width = Width
      .Rotation = Rotation
      .Altitude = Altitude
      .BuildType = BuildType
      .RoofLight = RoofLight
      .RoofLength = RoofLength
      .RoofWidth = RoofWidth
      .Synchronize = Synchronize
      For I = 0 To 3
        myLevels(I).CopyTo .Level(I)
      Next I
    End With
  End If
End Sub

Private Sub clsObject_DrawBottom(Optional ByVal DrawX As Single = -999999#, Optional ByVal DrawY As Single)
  Dim myColor As Long
  
  If Not mDrawable Or Not Options.ObjectVisible(OT_Building) Then Exit Sub

  glPushMatrix
  If DrawX = -999999 Then
    ' No drag
    If Scenery.Focus(SceneryIndex) Then
      myColor = COLOR_FocusColor
    Else
      myColor = OT_Building
    End If
    glTranslatef X, Y, 0
  Else
    ' Drag
    myColor = COLOR_DragColor
    glTranslatef DrawX, DrawY, 0
  End If
  glRotatef Scenery.Header.Rotation - Rotation, 0, 0, 1
  
  Options.SetGLColor myColor
  
  If myColor <> COLOR_DragColor Then
    If UBound(TexCoords) > 0 Then
      If glBitmapID > -1 Then
        picEditor.SelectBitmap glBitmapID
        glPaintTexturedRegion2 Pts, TexCoords
      End If
    Else
      If BuildType <= 3 And Options.FillObjects Then
        glForeColor Building(myLevels(0).TexID).Color
        glPaintRegion Pts
      End If
    End If
  End If
  glPopMatrix
End Sub

Private Sub clsObject_DrawTop(Optional ByVal DrawX As Single = -999999#, Optional ByVal DrawY As Single)
  Dim myColor As Long
  
  If Not mDrawable Or Not Options.ObjectVisible(OT_Building) Then Exit Sub

  glPushMatrix
  If DrawX = -999999 Then
    ' No drag
    If Scenery.Focus(SceneryIndex) Then
      myColor = COLOR_FocusColor
    Else
      myColor = OT_Building
    End If
    glTranslatef X, Y, 0
  Else
    ' Drag
    myColor = COLOR_DragColor
    glTranslatef DrawX, DrawY, 0
  End If
  glRotatef Scenery.Header.Rotation - Rotation, 0, 0, 1
  
  Options.SetGLColor myColor
  
  If myColor <> COLOR_DragColor Then
    If UBound(TexCoords) > 0 Then
      If glBitmapID <= -1 Then
        glDrawPolygon Pts
      End If
    Else
      If BuildType > 3 Or Not Options.FillObjects Then
        glDrawPolygon Pts
      End If
    End If
  Else
    glDrawPolygon Pts
  End If
  
  If DrawX = -999999 Then
    If Options.FocusCircle Then
      Options.SetGLColor myColor
      glDrawCircle X, Y, FocusWidth, Options.ObjectColor(myColor)
    End If
  End If
  glPopMatrix
End Sub

Private Function clsObject_EditProperties() As Boolean
  If frmBuilding.EditData(Me) Then
    clsObject_EditProperties = True
    clsObject_UpdateObject
    If Defaults(OT_Building) Is Nothing Then Set Defaults(OT_Building) = New clsBuilding
    clsObject_CopyTo Defaults(OT_Building)
  End If
End Function

Private Sub clsObject_LoadAPTData(FileData As clsAPTData)
  Dim intTemp As Integer, strTemp As String
  
  With FileData
    X = .GetData("PosX")
    Y = .GetData("PosY")
    Length = .GetData("Longueur")
    Width = .GetData("Largeur")
    myLevels(0).Height = .GetData("Hauteur") * 4
    intTemp = BinaryToDec(.GetData("Type"))
    myLevels(0).TexID = intTemp And 7
    BuildType = intTemp \ 8
    Rotation = .GetData("Orientation")
    Altitude = Meter(.GetData("Altitude"))
    Complexity = .GetComplexityData()
    Name = .GetData("Comment")
    
    strTemp = .GetData("AdvType")
    If strTemp <> "" And ValEx(strTemp) >= 0 Then
      BuildType = ValEx(strTemp) + 4
      Level(0).LoadAPTData (.GetData("AdvLevel1"))
      Level(1).LoadAPTData (.GetData("AdvLevel2"))
      Level(2).LoadAPTData (.GetData("AdvLevel3"))
      Level(3).LoadAPTData (.GetData("AdvRoof"))
      strTemp = .GetData("AdvOther")
      RoofWidth = ValEx(ReadNext(strTemp, " "))
      RoofLength = ValEx(ReadNext(strTemp, " "))
      If intTemp = 3 Then
        BuildType = ValEx(ReadNext(strTemp, " ")) + 4
      End If
      Synchronize = False
    Else
      Synchronize = True
    End If
  End With
  mDrawable = True
  clsObject_UpdateObject
End Sub

Private Sub clsObject_LoadBinaryData(ByVal FileNum As Integer, ByVal Version As Integer, ByVal UserData As Long)
  Dim intAlt As Integer, ID As Byte, Height As Byte, _
    sngWidth As Single, sngLength As Single
  
  If UserData = 0 Then
    Select Case Version
      Case Is >= 8
        GetBinaryData FileNum, Name, X, Y, Locked, _
          sngWidth, sngLength, Rotation, Altitude, ID, _
          BuildType, Height, Complexity, RoofLight
      Case Is >= 6
        GetBinaryData FileNum, X, Y, sngWidth, sngLength, _
          Rotation, Altitude, ID, BuildType, Height, _
          Complexity, RoofLight
      Case Is >= 3
        GetBinaryData FileNum, X, Y, sngWidth, sngLength, _
          Rotation, Altitude, ID, BuildType, Height, _
          Complexity
'      Case Else
'        GetBinaryData FileNum, X, Y, sngWidth, sngLength, _
'          Rotation, intAlt, ID, BuildType, Height, _
'          Complexity
'        Altitude = CSng(intAlt)
    End Select
    myLevels(0).Height = Height * 4
    myLevels(0).TexID = ID
  Else
    Dim BType As Byte
    Dim B As Byte, S As Single
    Dim sngRoofL As Single, sngRoofW As Single
    
    Select Case Version
      Case Is >= 8
        GetBinaryData FileNum, Name, X, Y, Locked, sngWidth, sngLength, _
          Rotation, Altitude, BType, Complexity, _
          Level(0), Level(1), Level(2), Level(3), sngRoofW, sngRoofL, RoofLight
      Case Is >= 6
        GetBinaryData FileNum, X, Y, sngWidth, sngLength, Rotation, Altitude, _
          BType, Complexity, Level(0), Level(1), Level(2), Level(3), sngRoofW, sngRoofL, RoofLight
      Case 5
        GetBinaryData FileNum, X, Y, sngWidth, sngLength, Rotation, Altitude, _
          BType, Complexity, Level(0), Level(1), Level(2), Level(3), sngRoofW, sngRoofL
      Case Else
        GetBinaryData FileNum, X, Y, sngWidth, sngLength, Rotation, Altitude, _
          BType, Complexity
        GetBinaryData FileNum, B, S
        Level(0).TexID = B: Level(0).Height = S
        GetBinaryData FileNum, B, S
        Level(1).TexID = B: Level(1).Height = S
        GetBinaryData FileNum, B, S
        Level(2).TexID = B: Level(2).Height = S
        GetBinaryData FileNum, B, S, sngRoofW, sngRoofL
        Level(3).TexID = B: Level(3).Height = S
    End Select
    Level(0).TexID = Level(0).TexID + 8
    Level(1).TexID = Level(1).TexID + 4
    Level(2).TexID = Level(2).TexID + 8
    Level(3).TexID = Level(3).TexID + 4
    
    If BType <= 3 Then
      BuildType = BType + 4
    Else
      BuildType = BType + 5
    End If
    RoofLength = sngRoofL
    RoofWidth = sngRoofW
  End If
  
  Length = sngLength
  Width = sngWidth
  
  mDrawable = True
  clsObject_UpdateObject
End Sub

Private Sub clsObject_LoadData(FileData As clsFileData, ByVal Version As Integer)
  Dim Temp As String
  With FileData
    .GetData "Name", Name, ""
    .GetData "Locked", Locked, False
    .GetData "X", X, 0
    .GetData "Y", Y, 0
    .GetData "Length", Length, 50
    .GetData "Width", Width, 50
    .GetData "Rotation", Rotation, 0
    .GetData "Altitude", Altitude, 0
    .GetData "Building Type", BuildType, 0
    .GetData "Roof Light", RoofLight, 0
    .GetData "Complexity", Complexity, 0
    .GetData "Synchronize", Synchronize, True
    If BuildType > 3 Then
      ' Advanced
      If Version > 19 Then
        .GetData "Lobby", Temp, "8 0 1 1 1"
        myLevels(0).DataStr = Temp
        .GetData "Windows", Temp, "8 0 1 1 1"
        myLevels(1).DataStr = Temp
        .GetData "Penthouse", Temp, "8 0 1 1 1"
        myLevels(2).DataStr = Temp
        .GetData "Roof", Temp, "8 0 1 1 1"
        myLevels(3).DataStr = Temp
      Else
        .GetData "Lobby", Temp, "8 0 1 1 1"
        myLevels(0).DataStrOld = Temp
        .GetData "Windows", Temp, "8 0 1 1 1"
        myLevels(1).DataStrOld = Temp
        .GetData "Penthouse", Temp, "8 0 1 1 1"
        myLevels(2).DataStrOld = Temp
        .GetData "Roof", Temp, "8 0 1 1 1"
        myLevels(3).DataStrOld = Temp
      End If
      If BuildType = 6 Then
        ' Roof
        .GetData "RoofLength", RoofLength, 0
        .GetData "RoofWidth", RoofWidth, 0
      End If
    Else
      .GetData "Windows", Temp, "0 0 1 1 1"
      myLevels(0).DataStr = Temp
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
  clsObject_ObjectType = OT_Building
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
  clsObject_Rotation = Rotation
End Property

Private Property Let clsObject_Rotation(ByVal RHS As Single)
  Rotation = RHS
End Property

Private Sub clsObject_SaveData(FileData As clsFileData)
  With FileData
    .InitializeForWrite ObjectNames(OT_Building), Caption
    .SaveData "Name", Name, ""
    .SaveData "Locked", Locked, False
    .SaveData "X", X, 0
    .SaveData "Y", Y, 0
    .SaveData "Length", Length, 50
    .SaveData "Width", Width, 50
    .SaveData "Rotation", Rotation, 0
    .SaveData "Altitude", Altitude, 0
    .SaveData "Building Type", BuildType, 0
    .SaveData "Roof Light", RoofLight, 0
    .SaveData "Complexity", Complexity, 0
    .SaveData "Synchronize", Synchronize, True
    If BuildType > 3 Then
      ' Advanced
      .SaveData "Lobby", myLevels(0).DataStr, "8 0 1 1 1"
      .SaveData "Windows", myLevels(1).DataStr, "8 0 1 1 1"
      .SaveData "Penthouse", myLevels(2).DataStr, "8 0 1 1 1"
      .SaveData "Roof", myLevels(3).DataStr, "8 0 1 1 1"
      If BuildType = 6 Then
        ' Roof
        .SaveData "RoofLength", RoofLength, 0
        .SaveData "RoofWidth", RoofWidth, 0
      End If
    Else
      .SaveData "Windows", myLevels(0).DataStr, "0 0 1 1 1"
    End If
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
  clsObject_SortIndex = &H8000& * &H80& + SceneryIndex
End Property

Private Sub clsObject_UpdateObject()
  Dim Sides As Integer, Angle As Single, _
    DeltaAngle As Single, I As Integer, _
    HalfL As Single, HalfW As Single
    
  Dim NewglID As Integer
  NewglID = -1
    
  mDrawable = True
  
  ReDim TexCoords(0)

  HalfL = Length / 2
  HalfW = Width / 2
  
  picEditor.StartDraw
  
  If BuildType <= 6 Then
    If Not Options.FillObjects Then
    
    ElseIf Options.FSVersion >= Version_FS2K And BuildType = 0 Then
      ReDim TexCoords(3)
      TexCoords(1).X = 1
      TexCoords(2).X = 1
      TexCoords(2).Y = 1
      TexCoords(3).Y = 1
      NewglID = picEditor.LoadBitmap(AddDir(Options.FSPath, "Texture\" & BuildingR(Building(myLevels(0).TexID).Roof)))
    ElseIf BuildType >= 4 Then
      ReDim TexCoords(3)
      TexCoords(1).X = myLevels(3).X
      TexCoords(2).X = myLevels(3).X
      TexCoords(2).Y = myLevels(3).Y
      TexCoords(3).Y = myLevels(3).Y
      NewglID = picEditor.LoadBitmap(AddDir(Options.FSPath, "Texture\" & BuildingR(myLevels(3).TexID)))
    Else
      ' Nothing else needed
    End If
    
    If BuildType = 3 Then
      ' Octagonal
      ReDim Pts(10)
      Pts(0) = MakePoint(-HalfW / 2, HalfL)
      Pts(1) = MakePoint(HalfW / 2, HalfL)
      Pts(2) = MakePoint(HalfW, HalfL / 2)
      Pts(3) = MakePoint(HalfW, -HalfL / 2)
      Pts(4) = MakePoint(HalfW / 2, -HalfL)
      Pts(5) = MakePoint(-HalfW / 2, -HalfL)
      Pts(6) = MakePoint(-HalfW, -HalfL / 2)
      Pts(7) = MakePoint(-HalfW, HalfL / 2)
      Pts(8) = Pts(0)
      Pts(10) = Pts(1)
'      MultiRotate Pts, Rotation
    Else
      ReDim Pts(6)
      Pts(0) = MakePoint(-HalfW, HalfL)
      Pts(1) = MakePoint(HalfW, HalfL)
      Pts(2) = MakePoint(HalfW, -HalfL)
      Pts(3) = MakePoint(-HalfW, -HalfL)
      Pts(4) = Pts(0)
      Pts(6) = Pts(1)
'      MultiRotate Pts, Rotation
    End If

    ' Error loading bitmap
    If NewglID = -1 Then
      ReDim TexCoords(0)
    Else
      ReDim Preserve Pts(UBound(Pts) - 3)
    End If
  Else
    Sides = BuildType - 4
    ReDim Pts(Sides - 1)
    DeltaAngle = 360 / Sides
    Angle = DeltaAngle / 2 - Rotation
    
    If Options.FillObjects Then
      ReDim TexCoords(Sides - 1)
      For I = 0 To Sides - 1
        Pts(I) = MakePoint(HalfW * Cos(Angle * DegToRad), HalfL * Sin(Angle * DegToRad))
        TexCoords(I) = MakePoint((HalfW * Cos((Angle + Rotation) * DegToRad) + HalfW) * myLevels(3).X / Width, (HalfL * Sin((Angle + Rotation) * DegToRad) + HalfL) * myLevels(3).Y / Length)
        Angle = DeltaAngle + Angle
      Next I

      NewglID = picEditor.LoadBitmap(AddDir(Options.FSPath, "Texture\" & BuildingR(myLevels(3).TexID)))
    Else
      For I = 0 To Sides - 1
        Pts(I) = MakePoint(HalfL * Cos(Angle * DegToRad), HalfW * Sin(Angle * DegToRad))
        Angle = DeltaAngle + Angle
      Next I
    End If
  End If

  If glBitmapID > -1 Then picEditor.DeleteBitmap glBitmapID
  glBitmapID = NewglID
End Sub

Private Sub clsObject_VersionChanged(ByVal NewVersion As Integer)
  If NewVersion < Version_FS2K Then
    If BuildType > 3 Then
      BuildType = 0
      myLevels(0).TexID = 0
      Scenery.Changed = True
    End If
  End If
End Sub
