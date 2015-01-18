VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSurfaceArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Exclusion Class
' Provides functions for using Surface Areas in the program

Option Explicit
Implements clsObject

' Common properties
Public SceneryIndex As Integer, _
       ObjectIndex As Integer, _
       Name As String, _
       Locked As Boolean, _
       X As Single, _
       Y As Single, _
       Rotation As Single

' General properties
Public Horz As Single, _
       Vert As Single, _
       SurfaceType As Byte, _
       Height As Single

Public mDrawable As Boolean

Public Sub LoadAPTFromPolygon(FileData As clsAPTData)
  Dim Temp As String
  Temp = FileData.GetData("SurfaceType2")
  If Temp <> "" Then
    SurfaceType = ValEx(ReadNext(Temp, " "))
    If SurfaceType = 3 Then
      SurfaceType = 1
    Else
      SurfaceType = SurfaceType + 1
    End If
    Rotation = ValEx(ReadNext(Temp, " "))
    Vert = ValEx(ReadNext(Temp, " "))
    Horz = ValEx(ReadNext(Temp, " "))
    X = ValEx(ReadNext(Temp, " "))
    Y = ValEx(ReadNext(Temp, " "))
    mDrawable = True
    clsObject_UpdateObject
  End If
End Sub

Public Property Get Caption(Optional ByVal Default As Boolean = False)
  Caption = clsObject_Caption(Default)
End Property

Private Sub Class_Terminate()
'  If Closing Then Error 1
End Sub

Private Function clsObject_Add(ByVal NewX As Single, ByVal NewY As Single) As Boolean
  Dim I As Integer
  If Defaults(OT_SurfaceArea) Is Nothing Or Not Options.Remember Then
    Horz = 1000
    Vert = 1000
    SurfaceType = 1
  Else
    Defaults(OT_SurfaceArea).CopyTo Me
  End If
  X = NewX
  Y = NewY
  mDrawable = True
  clsObject_Add = clsObject_EditProperties()
End Function

Private Property Get clsObject_Caption(Optional ByVal Default As Boolean = False) As String
  If Name = "" Or Default Then
    clsObject_Caption = Lang.GetString(RES_Obj_SurfaceArea) & " " & ObjectIndex
  Else
    clsObject_Caption = Name
  End If
End Property

Private Sub clsObject_Compile(ByVal FileNum As Integer)
  Dim Corners(4) As clsLatLon
  Dim TopLeft As clsLatLon, BottomRight As clsLatLon
  Dim HalfX As Integer, HalfY As Integer
  Dim Pos As String
  Dim I As Integer
  
  Print #FileNum, "; " & clsObject_Caption
  Pos = ReturnPoint(X, Y).LatLon

  If SurfaceType = 0 Then
    HalfX = Horz / 2
    HalfY = Vert / 2
    Set Corners(0) = ReturnPoint(X - Horz / 2, Y + Vert / 2)
    Set Corners(1) = ReturnPoint(X - Horz / 2, Y - Vert / 2)
    Set Corners(2) = ReturnPoint(X + Horz / 2, Y + Vert / 2)
    Set Corners(3) = ReturnPoint(X + Horz / 2, Y - Vert / 2)
    Set TopLeft = New clsLatLon
    Set BottomRight = New clsLatLon
    
    Corners(0).CopyTo TopLeft
    Corners(0).CopyTo BottomRight
    
    For I = 1 To 3
      If Corners(I).NumLatitude > TopLeft.NumLatitude Then _
        TopLeft.NumLatitude = Corners(I).NumLatitude
      If Corners(I).NumLatitude < BottomRight.NumLatitude Then _
        BottomRight.NumLatitude = Corners(I).NumLatitude
        
      If Corners(I).NumLongitude < TopLeft.NumLongitude Then _
        TopLeft.NumLongitude = Corners(I).NumLongitude
      If Corners(I).NumLongitude > TopLeft.NumLongitude Then _
        BottomRight.NumLongitude = Corners(I).NumLongitude
    Next I

    Print #FileNum, "Area16( " & TopLeft.Latitude() & " " & BottomRight.Latitude() & " " & BottomRight.Longitude() & " " & TopLeft.Longitude() & " )"
    Print #FileNum, "    RefPoint( abs : 1 " & Pos & " )"
    If Rotation > 0 Then
      Print #FileNum, "    RotatedCall( :lblRotated 0 0 " & FloatFormat(Rotation, "##0.0#") & " )"
      Print #FileNum, "    Jump( : )"
      Print #FileNum, ":lblRotated"
    End If
    Print #FileNum, "    SenseBorder( : " & -HalfX & " " & -HalfY
    Print #FileNum, "                   " & -HalfX & " " & HalfY
    Print #FileNum, "                   " & HalfX & " " & HalfY
    Print #FileNum, "                   " & HalfX & " " & -HalfY
    Print #FileNum, "    )"
    Print #FileNum, "    SetElevation( " & CInt(Height) & " )"
    If Rotation > 0 Then
      Print #FileNum, "    Return"
    End If
    Print #FileNum, "End16" & vbCrLf
    
    For I = 0 To 3
      Set Corners(I) = Nothing
    Next I
    Set TopLeft = Nothing
    Set BottomRight = Nothing
  End If

  Print #FileNum, "Area( 5 " & Pos & " 12 )"
  Print #FileNum, "    RefPoint( rel : 1 " & Pos & " )"
  If Rotation > 0 Then
    Print #FileNum, "    RotatedCall( :lblRotated 0 0 " & FloatFormat(Rotation, "##0.0#") & " )"
    Print #FileNum, "    Jump( : )"
    Print #FileNum, ":lblRotated"
  End If
  If SurfaceType > 0 Then
    Print #FileNum, "    SurfaceType( " & SurfaceType - 1 & " " & CLng(Horz) & " " & CLng(Vert) & " 0 )"
  Else
    Print #FileNum, "    SurfaceType( 0 " & CLng(Horz) & " " & CLng(Vert) & " 0 )"
  End If
  
  If Rotation > 0 Then
    Print #FileNum, "    Return"
  End If
  Print #FileNum, "EndA" & vbCrLf

End Sub

Private Property Get clsObject_Complexity() As Byte

End Property

Private Property Let clsObject_Complexity(ByVal RHS As Byte)

End Property

Private Sub clsObject_CopyTo(Destination As clsObject)
  Dim Temp As clsSurfaceArea
  
  If TypeOf Destination Is clsSurfaceArea Then
    Set Temp = Destination
    With Temp
      .Name = Name
      .Locked = Locked
      .X = X
      .Y = Y
      .Rotation = Rotation
      .Horz = Horz
      .Vert = Vert
      .SurfaceType = SurfaceType
    End With
  End If
End Sub

Private Sub clsObject_DrawBottom(Optional ByVal DrawX As Single = -999999#, Optional ByVal DrawY As Single)

End Sub

Private Sub clsObject_DrawTop(Optional ByVal DrawX As Single = -999999#, Optional ByVal DrawY As Single)
  Dim myColor As Long
  Dim HalfH As Single, HalfV As Single
  
  If Not mDrawable Or Not Options.ObjectVisible(OT_SurfaceArea) Then Exit Sub

  glPushMatrix
  If DrawX = -999999 Then
    ' No drag
    If Scenery.Focus(SceneryIndex) Then
      myColor = COLOR_FocusColor
    Else
      myColor = OT_SurfaceArea
    End If
    glTranslatef X, Y, 0
  Else
    ' Drag
    myColor = COLOR_DragColor
    glTranslatef DrawX, DrawY, 0
  End If
  glRotatef Scenery.Header.Rotation - Rotation, 0, 0, 1
  
  Options.SetGLColor myColor
  
  HalfH = Horz / 2
  HalfV = Vert / 2
  glDrawRect -HalfH, -HalfV, HalfH, HalfV
  
  If DrawX = -999999 Then
    If Options.FocusCircle Then
      Options.SetGLColor myColor
      glDrawCircle X, Y, FocusWidth, Options.ObjectColor(myColor)
    End If
  End If
  glPopMatrix
End Sub

Private Function clsObject_EditProperties() As Boolean
  If frmExclusion.EditDataSurface(Me) Then
    clsObject_EditProperties = True
    clsObject_UpdateObject
    If Defaults(OT_SurfaceArea) Is Nothing Then Set Defaults(OT_SurfaceArea) = New clsSurfaceArea
    clsObject_CopyTo Defaults(OT_SurfaceArea)
  End If
End Function

Private Sub clsObject_LoadAPTData(FileData As clsAPTData)
  With FileData
    Height = Meter(.GetData("Altitude"))
    SurfaceType = ValEx(.GetData("SurfaceType"))
    If SurfaceType = 3 Then
      SurfaceType = 1
    Else
      SurfaceType = SurfaceType + 1
    End If
    Horz = Meter(.GetData("Length"))
    Vert = Meter(.GetData("Width"))
    Rotation = ValEx(.GetData("Orientation"))
    X = Meter(.GetData("PosX"))
    Y = Meter(.GetData("PosY"))
    Name = .GetData("Comment")
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
    .GetData "Rotation", Rotation, 0
    .GetData "Horizontal", Horz, 0
    .GetData "Vertical", Vert, 0
    .GetData "Type", SurfaceType, 255
    .GetData "Height", Height, 0
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
  clsObject_ObjectType = OT_SurfaceArea
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
    .InitializeForWrite ObjectNames(OT_SurfaceArea), Caption
    .SaveData "Name", Name, ""
    .SaveData "Locked", Locked, False
    .SaveData "X", X, 0
    .SaveData "Y", Y, 0
    .SaveData "Rotation", Rotation, 0
    .SaveData "Horizontal", Horz, 0
    .SaveData "Vertical", Vert, 0
    .SaveData "Type", SurfaceType, 255
    .SaveData "Height", Height, 0
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
  clsObject_SortIndex = &H8000& * &H110& + SceneryIndex
End Property

Private Sub clsObject_UpdateObject()
  mDrawable = True
End Sub

Private Sub clsObject_VersionChanged(ByVal NewVersion As Integer)

End Sub

