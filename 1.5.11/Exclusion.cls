VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExclusion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Exclusion Class
' Provides functions for using exclusions in the program

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
Public Horz As Single, _
       Vert As Single, _
       Exclusion As Byte

Public mDrawable As Boolean

Public Property Get Caption(Optional ByVal Default As Boolean = False)
  Caption = clsObject_Caption(Default)
End Property

Public Sub Compile(ByVal FileNum As Integer)
  clsObject_Compile FileNum
End Sub

Private Sub Class_Terminate()
'  If Closing Then Error 1
End Sub

Private Function clsObject_Add(ByVal NewX As Single, ByVal NewY As Single) As Boolean
  Dim I As Integer
  If Defaults(OT_Exclusion) Is Nothing Or Not Options.Remember Then
    Horz = 1000
    Vert = 1000
    If Options.FSVersion >= Version_FS2K Then
      Exclusion = &H7
    Else
      Exclusion = &HF
    End If
  Else
    Defaults(OT_Exclusion).CopyTo Me
  End If
  X = NewX
  Y = NewY
  mDrawable = True
  clsObject_Add = clsObject_EditProperties()
End Function

Private Property Get clsObject_Caption(Optional ByVal Default As Boolean = False) As String
  If Name = "" Or Default Then
    clsObject_Caption = Lang.GetString(RES_Obj_Exclusion) & " " & ObjectIndex
  Else
    clsObject_Caption = Name
  End If
End Property

Private Sub clsObject_Compile(ByVal FileNum As Integer)
  Dim TopLeft As clsLatLon, BottomRight As clsLatLon
  Dim HalfWidth As Single, HalfHeight As Single
  Dim Pt As PointType
  
  Print #FileNum, "; " & clsObject_Caption
  Scenery.ExcludeCount = Scenery.ExcludeCount + 1
  
  ' Make adjustment since ReturnPoint automatically adjusts for rotation
  Pt = MakePoint(Horz / 2, Vert / 2)
  Rotate Pt, Scenery.Header.Rotation
  Set TopLeft = ReturnPoint(X - Pt.X, Y + Pt.Y)
  Set BottomRight = ReturnPoint(X + Pt.X, Y - Pt.Y)

  Print #FileNum, "GenExcl( " & QuoteString(AddDir(GetShortName(Scenery.ExportPath), "Scenery\" & GetFileTitle(Scenery.FileBase) & "_exc" & Scenery.ExcludeCount) & ".bgl") & " " & Hex$(IIf(Options.FSVersion >= Version_FS2K, Exclusion And 7, Exclusion And 15)) & vbCrLf & _
    "         " & TopLeft.Latitude() & " " & BottomRight.Latitude() & " " & BottomRight.Longitude() & " " & TopLeft.Longitude() & vbCrLf & _
    "         " & TopLeft.Latitude() & " " & BottomRight.Latitude() & " " & BottomRight.Longitude() & " " & TopLeft.Longitude() & vbCrLf & _
    ")" & vbCrLf

  Set TopLeft = Nothing
  Set BottomRight = Nothing
End Sub

Private Property Get clsObject_Complexity() As Byte

End Property

Private Property Let clsObject_Complexity(ByVal RHS As Byte)

End Property

Private Sub clsObject_CopyTo(Destination As clsObject)
  Dim Temp As clsExclusion
  
  If TypeOf Destination Is clsExclusion Then
    Set Temp = Destination
    With Temp
      .Name = Name
      .Locked = Locked
      .X = X
      .Y = Y
      .Horz = Horz
      .Vert = Vert
      .Exclusion = Exclusion
    End With
  End If
End Sub

Private Sub clsObject_DrawBottom(Optional ByVal DrawX As Single = -999999#, Optional ByVal DrawY As Single)

End Sub

Private Sub clsObject_DrawTop(Optional ByVal DrawX As Single = -999999#, Optional ByVal DrawY As Single)
  Dim myColor As Long
  Dim HalfH As Single, HalfV As Single
  
  If Not mDrawable Or Not Options.ObjectVisible(OT_Exclusion) Then Exit Sub

  glPushMatrix
  If DrawX = -999999 Then
    ' No drag
    If Scenery.Focus(SceneryIndex) Then
      myColor = COLOR_FocusColor
    Else
      myColor = OT_Exclusion
    End If
    glTranslatef X, Y, 0
  Else
    ' Drag
    myColor = COLOR_DragColor
    glTranslatef DrawX, DrawY, 0
  End If
  glRotatef Scenery.Header.Rotation, 0, 0, 1
  
  Options.SetGLColor myColor
  
  HalfH = Horz / 2
  HalfV = Vert / 2
  glDrawRect -HalfH, -HalfV, HalfH, HalfV
  glBeginLines
  glDrawLine -HalfH, -HalfV, HalfH, HalfV
  glDrawLine -HalfH, HalfV, HalfH, -HalfV
  glEnd
  
  If DrawX = -999999 Then
    If Options.FocusCircle Then
      Options.SetGLColor myColor
      glDrawCircle X, Y, FocusWidth, Options.ObjectColor(myColor)
    End If
  End If
  glPopMatrix
End Sub

Private Function clsObject_EditProperties() As Boolean
  If frmExclusion.EditData(Me) Then
    clsObject_EditProperties = True
    clsObject_UpdateObject
    If Defaults(OT_Exclusion) Is Nothing Then Set Defaults(OT_Exclusion) = New clsExclusion
    clsObject_CopyTo Defaults(OT_Exclusion)
  End If
End Function

Private Sub clsObject_LoadAPTData(FileData As clsAPTData)
  Dim strTemp As String, X1 As Single, X2 As Single, Y1 As Single, Y2 As Single
  With FileData
    Name = .GetData("Comment")
    strTemp = .GetData("Point")
    ReadNext strTemp, " "
    X1 = ReadNext(strTemp, " ")
    Y1 = ReadNext(strTemp, " ")
    strTemp = .GetData("Point")
    strTemp = .GetData("Point")
    ReadNext strTemp, " "
    X2 = ReadNext(strTemp, " ")
    Y2 = ReadNext(strTemp, " ")
    
    Exclusion = -(.GetData("EXCLATIS") = "1" And Options.FSVersion <= Version_CFS1) * 8 + _
                -(.GetData("EXCLVOR") = "1") * 2 + _
                -(.GetData("EXCLNDB") = "1") * 4 + _
                -(.GetData("EXCLObject") = "1") * 1
                
  End With
  X = (X1 + X2) / 2
  Y = (Y1 + Y2) / 2
  Horz = Abs(X1 - X2)
  Vert = Abs(Y1 - Y2)
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
    .GetData "Horizontal", Horz, 0
    .GetData "Vertical", Vert, 0
    .GetData "Exclusion", Exclusion, 15
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
  clsObject_ObjectType = OT_Exclusion
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
    .InitializeForWrite ObjectNames(OT_Exclusion), Caption
    .SaveData "Name", Name, ""
    .SaveData "Locked", Locked, False
    .SaveData "X", X, 0
    .SaveData "Y", Y, 0
    .SaveData "Horizontal", Horz, 0
    .SaveData "Vertical", Vert, 0
    .SaveData "Exclusion", Exclusion, 15
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
  clsObject_SortIndex = &H8000& * -3 + SceneryIndex
End Property

Private Sub clsObject_UpdateObject()
  mDrawable = True
End Sub

Private Sub clsObject_VersionChanged(ByVal NewVersion As Integer)

End Sub

