VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenuEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' MenuEntry Class
' Provides functions for using Menu Entries in the program

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
Public Rotation As Single

Public mDrawable As Boolean

' Drawing cache
Private Pts() As PointType

Public Property Get Caption(Optional ByVal Default As Boolean = False)
  Caption = clsObject_Caption(Default)
End Property

Private Sub Class_Terminate()
'  If Closing Then Error 1
End Sub

Private Function clsObject_Add(ByVal NewX As Single, ByVal NewY As Single) As Boolean
  Dim I As Integer
'  If Defaults(OT_MenuEntry) Is Nothing Or Not Options.Remember Then
'
'  Else
'    Defaults(OT_MenuEntry).CopyTo Me
'  End If
  X = NewX
  Y = NewY
  mDrawable = True
  clsObject_Add = clsObject_EditProperties()
End Function

Private Property Get clsObject_Caption(Optional ByVal Default As Boolean = False) As String
  If Name = "" Or Default Then
    clsObject_Caption = Lang.GetString(RES_Obj_MenuEntry) & " " & ObjectIndex
  Else
    clsObject_Caption = Name
  End If
End Property

Private Sub clsObject_Compile(ByVal FileNum As Integer)
  Dim TempStr As String
  Dim Pos As clsLatLon
  
  Print #FileNum, "; " & clsObject_Caption
  
  ' SCASM, FreeSC incompatibility
  If Scenery.QuoteStrings Then
    TempStr = QuoteString(Name)
  Else
    TempStr = ReplaceParens(Replace(Name, " ", "_"))
  End If

  Set Pos = ReturnPoint(X, Y)
  Print #FileNum, "MenuEntry( " & Pos.LatLon & " " & FloatFormat(Scenery.Header.Altitude, "####0.0#") & " " & FloatFormat(Rotation + IIf(Options.FSVersion = Version_FS2K, 0, Pos.MagVar()), "##0.0#") & " 127.30 111.50 360 111.50 360 " & TempStr & " )" & vbCrLf
  Set Pos = Nothing
End Sub

Private Property Get clsObject_Complexity() As Byte

End Property

Private Property Let clsObject_Complexity(ByVal RHS As Byte)

End Property

Private Sub clsObject_CopyTo(Destination As clsObject)
  Dim Temp As clsMenuEntry
  
  If TypeOf Destination Is clsMenuEntry Then
    Set Temp = Destination
    With Temp
      .Name = Name
      .Locked = Locked
      .X = X
      .Y = Y
      .Rotation = Rotation
    End With
  End If
End Sub

Private Sub clsObject_DrawBottom(Optional ByVal DrawX As Single = -999999#, Optional ByVal DrawY As Single)

End Sub

Private Sub clsObject_DrawTop(Optional ByVal DrawX As Single = -999999#, Optional ByVal DrawY As Single)
  Dim myColor As Long
  
  If Not mDrawable Or Not Options.ObjectVisible(OT_MenuEntry) Then Exit Sub

  glPushMatrix
  If DrawX = -999999 Then
    ' No drag
    If Scenery.Focus(SceneryIndex) Then
      myColor = COLOR_FocusColor
    Else
      myColor = OT_MenuEntry
    End If
    glTranslatef X, Y, 0
  Else
    ' Drag
    myColor = COLOR_DragColor
    glTranslatef DrawX, DrawY, 0
  End If
  glRotatef Scenery.Header.Rotation - Rotation, 0, 0, 1
  
  Options.SetGLColor myColor
  
  glDrawPolygon Pts
  
  If DrawX = -999999 Then
    If Options.FocusCircle Then
      Options.SetGLColor myColor
      glDrawCircle X, Y, FocusWidth, Options.ObjectColor(myColor)
    End If
  End If
  glPopMatrix
End Sub

Private Function clsObject_EditProperties() As Boolean
  If frmProperties.EditDataMenuEntry(Me) Then
    clsObject_EditProperties = True
    clsObject_UpdateObject
    
    ' Nothing to remember
'    If Defaults(OT_MenuEntry) Is Nothing Then Set Defaults(OT_MenuEntry) = New clsMenuEntry
'    clsObject_CopyTo Defaults(OT_MenuEntry)
  End If
End Function

Private Sub clsObject_LoadAPTData(FileData As clsAPTData)
  With FileData
    X = ValEx(.GetData("PosX"))
    Y = ValEx(.GetData("PosY"))
    Name = Replace(.GetData("Nom"), "_", " ")
    Rotation = ValEx(.GetData("Cap")) - Scenery.Header.APTMagVar
  End With
  mDrawable = True
  clsObject_UpdateObject
End Sub

Private Sub clsObject_LoadBinaryData(ByVal FileNum As Integer, ByVal Version As Integer, ByVal UserData As Long)
  Select Case Version
    Case Is >= 8
      GetBinaryData FileNum, X, Y, Locked, Rotation, Name
    Case Else
      GetBinaryData FileNum, X, Y, Rotation, Name
  End Select
  mDrawable = True
  clsObject_UpdateObject
End Sub

Private Sub clsObject_LoadData(FileData As clsFileData, ByVal Version As Integer)
  With FileData
    .GetData "Name", Name, ""
    .GetData "Locked", Locked, False
    .GetData "X", X, 0
    .GetData "Y", Y, 0
    .GetData "Rotation", Rotation, 0
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
  clsObject_ObjectType = OT_MenuEntry
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
    .InitializeForWrite ObjectNames(OT_MenuEntry), Caption
    .SaveData "Name", Name, ""
    .SaveData "Locked", Locked, False
    .SaveData "X", X, 0
    .SaveData "Y", Y, 0
    .SaveData "Rotation", Rotation, 0
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
  clsObject_SortIndex = &H8000& * -2 + SceneryIndex
End Property

Private Sub clsObject_UpdateObject()
  mDrawable = True
  
  ReDim Pts(3)
  Pts(0) = MakePoint(0, 40)
  Pts(1) = MakePoint(-20, -10)
  Pts(2) = MakePoint(20, -10)
  Pts(3) = Pts(0)
'  MultiRotate Pts(), Rotation
End Sub

Private Sub clsObject_VersionChanged(ByVal NewVersion As Integer)

End Sub

