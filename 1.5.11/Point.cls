VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Point Class
' Provides functions for using points in the program

Option Explicit
Implements clsObject

Public Enum PointTypeEnum
  FSNormal
  FSMoveTaxi
End Enum

' Common properties
Public SceneryIndex As Integer, _
       ObjectIndex As Integer, _
       Locked As Boolean, _
       X As Single, _
       Y As Single, _
       Z As Single

' General properties
Public PtType As PointTypeEnum

' Parent Object
Public Parent As clsShape

Private DragX As Single, DragY As Single

Public mDrawable As Boolean

Public Property Get Caption(Optional ByVal Default As Boolean = False)
  Caption = clsObject_Caption(Default)
End Property

Public Sub DrawBottom(Optional ByVal DrawX As Single = -999999#, Optional ByVal DrawY As Single)
  clsObject_DrawBottom DrawX, DrawY
End Sub

Public Sub DrawTop(Optional ByVal DrawX As Single = -999999#, Optional ByVal DrawY As Single)
  clsObject_DrawTop DrawX, DrawY
End Sub

Public Sub GetPoint(ByVal Color As Long, ByRef NewX As Single, ByRef NewY As Single)
  If Color = COLOR_DragColor Then
    NewX = DragX
    NewY = DragY
  Else
    NewX = X
    NewY = Y
  End If
End Sub

Public Sub LoadAPTData(FileData As clsAPTData)
  clsObject_LoadAPTData FileData
End Sub

Public Sub SavePointData(FileData As clsFileData)
  With FileData
    .InitializeForWrite ObjectNames(OT_Point), Caption
    .SaveData "Locked", Locked, False
    .SaveData "X", X, 0
    .SaveData "Y", Y, 0
    .SaveData "Z", Z, 0
    .SaveData "Type", PtType, 0
    .WriteSection
  End With
End Sub

Public Sub SetAsNewPoint(NewParent As clsShape, ByVal NewX As Single, ByVal NewY As Single, NewIndex As Integer)
  X = NewX
  Y = NewY
  ObjectIndex = NewIndex
  Set Parent = NewParent
  mDrawable = True
  Scenery.AddPoint Me
End Sub

Public Sub UpdateObject()
  If ObjectIndex = 1 Then Parent.UpdateObject
End Sub

Private Sub Class_Terminate()
  Set Parent = Nothing
  If Closing Then Error 1
End Sub

Private Function clsObject_Add(ByVal NewX As Single, ByVal NewY As Single) As Boolean
  Parent.FillDefaults
  X = NewX
  Y = NewY
  mDrawable = True
  If clsObject_EditProperties() Then
    clsObject_Add = True
    Parent.AddSuccess Me
  End If
End Function

Private Property Get clsObject_Caption(Optional ByVal Default As Boolean = False) As String
  If Default Then
    clsObject_Caption = Lang.GetString(RES_Obj_Point) & " " & ObjectIndex
  Else
    clsObject_Caption = Lang.ResolveString(Res_Obj_Point2, ObjectIndex, Parent.Caption)
  End If
End Property

Private Sub clsObject_Compile(ByVal FileNum As Integer)
  If ObjectIndex = 1 Then
    Parent.Compile FileNum
  End If
End Sub

Private Property Get clsObject_Complexity() As Byte
  ' Nothing
End Property

Private Property Let clsObject_Complexity(ByVal RHS As Byte)
  ' Nothing
End Property

Private Sub clsObject_CopyTo(Destination As clsObject)
  Dim Temp As clsPoint
  
  If TypeOf Destination Is clsPoint Then
    Set Temp = Destination
    With Temp
      .ObjectIndex = ObjectIndex
      .Locked = Locked
      .X = X
      .Y = Y
      .Z = Z
      .PtType = PtType
    End With
  End If
End Sub

Private Sub clsObject_DrawBottom(Optional ByVal DrawX As Single = -999999#, Optional ByVal DrawY As Single)
  If DrawX = -999999 Then
    If ObjectIndex = 1 Then Parent.DrawBottom
    DragX = X
    DragY = Y
  Else
    DragX = DrawX
    DragY = DrawY
  End If
End Sub

Private Sub clsObject_DrawTop(Optional ByVal DrawX As Single = -999999#, Optional ByVal DrawY As Single)
  Dim myColor As Long
  
'  glPushMatrix
  If DrawX = -999999 Then
    If ObjectIndex = 1 Then Parent.DrawTop
    
    If Not mDrawable Or Not Options.ObjectVisible(OT_Point) Then Exit Sub
    
    ' No drag
    If Scenery.Focus(SceneryIndex) Then
      myColor = COLOR_FocusColor
    Else
      myColor = OT_Point
    End If
'    glTranslatef X, Y, 0
    If Options.FocusCircle Then glDrawCircle X, Y, FocusWidth, Options.ObjectColor(myColor)
    If Options.PointCircle Then glDrawCircle X, Y, FocusWidth / 5, Options.ObjectColor(myColor)
  Else
  
    If Not mDrawable Or Not Options.ObjectVisible(OT_Point) Then Exit Sub
    
    ' Drag
'    glTranslatef DrawX, DrawY, 0
    Options.SetGLColor COLOR_DragColor
    If Options.PointCircle Then glDrawCircle DragX, DragY, FocusWidth / 5, Options.ObjectColor(COLOR_DragColor)
  End If
'  glPopMatrix
End Sub

Private Function clsObject_EditProperties() As Boolean
  Dim Result As Boolean
  If Parent.ShapeType = OT_FlatArea Then
    Result = frmFlatArea.EditData(Me)
  ElseIf Parent.ShapeType = OT_TaxiwayLine Then
    Result = frmFlatArea.EditDataTaxiLine(Me)
  Else
    Result = frmProperties.EditData(Me)
  End If
    
  If Result Then
    clsObject_EditProperties = True
    Parent.UpdateObject
    If Defaults(Parent.ShapeType) Is Nothing Then Set Defaults(Parent.ShapeType) = New clsShape
    Parent.CopyTo Defaults(Parent.ShapeType)
  End If
End Function

Private Sub clsObject_LoadAPTData(FileData As clsAPTData)
  Dim Temp As String
  Temp = FileData.GetData("Point")
  ObjectIndex = ValEx(ReadNext(Temp, " "))
  X = ValEx(ReadNext(Temp, " "))
  Y = ValEx(ReadNext(Temp, " "))
  PtType = IIf(Temp = "1", FSMoveTaxi, FSNormal)

  mDrawable = True
End Sub

Private Sub clsObject_LoadBinaryData(ByVal FileNum As Integer, ByVal Version As Integer, ByVal UserData As Long)
  Dim A As Byte, PrIndex As Integer
  Select Case Version
    Case Is >= 8
      GetBinaryData FileNum, X, Y, Locked, ObjectIndex, A, PrIndex
    Case Else
      GetBinaryData FileNum, X, Y, ObjectIndex, A, PrIndex
  End Select
'  If Version = 0 Then If A = 2 Then A = 3
'  If Version < 3 Then B = B - 1
  
  ' FSLineTo -> FSNormal
  ' FSMoveTo -> FSMoveTaxi
  ' FSDotted -> FSNormal
  ' FSTaxiway -> FSMoveTaxi
  PtType = CLng(A) Mod 2
  
  ' Dummy store
  SceneryIndex = PrIndex
  
  mDrawable = True
End Sub

Private Sub clsObject_LoadData(FileData As clsFileData, ByVal Version As Integer)
  With FileData
    .GetData "Locked", Locked, False
    .GetData "X", X, 0
    .GetData "Y", Y, 0
    .GetData "Z", Z, 0
    .GetData "Type", PtType, 0
  End With
  mDrawable = True
End Sub

Private Property Get clsObject_Locked() As Boolean
  clsObject_Locked = Locked
End Property

Private Property Let clsObject_Locked(ByVal RHS As Boolean)
  Locked = RHS
End Property

Private Property Get clsObject_Name() As String
  ' Nothing
End Property

Private Property Let clsObject_Name(ByVal RHS As String)
  ' Nothing
End Property

Private Property Get clsObject_ObjectIndex() As Integer
  clsObject_ObjectIndex = ObjectIndex
End Property

Private Property Let clsObject_ObjectIndex(ByVal RHS As Integer)
  ObjectIndex = RHS
End Property

Private Property Get clsObject_ObjectType() As ObjectTypeEnum
  clsObject_ObjectType = OT_Point
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
  ' Nothing
End Property

Private Property Let clsObject_Rotation(ByVal RHS As Single)
  ' Nothing
End Property

Private Sub clsObject_SaveData(FileData As clsFileData)
  If ObjectIndex = 1 Then Parent.SaveData FileData
End Sub

Private Property Get clsObject_SceneryIndex() As Integer
  clsObject_SceneryIndex = SceneryIndex
End Property

Private Property Let clsObject_SceneryIndex(ByVal RHS As Integer)
  SceneryIndex = RHS
End Property

Private Property Get clsObject_SortIndex() As Long
  clsObject_SortIndex = &H8000& * Parent.Layer + SceneryIndex
End Property

Private Sub clsObject_UpdateObject()
  mDrawable = True
  UpdateObject
End Sub

Private Sub clsObject_VersionChanged(ByVal NewVersion As Integer)
  If ObjectIndex = 1 Then Parent.VersionChanged NewVersion
End Sub

