VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBackground"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Background Class
' Provides functions for using background images
' in the program

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
Public Rotation As Single, _
       File As String, _
       ZoomX As Single, _
       ZoomY As Single, _
       LockAspectRatio As Boolean, _
       Visible As Boolean

Public mDrawable As Boolean

' Drawing cache
Private Pts() As PointType
Private TexCoords(3) As PointType
Private glBitmapID As Integer

Public Property Get Caption(Optional ByVal Default As Boolean = False)
  Caption = clsObject_Caption(Default)
End Property

Public Sub ScanTextures(ByRef Files() As String, ByRef FileCount As Integer)
  Dim W As Integer, L As Integer
  Dim Width2Up As Long, Height2Up As Long
  Dim Result As String
  Dim Dest As String, Res As Long, Pic As PicType
  
  If Visible And glBitmapID <> -1 Then
    picEditor.GetBitmapDimensions glBitmapID, W, L
    Width2Up = picEditor.PowerOf2Up(W)
    Height2Up = picEditor.PowerOf2Up(L)
    If Width2Up <> W Or Height2Up <> L Then
      If MsgBoxEx(frmMain, Lang.GetString(RES_ERR_BitmapConvert), vbInformation Or vbOKCancel, RES_ERR_BitmapConvert) = vbCancel Then
        Exit Sub
      End If

      Dest = AddDir(GetDir(File), "bkgnd" & ObjectIndex & ".bmp")
      With cDialog
        .Filter = Lang.GetString(RES_TEX_BitmapFilter)
        .FilterIndex = 1
        .DefExt = "bmp"
        Dest = .SaveDialog(Dest, Lang.GetString(RES_TEX_SaveAsBitmap))
      End With
      If Dest = "" Then
        Exit Sub
      End If

      SetScreenMousePointer vbHourglass
      LoadTexture File, Result
      If FileExists(Result) Then
        On Error Resume Next
        BmpResize Result, Result, Pic, Width2Up, Height2Up
        Res = BmptoAnys(Result, Dest, Pic, 20, 0)
        If Res > 0 Then
          File = Dest
          ZoomX = ZoomX / Width2Up * W
          ZoomY = ZoomY / Height2Up * L
          LockAspectRatio = False
          clsObject_UpdateObject
        Else
          SetScreenMousePointer vbDefault
          Exit Sub
        End If
        AddTexFile Files(), FileCount, File
      End If
      SetScreenMousePointer vbDefault
    End If
  End If
End Sub

Private Sub Class_Initialize()
  glBitmapID = -1
  TexCoords(1).X = 1
  TexCoords(2).X = 1
  TexCoords(2).Y = 1
  TexCoords(3).Y = 1
End Sub

Private Sub Class_Terminate()
  If glBitmapID > -1 Then picEditor.DeleteBitmap glBitmapID
'  If Closing Then Error 1
End Sub

Private Function clsObject_Add(ByVal NewX As Single, ByVal NewY As Single) As Boolean
  Dim I As Integer
  If Defaults(OT_Background) Is Nothing Or Not Options.Remember Then
    ' Default values:
    ZoomX = 1
    ZoomY = 1
    LockAspectRatio = True
  Else
    Defaults(OT_Background).CopyTo Me
  End If
  X = NewX
  Y = NewY
  mDrawable = True
  clsObject_Add = clsObject_EditProperties()
End Function

Private Property Get clsObject_Caption(Optional ByVal Default As Boolean = False) As String
  If Name = "" Or Default Then
    clsObject_Caption = Lang.GetString(RES_Obj_Background) & " " & ObjectIndex
  Else
    clsObject_Caption = Name
  End If
End Property

Private Sub clsObject_Compile(ByVal FileNum As Integer)
  Dim Pos As String
  If Visible Then
    If Options.FSVersion >= Version_FS2K Then
      If glBitmapID <> -1 Then
        Dim W As Integer, L As Integer
        Dim Width2Up As Long, Height2Up As Long
        picEditor.GetBitmapDimensions glBitmapID, W, L
        Width2Up = picEditor.PowerOf2Up(W)
        Height2Up = picEditor.PowerOf2Up(L)
        If Width2Up <> W Or Height2Up <> L Then
          MsgBoxEx frmMain, "Bitmap Error 1 ", vbInformation, 0
          Exit Sub
        End If
        
        Pos = ReturnPoint(X, Y).LatLon
        
        Print #FileNum, "; " & clsObject_Caption
        Print #FileNum, "Area( B " & Pos & " 40 )"
        Print #FileNum, "    LayerCall( :lblLayer )"
        Print #FileNum, "    Jump( : )"
        Print #FileNum, ":lblLayer"
        Print #FileNum, "    RefPoint( rel :lblReturn 1.00 " & Pos & " V1= 0 V2= 0 )"
        
        If Rotation > 0 Then
          Print #FileNum, "    RotatedCall( :lblRotated 0 0 " & FloatFormat(Rotation, "##0.0#") & " )"
          Print #FileNum, "    Return"
          Print #FileNum, ":lblRotated"
        End If
        Print #FileNum, "    Points( 0 " & vbCrLf & _
                        "            " & CInt(Pts(0).X) & " 0 " & CInt(Pts(0).Y) & vbCrLf & _
                        "            " & CInt(Pts(1).X) & " 0 " & CInt(Pts(1).Y) & vbCrLf & _
                        "            " & CInt(Pts(2).X) & " 0 " & CInt(Pts(2).Y) & vbCrLf & _
                        "            " & CInt(Pts(3).X) & " 0 " & CInt(Pts(3).Y) & vbCrLf & _
                        "           )"
        
          Print #FileNum, "    LoadBitmap( 0 5 0 1 1 1 " & QuoteString(GetFileTitle(File)) & " )"
          Print #FileNum, "    Smoothing( 1 )"
          Print #FileNum, "    TexPoly( au 0 0 255  1 255 255  2 255 0 3 0 0 )"
          Print #FileNum, "    Smoothing( 0 )"
      
        Print #FileNum, ":lblReturn"
        Print #FileNum, "    Return"
        Print #FileNum, "EndA" & vbCrLf
      End If
    End If
  End If
End Sub

Private Property Get clsObject_Complexity() As Byte

End Property

Private Property Let clsObject_Complexity(ByVal RHS As Byte)

End Property

Private Sub clsObject_CopyTo(Destination As clsObject)
  Dim Temp As clsBackground

  If TypeOf Destination Is clsBackground Then
    Set Temp = Destination
    With Temp
      .Name = Name
      .Locked = Locked
      .X = X
      .Y = Y
      .Rotation = Rotation
      .File = File
      .ZoomX = ZoomX
      .ZoomY = ZoomY
      .LockAspectRatio = LockAspectRatio
      .Visible = Visible
    End With
  End If
End Sub

Private Sub clsObject_DrawBottom(Optional ByVal DrawX As Single = -999999#, Optional ByVal DrawY As Single)
  Dim myColor As Long

  If Not mDrawable Or Not Options.ObjectVisible(OT_Background) Then Exit Sub

  glPushMatrix
  If DrawX = -999999 Then
    ' No drag
    If Scenery.Focus(SceneryIndex) Then
      myColor = COLOR_FocusColor
    Else
      myColor = OT_Background
    End If
    glTranslatef X, Y, 0
  Else
    ' Drag
    myColor = COLOR_DragColor
    glTranslatef DrawX, DrawY, 0
  End If
  glRotatef Scenery.Header.Rotation - Rotation, 0, 0, 1

  Options.SetGLColor myColor

  If myColor <> COLOR_DragColor And glBitmapID > -1 Then
    picEditor.SelectBitmap glBitmapID
    glPaintTexturedRegion2 Pts, TexCoords
  End If
  glPopMatrix
End Sub

Private Sub clsObject_DrawTop(Optional ByVal DrawX As Single = -999999#, Optional ByVal DrawY As Single)
  Dim myColor As Long

  If Not mDrawable Or Not Options.ObjectVisible(OT_Background) Then Exit Sub

  glPushMatrix
  If DrawX = -999999 Then
    ' No drag
    If Scenery.Focus(SceneryIndex) Then
      myColor = COLOR_FocusColor
    Else
      myColor = OT_Background
    End If
    glTranslatef X, Y, 0
  Else
    ' Drag
    myColor = COLOR_DragColor
    glTranslatef DrawX, DrawY, 0
  End If
  glRotatef Scenery.Header.Rotation - Rotation, 0, 0, 1

  Options.SetGLColor myColor

  If myColor = COLOR_DragColor Or glBitmapID <= -1 Then
    glDrawPolygon Pts
  End If

  If DrawX = -999999 Then
    If Options.FocusCircle Then
      glDrawCircle X, Y, FocusWidth, Options.ObjectColor(myColor)
    End If
  End If
  glPopMatrix
End Sub

Private Function clsObject_EditProperties() As Boolean
  If frmBackground.EditData(Me) Then
    clsObject_EditProperties = True
    clsObject_UpdateObject
    If Defaults(OT_Background) Is Nothing Then Set Defaults(OT_Background) = New clsBackground
    clsObject_CopyTo Defaults(OT_Background)
  End If
End Function

Private Sub clsObject_LoadAPTData(FileData As clsAPTData)
'  Dim Temp As Integer
'
'  With FileData
'    Name = .GetData("Comment")
'    X = .GetData("PosX")
'    Y = .GetData("PosY")
'  End With
'  mDrawable = True
End Sub

Private Sub clsObject_LoadBinaryData(ByVal FileNum As Integer, ByVal Version As Integer, ByVal UserData As Long)
  Dim MScale As Single, myFile As String
  GetBinaryData FileNum, Name, X, Y, Locked, File, MScale
  ZoomX = MScale
  ZoomY = MScale
  LockAspectRatio = True
  Visible = False
  Rotation = 0
  
  If Not FileExists(File) Then
    myFile = GetFileTitle(File)
    myFile = MultiDir(myFile, GetDir(Scenery.File))
    If myFile <> "" Then File = GetRealName(myFile)
  End If

  mDrawable = True
  clsObject_UpdateObject
End Sub

Private Sub clsObject_LoadData(FileData As clsFileData, ByVal Version As Integer)
  Dim myFile As String
  
  With FileData
    .GetData "Name", Name, ""
    .GetData "Locked", Locked, False
    .GetData "X", X, 0
    .GetData "Y", Y, 0
    .GetData "Rotation", Rotation, 0
    .GetData "File", File, ""
    .GetData "ZoomX", ZoomX, 1
    .GetData "ZoomY", ZoomY, 1
    .GetData "Lock Aspect Ratio", LockAspectRatio, True
    .GetData "Visible", Visible, False
  End With

  If Not FileExists(File) Then
    myFile = GetFileTitle(File)
    myFile = MultiDir(myFile, GetDir(Scenery.File))
    If myFile <> "" Then File = GetRealName(myFile)
  End If

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
  clsObject_ObjectType = OT_Background
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
    .InitializeForWrite ObjectNames(OT_Background), Caption
    .SaveData "Name", Name, ""
    .SaveData "Locked", Locked, False
    .SaveData "X", X, 0
    .SaveData "Y", Y, 0
    .SaveData "Rotation", Rotation, 0
    .SaveData "File", File, ""
    .SaveData "ZoomX", ZoomX, 1
    .SaveData "ZoomY", ZoomY, 1
    .SaveData "Lock Aspect Ratio", LockAspectRatio, True
    .SaveData "Visible", Visible, False
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
  clsObject_SortIndex = 0 + SceneryIndex
End Property

Private Sub clsObject_UpdateObject()
  Dim L As Integer, W As Integer
  Dim NewglID As Integer
  
  mDrawable = True
  
  NewglID = -1
  NewglID = picEditor.LoadBitmap(File)

  If NewglID > -1 Then
    picEditor.GetBitmapDimensions NewglID, W, L
  Else
    W = 100
    L = 100
  End If
  ReDim Pts(3)
  Pts(0) = MakePoint(-W / 2 * ZoomX, L / 2 * ZoomY)
  Pts(1) = MakePoint(-Pts(0).X, Pts(0).Y)
  Pts(2) = MakePoint(-Pts(0).X, -Pts(0).Y)
  Pts(3) = MakePoint(Pts(0).X, -Pts(0).Y)

  If glBitmapID > -1 Then picEditor.DeleteBitmap glBitmapID
  glBitmapID = NewglID
End Sub

Private Sub clsObject_VersionChanged(ByVal NewVersion As Integer)
  If NewVersion < Version_FS2K Then
    Visible = False
  End If
End Sub
