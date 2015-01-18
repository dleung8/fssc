VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsATIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ATIS Class
' Provides functions for using ATIS in the program

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
Public Frequency As Single, _
       Range As Byte, _
       Text As String, _
       NW As String, NE As String, _
       SW As String, SE As String

Public mDrawable As Boolean

Public Property Get Caption(Optional ByVal Default As Boolean = False)
  Caption = clsObject_Caption(Default)
End Property

Private Sub Class_Terminate()
'  If Closing Then Error 1
End Sub

Private Function clsObject_Add(ByVal NewX As Single, ByVal NewY As Single) As Boolean
  Dim I As Integer
  If Defaults(OT_ATIS) Is Nothing Or Not Options.Remember Then
    ' Default values:
    Frequency = 127.3
    Range = 100
    Text = Scenery.Header.Name & " %A%E%I%J%K%L %C %G %H%M"
    NW = 36: NE = 36: SW = 36: SE = 36
  Else
    Defaults(OT_ATIS).CopyTo Me
  End If
  X = NewX
  Y = NewY
  mDrawable = True
  clsObject_Add = clsObject_EditProperties()
End Function

Private Property Get clsObject_Caption(Optional ByVal Default As Boolean = False) As String
  If Name = "" Or Default Then
    clsObject_Caption = Lang.GetString(RES_Obj_ATIS) & " " & ObjectIndex
  Else
    clsObject_Caption = Name
  End If
End Property

Private Sub clsObject_Compile(ByVal FileNum As Integer)
  Dim TempStr As String

  ' SCASM, FreeSC incompatibility
  If Scenery.QuoteStrings Then
    TempStr = QuoteString(Text)
  Else
    TempStr = ReplaceParens(Replace(Text, " ", "_"))
  End If
  
  Print #FileNum, "; " & clsObject_Caption
  Print #FileNum, "ATIS( " & CLng(Range * NmToM / 2048) & " " & ReturnPoint(X, Y).LatLon & " " & FloatFormat(Frequency, "000.00") & " " & NE & " " & SE & " " & SW & " " & NW & " " & TempStr & " )" & vbCrLf
End Sub

Private Property Get clsObject_Complexity() As Byte

End Property

Private Property Let clsObject_Complexity(ByVal RHS As Byte)

End Property

Private Sub clsObject_CopyTo(Destination As clsObject)
  Dim Temp As clsATIS
  
  If TypeOf Destination Is clsATIS Then
    Set Temp = Destination
    With Temp
      .Name = Name
      .Locked = Locked
      .X = X
      .Y = Y
      .Frequency = Frequency
      .Range = Range
      .Text = Text
      .NW = NW
      .NE = NE
      .SW = SW
      .SE = SE
    End With
  End If
End Sub

Private Sub clsObject_DrawBottom(Optional ByVal DrawX As Single = -999999#, Optional ByVal DrawY As Single)

End Sub

Private Sub clsObject_DrawTop(Optional ByVal DrawX As Single = -999999#, Optional ByVal DrawY As Single)
  Dim myColor As Long, Buffer As Single
  
  If Not mDrawable Or Not Options.ObjectVisible(OT_ATIS) Then Exit Sub

  glPushMatrix
  If DrawX = -999999 Then
    ' No drag
    If Scenery.Focus(SceneryIndex) Then
      myColor = COLOR_FocusColor
    Else
      myColor = OT_ATIS
    End If
    glTranslatef X, Y, 0
  Else
    ' Drag
    myColor = COLOR_DragColor
    glTranslatef DrawX, DrawY, 0
  End If
  
  Buffer = picEditor.ScaleX(2)
  
  Options.SetGLColor myColor
  glRotatef Scenery.Header.Rotation, 0, 0, 1
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
  If frmATIS.EditData(Me) Then
    clsObject_EditProperties = True
    clsObject_UpdateObject
    If Defaults(OT_ATIS) Is Nothing Then Set Defaults(OT_ATIS) = New clsATIS
    clsObject_CopyTo Defaults(OT_ATIS)
  End If
End Function

Private Sub clsObject_LoadAPTData(FileData As clsAPTData)
  Dim Temp As String
  
  With FileData
    X = .GetData("PosX")
    Y = .GetData("PosY")
    Frequency = ValEx(.GetData("Frequence"))
    Range = Nautical(.GetData("Portee"))
    Temp = .GetData("Pistes")
    NE = ReadNext(Temp, " ")
    SE = ReadNext(Temp, " ")
    SW = ReadNext(Temp, " ")
    NW = ReadNext(Temp, " ")
    Text = Replace(.GetData("Texte"), "_", " ")
    Name = .GetData("Comment")
  End With
  mDrawable = True
End Sub

Private Sub clsObject_LoadBinaryData(ByVal FileNum As Integer, ByVal Version As Integer, ByVal UserData As Long)
  Dim DummyByte1 As Byte, DummyByte2 As Byte, _
    DummyBool1 As Boolean, DummyStr As String, _
    Temp As String
  Select Case Version
    Case Is >= 8
      GetBinaryData FileNum, Name, X, Y, Locked, Frequency, _
        Range, Text, DummyStr, NW, NE, SW, SE, _
        DummyByte1, DummyByte2, DummyBool1
    Case Else
      GetBinaryData FileNum, X, Y, Frequency, _
        Range, Text, DummyStr, NW, NE, SW, SE, _
        DummyByte1, DummyByte2
  End Select
  
  If Version < 3 Then
    Temp = NE
    NE = NW
    NW = SE
    SE = Temp
  End If
  mDrawable = True
End Sub

Private Sub clsObject_LoadData(FileData As clsFileData, ByVal Version As Integer)
  With FileData
    .GetData "Name", Name, ""
    .GetData "Locked", Locked, False
    .GetData "X", X, 0
    .GetData "Y", Y, 0
    .GetData "Frequency", Frequency, 127.3
    .GetData "Range", Range, 100
    .GetData "Text", Text, ""
    .GetData "NW", NW, "36"
    .GetData "NE", NE, "36"
    .GetData "SW", SW, "36"
    .GetData "SE", SE, "36"
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
  clsObject_ObjectType = OT_ATIS
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
    .InitializeForWrite ObjectNames(OT_ATIS), Caption
    .SaveData "Name", Name, ""
    .SaveData "Locked", Locked, False
    .SaveData "X", X, 0
    .SaveData "Y", Y, 0
    .SaveData "Frequency", Frequency, 127.3
    .SaveData "Range", Range, 100
    .SaveData "Text", Text, ""
    .SaveData "NW", NW, "36"
    .SaveData "NE", NE, "36"
    .SaveData "SW", SW, "36"
    .SaveData "SE", SE, "36"
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
  If NewVersion < Version_CFS2 Then
  
  End If
End Sub
