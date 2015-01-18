VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMacro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Macro Class
' Provides functions for using Macros in the program

Option Explicit
Option Compare Text
Implements clsObject

Private Type MacroPointType
  Pts() As PointType
End Type

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
       File As String, _
       Range As Integer, _
       MScale As Single, _
       Altitude As Single, _
       V1 As Long, _
       V2 As Long
       
Public FileExistsCheck As Integer

Private P(9) As String

Public mDrawable As Boolean

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

' Drawing cache
Private AllPts() As MacroPointType

Public Property Get Caption(Optional ByVal Default As Boolean = False)
  Caption = clsObject_Caption(Default)
End Property

' Modify Macro Parameters for compilation
Private Function Modify(ByVal Str As String)
  If Str = "" Then Str = "0"
  If InStr(Str, " ") > 0 Then Str = QuoteString(Str)
  Modify = Str
End Function

Public Property Get Param(Index As Integer) As String
  Param = P(Index)
End Property

Public Property Let Param(Index As Integer, X As String)
  P(Index) = X
End Property

Public Sub ScanTextures(ByRef Files() As String, ByRef FileCount As Integer)
  Dim Meta As MacroMetaType
  
  Dim Contents As String, FileNum As Integer, _
    myStr As String, X As String, I As Integer
  Dim myDir As String
  Dim TempStr As String

  If FileExistsCheck = MUNKNOWN Then
    FileExistsCheck = IIf(FileExists(File), MFILEEXIST, MFILENOTEXIST)
  End If

  Meta = SearchMeta(File)
  myDir = GetDir(File)
  If Meta.Textures <> "" Then
    X = Meta.Textures
    Do Until X = ""
      myStr = ReadNext(X, ",")
      For I = 0 To 9
        If X = "%" & Format$(IIf(I < 3, I + 6, I + 11)) Then
          AddTexFile Files(), FileCount, AddDir(myDir, P(I)), GetDir(File)
          GoTo NextLoop1:
        End If
      Next I
      AddTexFile Files(), FileCount, AddDir(myDir, myStr), GetDir(File)
NextLoop1:
    Loop
  ElseIf FileExistsCheck = MFILEEXIST Then
    FileNum = FreeFile
    Open File For Input As #FileNum
    'Contents = LCase$(Input(LOF(FileNum) - 20, FileNum))
    Do Until EOF(FileNum)
      LineInputEx FileNum, Contents
      If Left$(Contents, 15) = "; BGL_TEXTURE: " Then
        ReadNext Contents, "name="
        AddTexFile Files(), FileCount, AddDir(myDir, Contents), GetDir(File)
      Else
        Do Until Contents = ""
          TempStr = ReadNext(Contents, "bitmap( ")
          
          If Right$(TempStr, 4) = "load" Then
            For I = 1 To 6
              ReadNext Contents, " "
            Next I
            X = UnQuoteString(ReadNext(Contents, " "))
          ElseIf Right$(TempStr, 6) = "repeat" Then
            ' Nothing
          Else
            X = ReadNext(Contents, " ")
          End If
          
          If X <> "" Then
            For I = 0 To 9
              If X = "%" & Format$(IIf(I < 3, I + 6, I + 11)) Then
                AddTexFile Files(), FileCount, AddDir(myDir, P(I)), GetDir(File)
                GoTo NextLoop2:
              End If
            Next I
            AddTexFile Files(), FileCount, AddDir(myDir, X), GetDir(File)
NextLoop2:
          End If
        Loop
      End If
    Loop
    Close #FileNum
  End If
End Sub

Private Sub Class_Terminate()
'  If Closing Then Error 1
End Sub

Private Function clsObject_Add(ByVal NewX As Single, ByVal NewY As Single) As Boolean
  Dim I As Integer
  If Defaults(OT_Macro) Is Nothing Or Not Options.Remember Then
    ' Default values:
    MScale = 1
    Range = 20
    V1 = 0
    V2 = 0
  Else
    Defaults(OT_Macro).CopyTo Me
  End If
  X = NewX
  Y = NewY
  mDrawable = True
  TabValue = 1
  clsObject_Add = clsObject_EditProperties()
End Function

Private Property Get clsObject_Caption(Optional ByVal Default As Boolean = False) As String
  If Name = "" Or Default Then
    clsObject_Caption = Lang.GetString(RES_Obj_Macro) & " " & ObjectIndex
  Else
    clsObject_Caption = Name
  End If
End Property

Private Sub clsObject_Compile(ByVal FileNum As Integer)
  Dim Pos As String
  Dim TempArr(4) As String, I As Integer
  
  Print #FileNum, "; " & clsObject_Caption
  Pos = ReturnPoint(X, Y).LatLon

  If Left$(File, 8) = "LibObj: " Then
    Print #FileNum, "Area( 5 " & Pos & " 12 )"
    If Complexity > 0 Then Print #FileNum, "    IfVarRange( : 346 " & Complexity & " 32767 )"
    Print #FileNum, "    CrashIndirect( :lblObj :lblDraw :lblRotate 0 " & Complexity & " )"
    Print #FileNum, "    PerspectiveCall( :lblPerspective )"
    Print #FileNum, "    ShadowCall( :lblDraw )"
    Print #FileNum, "    Jump( : )"
    Print #FileNum, ":lblPerspective"
    Print #FileNum, "    Perspective"
    Print #FileNum, ":lblDraw"
    Print #FileNum, "    RefPoint( rel :lblReturn " & FloatFormat(MScale, "#0.0########") & " " & Pos & " v1= " & V1 & " v2= " & V2 & IIf(Altitude > 0, " E= " & FloatFormat(Altitude, "####0.0#"), "") & " )"
    Print #FileNum, ":lblRotate"
    Print #FileNum, "    RotatedCall( :lblObj 0 0 " & FloatFormat(Rotation, "##0.0#") & " )"
    Print #FileNum, ":lblReturn"
    Print #FileNum, "    Return"
    Print #FileNum, ":lblObj"
    Print #FileNum, "    CallLibObj( 0 " & Mid$(File, 9) & " )"
    Print #FileNum, "    Return"
    Print #FileNum, "EndA" & vbCrLf
  ElseIf StrComp(Right$(File, 4), ".scm", vbTextCompare) = 0 Then
    If GetShortName(File) = "" Then
      MsgBoxEx frmMain, Lang.ResolveString(RES_ERR_MacroCompileFail, Caption), vbCritical, RES_ERR_MacroCompileFail
    Else
      For I = 0 To 4
        TempArr(I) = P(I)
      Next I
      If P(0) = "" Then TempArr(0) = FloatFormat(MScale, "##0.0#########")
      If P(1) = "" Then TempArr(1) = Complexity
      If P(2) = "" Then TempArr(2) = V1
      If P(3) = "" Then TempArr(3) = Range
      If P(4) = "" Then TempArr(4) = Altitude
      Print #FileNum, "Macro( " & GetShortName(File) & " " & Pos & " " & FloatFormat(Rotation, "##0.0#") & " " & Modify(TempArr(0)) & " " & Modify(TempArr(1)) & " " & Modify(TempArr(2)) & " " & Modify(TempArr(3)) & " " & Modify(TempArr(4)) & " " & Modify(P(5)) & " " & Modify(P(6)) & " " & Modify(P(7)) & " " & Modify(P(8)) & " " & Modify(P(9)) & " )" & vbCrLf
    End If
  Else
    If GetShortName(File) = "" Then
      MsgBoxEx frmMain, Lang.ResolveString(RES_ERR_MacroCompileFail, Caption), vbCritical, RES_ERR_MacroCompileFail
    Else
      Print #FileNum, "Macro( " & GetShortName(File) & " " & Pos & " " & Range & " " & FloatFormat(MScale, "##0.0#########") & " " & FloatFormat(Rotation, "##0.0#") & " " & Modify(P(0)) & " " & Modify(P(1)) & " " & Modify(P(2)) & " " & Modify(P(3)) & " " & V1 & " " & FloatFormat(Altitude, "####0.0#") & " " & Complexity & " v2= " & V2 & " " & Modify(P(4)) & " " & Modify(P(5)) & " " & Modify(P(6)) & " " & Modify(P(7)) & " " & Modify(P(8)) & " " & Modify(P(9)) & " )" & vbCrLf
    End If
  End If
End Sub

Private Property Get clsObject_Complexity() As Byte
  clsObject_Complexity = Complexity
End Property

Private Property Let clsObject_Complexity(ByVal RHS As Byte)
  Complexity = RHS
End Property

Private Sub clsObject_CopyTo(Destination As clsObject)
  Dim Temp As clsMacro, I As Integer
  
  If TypeOf Destination Is clsMacro Then
    Set Temp = Destination
    With Temp
      .Name = Name
      .Locked = Locked
      .X = X
      .Y = Y
      .Complexity = Complexity
      .Rotation = Rotation
      .File = File
      .Range = Range
      .MScale = MScale
      .Altitude = Altitude
      .V1 = V1
      .V2 = V2
      For I = 0 To 9
        .Param(I) = P(I)
      Next I
    End With
  End If
End Sub

Private Sub clsObject_DrawBottom(Optional ByVal DrawX As Single = -999999#, Optional ByVal DrawY As Single)

End Sub

Private Sub clsObject_DrawTop(Optional ByVal DrawX As Single = -999999#, Optional ByVal DrawY As Single)
  Dim myColor As Long, I As Integer
  Dim NewX As Single, NewY As Single
  
  If Not mDrawable Or Not Options.ObjectVisible(OT_Macro) Then Exit Sub

  glPushMatrix
  If DrawX = -999999 Then
    ' No drag
    If Scenery.Focus(SceneryIndex) Then
      myColor = COLOR_FocusColor
    Else
      myColor = OT_Macro
    End If
    glTranslatef X, Y, 0
    NewX = X
    NewY = Y
  Else
    ' Drag
    myColor = COLOR_DragColor
    glTranslatef DrawX, DrawY, 0
    NewX = DrawX
    NewY = DrawY
  End If
  glRotatef Scenery.Header.Rotation - Rotation, 0, 0, 1

  Options.SetGLColor myColor
  
  For I = 0 To UBound(AllPts)
    glDrawPolygon2 AllPts(I).Pts, Options.ObjectColor(myColor), NewX, NewY
  Next I
  
  If DrawX = -999999 Then
    If Options.FocusCircle Then
      Options.SetGLColor myColor
      glDrawCircle X, Y, FocusWidth, Options.ObjectColor(myColor)
    End If
  End If
  glPopMatrix
End Sub

Private Function clsObject_EditProperties() As Boolean
  If frmMacro.EditData(Me) Then
    clsObject_EditProperties = True
    clsObject_UpdateObject
    If Defaults(OT_Macro) Is Nothing Then Set Defaults(OT_Macro) = New clsMacro
    clsObject_CopyTo Defaults(OT_Macro)
  End If
End Function

Private Sub clsObject_LoadAPTData(FileData As clsAPTData)
  Dim I As Integer, Temp As String, Temp2 As String, BaseScale As Single, AirportMacro As Boolean
  
  Dim myDirs As String, myFile As String
  
  If LCase$(FileData.Section) = "fuel" Then
    With FileData
      X = .GetData("PosX")
      Y = .GetData("PosY")
      File = "fuelarea.api"
      P(0) = Meter(.GetData("largeur"))
      Name = .GetData("Comment")
      P(1) = 0
      P(2) = 0
      P(3) = 2
      V1 = 0
      V2 = 0
      Range = 20
      MScale = 1
    End With
  Else
    With FileData
      X = .GetData("PosX")
      Y = .GetData("PosY")
      Temp = .GetData("Type")
      Select Case Left$(Temp, 1)
        Case "&"
          With Options
            If .AirportFolder = "" Then
              .AirportFolder = RegGetKey("MACRODirectory", "", "Software\Airport 2.xx")
              If .AirportFolder = "" Then
                .AirportFolder = RegGetKey("PGMDirectory", "", "Software\Airport 2.xx")
                If .AirportFolder = "" Then
                  MsgBoxEx frmMain, Lang.GetString(RES_ERR_NoAirport), vbCritical, RES_ERR_NoAirport
                Else
                  .AirportFolder = AddDir(.AirportFolder, "API")
                End If
              End If
            End If
            File = GetRealName(AddDir(.AirportFolder, ReadNext(Mid$(Temp, 2), " ") & ".api"))
            BaseScale = ValEx(SearchMeta(File).AirportScale)
            AirportMacro = True
          End With
        Case "@"
          ' File
          File = GetRealName(ReadNext(Mid$(Temp, 2), " "))
        Case Else
          ' Internal Airport Macros
      End Select
      MScale = ValEx(.GetData("Taille"))
      If BaseScale > 0 Then
        MScale = MScale * BaseScale
      End If
      Rotation = ValEx(.GetData("Orientation"))
      Temp = .GetData("Couleur")
      For I = 0 To 3
        Temp = .GetData("STDParams")
        ReadNext Temp, " "
        P(I) = Temp
      Next I
      Range = Nautical(.GetData("Range"))
      Altitude = Meter(.GetData("Altitude"))
      Complexity = .GetComplexityData()
      Name = .GetData("Comment")
      Temp = .GetData("Param")
      For I = 4 To 9
        P(I) = ReadNext(Temp, " ")
      Next I
      If AirportMacro Then
        For I = 0 To 9
          If P(I) = "o" Then P(I) = "1"
          If P(I) = "n" Then P(I) = "0"
        Next I
      End If
      V1 = 0
      V2 = 0
    End With
  End If
  
  If FileExistsCheck = MUNKNOWN Then
    FileExistsCheck = IIf(FileExists(File), MFILEEXIST, MFILENOTEXIST)
  End If
  
  If File <> "" And FileExistsCheck = MFILENOTEXIST Then
    myFile = GetFileTitle(File)
    myDirs = AddDir(App.Path, "Macros") & ";" & Options.MacroPath
    myFile = MultiDir(myFile, myDirs)
    If myFile <> "" Then File = GetRealName(myFile)
  End If
  
  mDrawable = True
  clsObject_UpdateObject
End Sub

Private Sub clsObject_LoadBinaryData(ByVal FileNum As Integer, ByVal Version As Integer, ByVal UserData As Long)
  Dim OldV1 As Integer, OldV2 As Integer, _
    sngDummy1 As Single, sngDummy2 As Single, _
    Alt As Integer, Ext As Integer, Extra As String, _
    I As Integer, bytData As Byte, bytData2 As Byte
  
  Dim myDirs As String, myFile As String
  
  Select Case UserData
    Case 0 ' Macro
      Select Case Version
        Case Is >= 12
          GetBinaryData FileNum, Name, X, Y, Locked, Rotation, _
            File, Range, MScale, Altitude
          GetBinaryData FileNum, P(0), P(1), P(2), P(3), Complexity, _
            P(4), P(5), P(6), P(7), P(8), P(9), V1, V2
        Case Is >= 8
          GetBinaryData FileNum, Name, X, Y, Locked, Rotation, _
            File, Range, MScale, Altitude
          GetBinaryData FileNum, P(0), P(1), P(2), P(3), Complexity, _
            P(4), P(5), P(6), P(7), P(8), P(9), OldV1, OldV2
          V1 = OldV1
          V2 = OldV2
        Case Is >= 3
          GetBinaryData FileNum, X, Y, Rotation, File, Ext, _
            Range, MScale, Altitude, P(0), P(1), P(2), P(3), _
            Complexity, P(4), P(5), P(6), P(7), P(8), P(9)
          ' Seems as if there is a limit to ParamArray arguments
          GetBinaryData FileNum, OldV1, OldV2
          V1 = OldV1
          V2 = OldV2
'        Case Else
'          GetBinaryData FileNum, X, Y, Rotation
'          If Version < 2 Then
'            GetBinaryData FileNum, sngDummy1, sngDummy2, File
'          Else
'            GetBinaryData FileNum, File, Ext
'          End If
'          GetBinaryData FileNum, Range, MScale, Alt, P(0), P(1), P(2), P(3), Complexity
'          Altitude = Alt
'          If Version > 0 Then
'            GetBinaryData FileNum, Extra
'            For I = 4 To 9
'              P(I) = ReadNext(Extra, ",")
'            Next I
'          End If
'          V1 = 6000
'          V2 = 200
      End Select
    Case 1 ' City
      Select Case Version
        Case Is >= 8
          GetBinaryData FileNum, Name, X, Y, Locked, Complexity
        Case Else
          GetBinaryData FileNum, X, Y, Complexity
      End Select
      File = "fscity.api"
      V1 = 0
      V2 = 0
      Range = 20
      MScale = 1
    Case 2 ' Fuel Area
      Select Case Version
        Case Is >= 8
          GetBinaryData FileNum, Name, X, Y, Locked, bytData2, Rotation, bytData
        Case Else
          GetBinaryData FileNum, X, Y, bytData2, Rotation, bytData
      End Select
      P(0) = bytData2
      P(1) = -((bytData And 1) > 0)
      P(2) = -((bytData And 2) > 0)
      P(3) = 2
      File = "fuelarea.api"
      V1 = 0
      V2 = 0
      Range = 20
      MScale = 1
  End Select

  If FileExistsCheck = MUNKNOWN Then
    FileExistsCheck = IIf(FileExists(File), MFILEEXIST, MFILENOTEXIST)
  End If

  If Not Left$(File, 7) = "LibObj:" And FileExistsCheck = MFILENOTEXIST Then
    myFile = GetFileTitle(File)
    myDirs = AddDir(App.Path, "Macros") & ";" & Options.MacroPath
    myFile = MultiDir(myFile, myDirs)
    If myFile <> "" Then File = GetRealName(myFile)
  End If
  
  mDrawable = True
  clsObject_UpdateObject
End Sub

Private Sub clsObject_LoadData(FileData As clsFileData, ByVal Version As Integer)
  Dim myDirs As String, myFile As String
  Dim I As Integer
  
  With FileData
    .GetData "Name", Name, ""
    .GetData "Locked", Locked, False
    .GetData "X", X, 0
    .GetData "Y", Y, 0
    .GetData "Rotation", Rotation, 0
    .GetData "File", File, ""
    .GetData "Range", Range, 20
    .GetData "MScale", MScale, 1
    .GetData "Altitude", Altitude, 0
    .GetData "V1", V1, 0
    .GetData "V2", V2, 0
    .GetData "Complexity", Complexity, 0
    
    For I = 0 To 9
      .GetData "Param" & I, P(I), ""
    Next I
  End With
  
  If FileExistsCheck = MUNKNOWN Then
    FileExistsCheck = IIf(FileExists(File), MFILEEXIST, MFILENOTEXIST)
  End If

  If Not Left$(File, 7) = "LibObj:" And FileExistsCheck = MFILENOTEXIST Then
    myFile = GetFileTitle(File)
    myDirs = AddDir(App.Path, "Macros") & ";" & Options.MacroPath
    If myFile <> "" Then
      myFile = MultiDir(myFile, myDirs)
    End If
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
  clsObject_ObjectType = OT_Macro
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
  Dim I As Integer
  With FileData
    .InitializeForWrite ObjectNames(OT_Macro), Caption
    .SaveData "Name", Name, ""
    .SaveData "Locked", Locked, False
    .SaveData "X", X, 0
    .SaveData "Y", Y, 0
    .SaveData "Rotation", Rotation, 0
    .SaveData "File", File, ""
    .SaveData "Range", Range, 20
    .SaveData "MScale", MScale, 1
    .SaveData "Altitude", Altitude, 0
    .SaveData "V1", V1, 0
    .SaveData "V2", V2, 0
    .SaveData "Complexity", Complexity, 0
    
    For I = 0 To 9
      .SaveData "Param" & I, P(I), ""
    Next I
    
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
  ' Calculate Coordinates before rotate
  ' (relative to polygon center)
  
  Dim Desc As String, Word As String, _
    Length As Single, Width As Single, PtStr As String, _
    I As Integer, Temp As String, FileNum As Integer
  Dim Pos As Integer, Pos2 As Integer, TempScale As Single, UseScale As Single, _
    X As String, Y As String, Z As String
  Dim Meta As MacroMetaType

  Dim Tokenizer As clsTokenizer, _
    Token As String, Token2 As String, _
    QuickPoints As Boolean, Num As Integer, _
    Status As Integer, Skip As Integer, _
    AllPtsCount As Integer, MyPtsCount As Integer
    
  Dim myPts() As PointType, Count As Integer
  
  SetScreenMousePointer vbHourglass
  
  ' Prevent nasty div/0 crash if this is a problem
  If MScale = 0 Then MScale = 1

  If FileExistsCheck = MUNKNOWN Then
    FileExistsCheck = IIf(FileExists(File), MFILEEXIST, MFILENOTEXIST)
  End If

  If FileExistsCheck = MFILENOTEXIST Then
    Width = 50
    Length = 50
    GoTo Done:
  End If

  Meta = SearchMeta(File)
  Desc = Meta.MacroDesc

  Pos = InStr(Desc, "using scale ")
  If Pos = 0 Then Pos = InStr(Desc, "using scale=")
  If Pos > 0 Then
    TempScale = Val(LTrim$(Mid$(Desc, Pos + 11)))
    If TempScale = 0 Then TempScale = 1
  Else
    TempScale = 1
  End If

  ' See if the macro description has something
  Do
    Word = ReadNext(Desc, " ")
    Pos = InStr(Word, "x")
    If Pos > 1 Then
      If InStr(Digits, Mid$(Word, Pos - 1, 1)) > 0 Then
        Width = Val(Word) / TempScale
        Word = Mid$(Word, Pos + 1)
        Length = Val(Word) / TempScale
        If Width = 0 Then Width = 50
        If Length = 0 Then Length = 50
        GoTo Done:
      End If
    End If
  Loop Until Desc = ""
  
  ' See if there is FileNum set of predefined points
  ' This is for the case where Param 6, and 7 hold the
  ' length and width
  PtStr = Meta.Points
  If PtStr Like "param w=%# l=%#" Then
    Width = CSng(P(Val(Mid$(PtStr, 10, 1)) - 6))
    Length = CSng(P(Val(Mid$(PtStr, 15, 1)) - 6))
    If Width = 0 Then Width = 50
    If Length = 0 Then Length = 50
    GoTo Done:
  End If
  
  ' This is the case where there are X set of points
  If PtStr <> "" Then
    ReDim AllPts(0)
    ReDim AllPts(0).Pts(Val(ReadNext(PtStr, ":")))
    For I = 0 To UBound(AllPts(0).Pts) - 1
      AllPts(0).Pts(I) = MakePoint(ReadNext(PtStr, ",") * MScale, ReadNext(PtStr, ";") * MScale)
    Next I
    AllPts(0).Pts(UBound(AllPts(0).Pts)) = AllPts(0).Pts(0)
    AllPtsCount = 1
    GoTo Done:
  End If
  
  ' Airport 2.60's point format
  PtStr = Meta.DesignShape
  If PtStr <> "" Then
    ReDim AllPts(0)
    ReDim AllPts(0).Pts(0)
    I = 0
    Do While PtStr <> ""
      X = ReadNext(PtStr, ",")
      Y = ReadNext(PtStr, ",")
      If InStr(X, " ") > 0 Then
        Temp = ReadNext(X, " ")
        Y = X
        X = Temp
      End If
      AllPts(0).Pts(I) = MakePoint(Val(Trim$(X)) * MScale, Val(Trim$(Y)) * MScale)
      I = I + 1
      ReDim Preserve AllPts(0).Pts(I)
    Loop
    AllPts(0).Pts(I) = AllPts(0).Pts(0)
    AllPtsCount = 1
    GoTo Done:
  End If

  ' We must go through the file itself to find points

  ReDim AllPts(10)
  ReDim myPts(20)
  AllPtsCount = 0
  MyPtsCount = 0
  Count = 0

  PtStr = Meta.QuickPoints
  If PtStr = "True" Or PtStr = "Yes" Or PtStr = "1" Then
    QuickPoints = True
  End If
  
  Set Tokenizer = New clsTokenizer
  Tokenizer.Filename = File
  
  UseScale = MScale
  
  Count = 0
  
  Do While Tokenizer.HasMoreTokens()
    Token = Tokenizer.Token()
    Select Case Token
      Case "RefPoint("
        Token2 = Tokenizer.Token()
        Token2 = Tokenizer.Token()
        Token2 = Tokenizer.Token(True)
        Pos2 = InStr(Token2, "*")
        If Pos2 > 0 Then
          ' Read in multipliers
          Pos = 1
          UseScale = 1
          Do
            Pos2 = InStr(Pos + 1, Token2, "*")
            If Pos2 = 0 Then
              Temp = Trim$(Mid$(Token2, Pos))
            Else
              Temp = Trim$(Mid$(Token2, Pos, Pos2 - Pos))
            End If
            If Temp = "%4" Then
              UseScale = UseScale * MScale
            ElseIf ValEx(Temp) > 0 Then
              UseScale = UseScale * ValEx(Temp)
            End If
            Pos = Pos2 + 1
          Loop Until Pos2 = 0
        Else
          UseScale = ValEx(Token2)
          If Token2 = "%4" Then
            UseScale = MScale
          ElseIf ValEx(Meta.APTMacroScale) > 0 Then
            UseScale = ValEx(Meta.APTMacroScale) * MScale
          ElseIf UseScale = 0 Then
            UseScale = MScale
          End If
        End If
      Case "SetScaleX(", "SuperScale("
        Token2 = Tokenizer.Token()
        Token2 = Tokenizer.Token()
        Token2 = Tokenizer.Token()
        Token2 = Tokenizer.Token()
        If ValEx(Token2) < 32 Then
          UseScale = 2 ^ (ValEx(Token2) - 16)
        End If
      Case "VecPoints(", "VectorPoints("
        Token2 = Tokenizer.Token()
        Select Case Token2
          Case "m"
            Skip = 3
            Token2 = Tokenizer.Token()
          Case "p"
            Skip = 2
            Token2 = Tokenizer.Token()
          Case "xy", "yz", "xz"
            Skip = 0
            Token2 = Tokenizer.Token()
          Case "a"
            Skip = 0
            Token2 = Tokenizer.Token()
          Case Else
            If Not IsNumeric(Token2) Then
              Token2 = Tokenizer.Token()
            Else
              Skip = 3
            End If
        End Select
        If Not QuickPoints Then Count = Val(Token2)
        GoSub ReadPoints:
      Case "Points("
        Skip = 0
        Token2 = Tokenizer.Token()
        If Not QuickPoints Then Count = Val(Token2)
        GoSub ReadPoints:
      Case "Poly(", "ShadedPoly("
        Token2 = Tokenizer.Token()
        Select Case Token2
          Case "m", "abL", "aibL"
            For I = 1 To 4
              Temp = Tokenizer.Token()
            Next I
          Case Else
            If IsNumeric(Token2) Then
              For I = 1 To 3
                Temp = Tokenizer.Token()
              Next I
            End If
        End Select
        Skip = 0
        GoSub ReadPoly:
      Case "TexPoly(", "ShadedTexPoly("
        Token2 = Tokenizer.Token()
        Select Case Token2
          Case "m", "abL", "aibL"
            For I = 1 To 4
              Temp = Tokenizer.Token()
            Next I
          Case Else
            If IsNumeric(Token2) Then
              For I = 1 To 3
                Temp = Tokenizer.Token()
              Next I
            End If
        End Select
        If InStr(Token2, "b") > 0 Or InStr(Token2, "t") > 0 Then
          Skip = 0
        Else
          Skip = 2
        End If
        GoSub ReadPoly:
      Case "VecPoly("
        Token2 = Tokenizer.Token()
        Select Case Token2
          Case "m"
            For I = 1 To 6
              Temp = Tokenizer.Token()
            Next I
          Case Else
        End Select
        Skip = 0
        GoSub ReadPoly:
      Case "TexVecPoly("
        Token2 = Tokenizer.Token()
        Select Case Token2
          Case "m"
            For I = 1 To 6
              Temp = Tokenizer.Token()
            Next I
          Case Else
        End Select
        Skip = 2
        GoSub ReadPoly:
      Case "MoveToPt("
        Token2 = Tokenizer.Token()
      
        If Not QuickPoints Then
          If AllPtsCount > UBound(AllPts) Then
            ReDim Preserve AllPts(AllPtsCount * 2)
          End If
          
          ' Have to incrementally redimension array by 1, since there's no actual "end" command
          ReDim AllPts(AllPtsCount).Pts(0)
          
          Num = Val(Token2)
                  
          With AllPts(AllPtsCount).Pts(0)
            .X = myPts(Num).X
            .Y = myPts(Num).Y
          End With
      
          AllPtsCount = AllPtsCount + 1
          MyPtsCount = 1
        End If
        
        ' Read ending parenthesis
        Token = Tokenizer.Token()
      Case "DrawToPt("
        If Not QuickPoints Then
          Token2 = Tokenizer.Token()
          Do Until Token2 = ")"
            Num = Val(Token2)
              
            If MyPtsCount > UBound(AllPts(AllPtsCount - 1).Pts) Then
              ReDim Preserve AllPts(AllPtsCount - 1).Pts(MyPtsCount * 2)
            End If
              
            With AllPts(AllPtsCount - 1).Pts(MyPtsCount)
              .X = myPts(Num).X
              .Y = myPts(Num).Y
            End With
            
            MyPtsCount = MyPtsCount + 1
            
            Token2 = Tokenizer.Token()
          Loop
          
          ReDim Preserve AllPts(AllPtsCount - 1).Pts(MyPtsCount - 1)
        End If
        
      Case "MoveTo("
        If Not QuickPoints Then
          X = Tokenizer.Token()
          Z = Tokenizer.Token()
          Y = Tokenizer.Token()
          If Not IsNumeric(X) Or Not IsNumeric(Y) Then
            X = "0": Y = "0"
          End If
          
          If AllPtsCount > UBound(AllPts) Then
            ReDim Preserve AllPts(AllPtsCount * 2)
          End If
          
          ' Have to incrementally redimension array by 1, since there's no actual "end" command
          ReDim AllPts(AllPtsCount).Pts(0)
                  
          With AllPts(AllPtsCount).Pts(0)
            .X = Val(X) * UseScale
            .Y = Val(Y) * UseScale
          End With
          
          AllPtsCount = AllPtsCount + 1
          MyPtsCount = 1
          
          ' Read ending parenthesis
          Token = Tokenizer.Token()
        End If
      Case "DrawTo("
        If Not QuickPoints Then
          X = Tokenizer.Token()
          Z = Tokenizer.Token()
          Y = Tokenizer.Token()
          If Not IsNumeric(X) Or Not IsNumeric(Y) Then
            X = "0": Y = "0"
          End If
          
          ReDim Preserve AllPts(AllPtsCount - 1).Pts(MyPtsCount)
          With AllPts(AllPtsCount - 1).Pts(MyPtsCount)
            .X = Val(X) * UseScale
            .Y = Val(Y) * UseScale
          End With
            
          MyPtsCount = MyPtsCount + 1
            
          ' Read ending parenthesis
          Token = Tokenizer.Token()
        End If
    End Select
  Loop
  GoTo Done:

ReadPoints:
  Status = 0
  Token = Tokenizer.Token()
  Do Until Token = ")"
    Select Case Status
      Case 0
        X = Token
        Status = 1
      Case 1
        Z = Token
        Status = 2
      Case 2
        Y = Token
        If Not IsNumeric(X) Or Not IsNumeric(Y) Then
          X = "0": Y = "0"
        End If
        
        If Not QuickPoints Or (QuickPoints And Val(Z) = 0) Then
          If Count > UBound(myPts) Then
            ReDim Preserve myPts(Count * 2)
          End If

          With myPts(Count)
            .X = Val(X) * UseScale
            .Y = Val(Y) * UseScale
          End With

          Count = Count + 1
        End If
        
        For I = 1 To Skip
          Temp = Tokenizer.Token()
        Next I
        
        Status = 0
    End Select
    Token = Tokenizer.Token()
  Loop
  
  Return

ReadPoly:

  If QuickPoints Then Return
  If AllPtsCount > UBound(AllPts) Then
    ReDim Preserve AllPts(AllPtsCount * 2)
  End If
  
  MyPtsCount = 0
  ReDim AllPts(AllPtsCount).Pts(50)
  Token = Tokenizer.Token()
  Do Until Token = ")"
    Num = Val(Token)
      
    If MyPtsCount > UBound(AllPts(AllPtsCount).Pts) Then
      ReDim Preserve AllPts(AllPtsCount).Pts(MyPtsCount * 2)
    End If
      
    With AllPts(AllPtsCount).Pts(MyPtsCount)
      .X = myPts(Num).X
      .Y = myPts(Num).Y
    End With
    
    MyPtsCount = MyPtsCount + 1
    
    For I = 1 To Skip
      Temp = Tokenizer.Token()
    Next I
    
    Token = Tokenizer.Token()
  Loop

  If Tokenizer.M98toBGLSkip Then
    Tokenizer.SetEnd
  End If

  ReDim Preserve AllPts(AllPtsCount).Pts(MyPtsCount)
  
  AllPts(AllPtsCount).Pts(MyPtsCount) = AllPts(AllPtsCount).Pts(0)
  AllPtsCount = AllPtsCount + 1
  Return
  
Done:
  If QuickPoints Then
    ReDim AllPts(0)
    ReDim AllPts(0).Pts(Count)
    CopyMemory AllPts(0).Pts(0), myPts(0), 8 * Count
    AllPts(0).Pts(Count) = AllPts(0).Pts(0)
    AllPtsCount = 1
  End If

  If Length = 0 Or Width = 0 Then
    If AllPtsCount = 0 Then
      Length = 50
      Width = 50
    Else
      If UBound(AllPts(0).Pts) < 1 Then
        ' If there are less than 2 points, then we should
        ' just create an object of length = width = 50
        Length = 50
        Width = 50
      End If
    End If
  End If
  
  If Width <> 0 And Length <> 0 Then
    ReDim AllPts(0)
    ReDim AllPts(0).Pts(6)
    If Width * MScale / 2 < 1 Then Width = 2 / MScale
    If Length * MScale / 2 < 1 Then Length = 2 / MScale
    With AllPts(0)
      .Pts(0) = MakePoint(-Width * MScale / 2, Length * MScale / 2)
      .Pts(1) = MakePoint(-.Pts(0).X, .Pts(0).Y)
      .Pts(2) = MakePoint(-.Pts(0).X, -.Pts(0).Y)
      .Pts(3) = MakePoint(.Pts(0).X, -.Pts(0).Y)
      .Pts(4) = .Pts(0)
      .Pts(5) = MakePoint(0, 0)
      .Pts(6) = .Pts(1)
    End With
    AllPtsCount = 1
  End If
  
  ReDim Preserve AllPts(AllPtsCount - 1)
  For I = 0 To UBound(AllPts)
    MultiRotate AllPts(I).Pts, Rotation - Scenery.Header.Rotation
  Next I

  SetScreenMousePointer vbDefault

  mDrawable = True
End Sub

Private Sub clsObject_VersionChanged(ByVal NewVersion As Integer)

End Sub
