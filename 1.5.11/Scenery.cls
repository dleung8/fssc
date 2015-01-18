VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScenery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ToDo
' Background
' linked polygons
' BGL import

' White roads
' screwy images
' bmp files

' Object Identifiers
Public Enum ObjectTypeEnum
  OT_Header = 0
  OT_Runway
  OT_Polygon
  OT_Taxiway
  OT_Road
  OT_River
  OT_Line
  OT_TaxiwayLine
  OT_Building
  OT_Macro
  OT_ATIS
  OT_VOR
  OT_NDB
  OT_TowerLocation
  OT_MenuEntry
  OT_Background
  OT_FlatArea
  OT_SurfaceArea
  OT_Exclusion
  OT_Code
  
  OT_Point
End Enum

Public Event Redraw()
Public Event ClearValues()
Public Event FocusChanged(NewFoci() As Integer)
Public Event TitleBarChange()

Public Changed As Boolean
Public SingleFocus As Integer

' The objects themselves
Private ScObjects() As clsObject
Private NumObjects As Integer

' If True, need to refresh the AFD information
Public AFDRefresh As Boolean

' Provides a number for polygon-polygon and line-macro
' link independent of position (useful when users edit
' the ASCII save file
Public LinkCounter As Integer

' The count number for accounting purposes
Private ObjectCounts(OT_Point) As Integer

Private mFile As String
Private mFocus() As Boolean

Public ExportPath As String

' Some flags from compiling
Public QuoteStrings As Boolean
Public FileBase As String
Public TitleBase As String
Public ExcludeCount As Integer
Public HeaderStr As String
Public AIHeaderStr As String
'Public PolygonLayer As Byte
'Public PolygonBuffer As String
'Public RunwayBuffer As String
Public AFDBuffer As String
Public RunwayLocBuffer As String
Public RunwayNAVBuffer As String
Public NameListBuffer As String
Public NameListABuffer As String
Public A16NBuffer As String

Public TowerIndex As Integer

Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

' Adds an FS object
Public Sub Add(ByVal ObjectType As ObjectTypeEnum, ByVal NewX As Single, ByVal NewY As Single)
  Dim NewObject As clsObject, _
      PointObject As clsPoint, _
      ShapeObject As clsShape, _
      Index As Integer

  Set NewObject = CreateNewInstance(ObjectType)

  If TypeOf NewObject Is clsShape Then
    Set PointObject = New clsPoint
    Set ShapeObject = NewObject
    Set PointObject.Parent = ShapeObject
    PointObject.Parent.ObjectIndex = ObjectCounts(ObjectType) + 1
    Set NewObject = PointObject
  Else
    NewObject.ObjectIndex = ObjectCounts(ObjectType) + 1
  End If

  Index = Count
  If NewObject.Add(NewX, NewY) Then
    NewObject.SceneryIndex = Index + 1
    ObjectCounts(ObjectType) = ObjectCounts(ObjectType) + 1
    
    If NewObject.ObjectType <> OT_Point Then
      Count = Index + 1
      Set ScObjects(Index + 1) = NewObject
      
      frmMain.UpdateObjectList
      ' forces a draw
      SetSingleFocus Index + 1
    ElseIf PointObject.Parent.NumPoints = 1 Then ' (2 points)
      frmMain.UpdateObjectList
      SetSingleFocus Index + 1
      ScObjects(Index + 2).DrawBottom
      ScObjects(Index + 2).DrawTop
    ElseIf PointObject.Parent.NumPoints = 2 Then ' (3 points)
      frmMain.UpdateObjectList
      SetSingleFocus Index + 1
      ScObjects(Index + 2).DrawBottom
      ScObjects(Index + 2).DrawTop
      ScObjects(Index + 3).DrawBottom
      ScObjects(Index + 3).DrawTop
    End If
    If TypeOf NewObject Is clsTower Then TowerIndex = Index + 1
    Changed = True
  End If
  Set NewObject = Nothing
End Sub

' Adds the specified point into the scenery objects list
Public Sub AddPoint(Point As clsPoint)
  Dim Index As Integer
  Index = Count
  Count = Index + 1
  Set ScObjects(Index + 1) = Point
  Point.SceneryIndex = Index + 1
End Sub

Public Sub BringToFront()
  Dim I As Integer
  Dim MyObject As clsObject, ShapeObject As clsShape, _
    PointObject As clsPoint, PointObject2 As clsPoint
  Dim Count As Integer, Num As Integer, UpTo As Integer
  
  Dim Saved() As clsObject
  
  UpTo = SingleFocus
  Set MyObject = Scenery(UpTo)
  
  ClearFocus True
  
  If MyObject.ObjectType = OT_Point Then
    Set PointObject = MyObject
    Set ShapeObject = PointObject.Parent
    With ShapeObject
      ReDim Saved(.NumPoints)
      For I = 0 To .NumPoints
        Set Saved(I) = .Point(I)
        Set ScObjects(.Point(I).SceneryIndex) = Nothing
      Next I
    End With
    Set MyObject = PointObject.Parent.Point(0)
    
    For I = 1 To NumObjects
      If ScObjects(I) Is Nothing Then
        Count = Count + 1
      Else
        Set ScObjects(I - Count) = ScObjects(I)
        With ScObjects(I - Count)
          .SceneryIndex = I - Count
          
          If .ObjectType = OT_Point Then
            If .ObjectIndex = 1 Then
              Set PointObject2 = ScObjects(I - Count)
              If PointObject.Parent.ShapeType = PointObject2.Parent.ShapeType Then
                PointObject2.Parent.ObjectIndex = PointObject2.Parent.ObjectIndex - 1
              End If
            End If
          End If
        End With
      End If
    Next I
    For I = 0 To UBound(Saved)
      Num = NumObjects - Count + I + 1
      Set ScObjects(Num) = Saved(I)
      ScObjects(Num).SceneryIndex = Num
    Next I
    PointObject.Parent.ObjectIndex = ObjectCounts(PointObject.Parent.ShapeType)
  Else
    For I = UpTo To NumObjects - 1
      Set ScObjects(I) = ScObjects(I + 1)
      ScObjects(I).SceneryIndex = I
      If ScObjects(I).ObjectType <> OT_Point Then
        If ScObjects(I).ObjectType = MyObject.ObjectType Then
          ScObjects(I).ObjectIndex = ScObjects(I).ObjectIndex - 1
        End If
      End If
    Next I
    MyObject.SceneryIndex = NumObjects
    MyObject.ObjectIndex = ObjectCounts(MyObject.ObjectType)
    Set ScObjects(NumObjects) = MyObject
  End If
  
  frmMain.UpdateObjectList
  Changed = True
  
  Scenery.Draw
End Sub

' Resets the Scenery object
Public Sub Clear()
  Class_Terminate
  
  ReDim ScObjects(10)
  ReDim mFocus(10)
  Erase ObjectCounts
  SingleFocus = 0
  Set ScObjects(0) = New clsHeader
  NumObjects = 0
  TowerIndex = 0
  
  AFDRefresh = False
  Changed = False
  mFile = Lang.GetString(RES_UntitledFile)
  ExportPath = ""
  
  RaiseEvent ClearValues
  RaiseEvent TitleBarChange
End Sub

' Sets all the Foci to False
Public Sub ClearFocus(Optional ByVal RaiseRedrawEvent As Boolean = True)
  Dim I As Integer
  picEditor.StartDraw
  For I = 0 To Count
    If mFocus(I) Then
      mFocus(I) = False
      ScObjects(I).DrawBottom
      ScObjects(I).DrawTop
    End If
  Next I
  SingleFocus = 0
  RaiseSingleFocusChangeEvent 0
  If RaiseRedrawEvent Then RaiseEvent Redraw
End Sub

Public Function Compile(ByVal DoAFD As Boolean) As Boolean
  Dim I As Integer, FileNum As Integer, _
    SourceFile As String, _
    BGLFile As String, _
    AFDSourceFile As String, A16NSourceFile As String, _
    AFDBGLFile As String, A16NBGLFile As String, _
    rc As Boolean, FileTitle As String, OldDir As String
  Dim TowerObject As clsTower
    
  OldDir = CurDir$
  
  Dim Objs() As Integer
  
  ' Progress bar
  With frmMain
    If .StatusBar.Visible And Count > 0 Then
      .barProgress.Value = 0
      .lblStatus.Caption = Lang.GetString(RES_Main_Compiling)
      .barProgress.Max = Count
      .StatusBar.Visible = False
      .picProgress.Visible = True
      .picProgress.Refresh
    End If
  End With

  SetScreenMousePointer vbHourglass

  FileTitle = GetFileTitle(mFile)
  FileBase = AddDir(ExportPath, "Scenery\" & Left$(FileTitle, InStrRev(FileTitle, ".") - 1))
  TitleBase = GetFileTitle(FileBase)
  
  ExcludeCount = 0
  AFDBuffer = ""
  RunwayLocBuffer = ""
  NameListBuffer = ""
  NameListABuffer = ""
  A16NBuffer = ""
  RunwayNAVBuffer = "Container( APT " & Scenery.Count + 1 & " )" & vbCrLf & _
                    "    APLocation( " & Header.Center.LatLon & " " & FloatFormat(Header.Altitude, "####0.0#") & " )" & vbCrLf & _
                    "    ICAO_ID( " & QuoteString(Header.ICAOID) & " )" & vbCrLf
  If TowerIndex > 0 Then
    Set TowerObject = Scenery(TowerIndex)
    With TowerObject
      RunwayNAVBuffer = RunwayNAVBuffer & "    Tower( " & ReturnPoint(.X, .Y).LatLon & " " & FloatFormat(.Height + Header.Altitude, "##0.0#") & " 0 )" & vbCrLf
      For I = 0 To 11
        If .COMFrequency(I) > 0 Then _
          RunwayNAVBuffer = RunwayNAVBuffer & "    Frequency( " & FloatFormat(.COMFrequency(I), "000.00") & " " & Choose(I + 1, "ATIS", "MULTICOM", "UNICOM", "CTAF", "GNDCON", "CT", "CDEL", "APCON", "DEPCON", "CENTER", "FSS", "AWOS") & " )" & vbCrLf
      Next I
    End With
    Set TowerObject = Nothing
  End If
  RunwayNAVBuffer = RunwayNAVBuffer & vbCrLf

  SourceFile = FileBase & ".sca"
  BGLFile = FileBase & ".bgl"
  
  AFDSourceFile = FileBase & "_AFD.sca"
  AFDBGLFile = FileBase & "_AFD.bgl"
  
  A16NSourceFile = FileBase & "_A16N.sca"
  A16NBGLFile = FileBase & "_A16N.bgl"
  
  ChangeDir GetDir(FileBase)
  
  QuoteStrings = (InStr(1, Options.Compiler, "SCASM.EXE", vbTextCompare) = 0)
  
  ReDim Objs(NumObjects)
  For I = 0 To NumObjects
    Objs(I) = I
  Next I
  QuickSort Objs, 0, NumObjects

  FileNum = FreeFile
  Open SourceFile For Output As #FileNum
  Print #FileNum, FileHeader
  For I = 0 To NumObjects
    ScObjects(Objs(I)).Compile FileNum
  Next I
  Close #FileNum
  
  If DoAFD Then
    Open AFDSourceFile For Output As #FileNum
    Print #FileNum, FileHeader
    
    If Options.FSVersion >= Version_FS2K2 Then
      Print #FileNum, "Set( FSVERS 0x800 )"
      Print #FileNum, AIHeaderStr
    End If
    
    Print #FileNum, HeaderStr

    If RunwayLocBuffer <> "" Then
      Print #FileNum, RunwayNAVBuffer
      Print #FileNum, RunwayLocBuffer
      Print #FileNum, "EndC" & vbCrLf
    End If
    
    Print #FileNum, AFDBuffer
    
    With Scenery.Header
      Print #FileNum, "NameList( " & Right$(.LangCode, 3) & " )"
      Print #FileNum, "  NameEntry( REGION 0 " & QuoteString(.Region) & " )"
      Print #FileNum, "  NameEntry( COUNTRY 0 " & QuoteString(.Country) & " )"
      If .State <> "" Then Print #FileNum, "  NameEntry( STATE 0 " & QuoteString(.State) & " )"
      Print #FileNum, NameListBuffer;
      If NameListABuffer <> "" Then
        Print #FileNum, "  NameEntry( CITY 0 " & QuoteString(.City) & " )"
        Print #FileNum, NameListABuffer;
      End If
      Print #FileNum, "EndNL"
    End With

    Close #FileNum
  End If
  
  If A16NBuffer <> "" Then
    Open A16NSourceFile For Output As #FileNum
    Print #FileNum, FileHeader
    Print #FileNum, HeaderStr
    Print #FileNum, A16NBuffer
    Close #FileNum
  End If

  If Not FileExists(Options.Compiler) Then
    Options.Compiler = AddDir(App.Path, "SCASM\SCASM.EXE")
  End If
  
  If FileExists(Options.Compiler) Then
    rc = RunSCASM(SourceFile, BGLFile)
    If rc Then
      If Options.FSVersion >= Version_FS2K And Options.AutoCompress And FileExists(Options.CompressName) Then
        RunDosFile Options.CompressName, "-f " & QuoteString(BGLFile), GetDir(BGLFile)
      End If

      If A16NBuffer <> "" Then
        rc = RunSCASM(A16NSourceFile, A16NBGLFile)
      End If
      If DoAFD Then
        With Scenery.Header
          If .Region = "" Or .Country = "" Or .City = "" Or .ICAOID = "" Then
            MsgBoxEx frmMain, Lang.GetString(RES_ERR_AFDData), vbCritical, RES_ERR_AFDData
          Else
            rc = RunSCASM(AFDSourceFile, AFDBGLFile)
            If rc Then AFDRefresh = False
          End If
        End With
      End If
    End If
    If rc And Options.EditConfig Then
      WriteSceneryCFGFile AddDir(Options.FSPath, "scenery.cfg")
    End If
    
    Compile = rc
  Else
    SetScreenMousePointer vbDefault
    MsgBoxEx frmMain, Lang.GetString(RES_ERR_CompilerPath), vbCritical, RES_ERR_CompilerPath
  End If

  ' Progress bar
  With frmMain
    If .picProgress.Visible And Count > 0 Then
      .barProgress.Value = Count
      .picProgress.Visible = False
      .StatusBar.Visible = True
    End If
  End With

  ChangeDir OldDir
  
  SetScreenMousePointer vbDefault
End Function

' Returns # of objects
Public Property Get Count() As Long
  Count = NumObjects
End Property

' Resizes the Scenery Object
Public Property Let Count(ByVal vData As Long)
  Dim I As Integer
  If vData > NumObjects Then
    If vData > UBound(ScObjects) Then
      ReDim Preserve ScObjects(vData * 2)
      ReDim Preserve mFocus(vData * 2)
    End If
  ElseIf vData < NumObjects Then
    For I = vData + 1 To NumObjects
      Set ScObjects(I) = Nothing
    Next I
    
    If vData * 4 < UBound(ScObjects) Then
      ReDim Preserve ScObjects(vData * 2)
      ReDim Preserve mFocus(vData * 2)
    End If
  End If
  NumObjects = vData
End Property

' Creates a new FS object in memory
Public Function CreateNewInstance(ByVal ObjType As ObjectTypeEnum) As clsObject
  Dim ShapeObject As clsShape
  Select Case ObjType
    Case OT_Header:      Set CreateNewInstance = New clsHeader
    Case OT_Runway:      Set CreateNewInstance = New clsRunway
    Case OT_Polygon, OT_Taxiway, OT_Road, _
         OT_River, OT_Line, OT_FlatArea, OT_TaxiwayLine:
                         Set ShapeObject = New clsShape
                         ShapeObject.ShapeType = ObjType
                         Set CreateNewInstance = ShapeObject
                         Set ShapeObject = Nothing
    Case OT_Building:    Set CreateNewInstance = New clsBuilding
    Case OT_Macro:       Set CreateNewInstance = New clsMacro
    Case OT_ATIS:        Set CreateNewInstance = New clsATIS
    Case OT_VOR:         Set CreateNewInstance = New clsVOR
    Case OT_NDB:         Set CreateNewInstance = New clsNDB
    Case OT_TowerLocation: Set CreateNewInstance = New clsTower
    Case OT_MenuEntry:   Set CreateNewInstance = New clsMenuEntry
    Case OT_Background:  Set CreateNewInstance = New clsBackground
    Case OT_SurfaceArea: Set CreateNewInstance = New clsSurfaceArea
    Case OT_Exclusion:   Set CreateNewInstance = New clsExclusion
    Case OT_Code:        Set CreateNewInstance = New clsCode
    Case OT_Point:       Set CreateNewInstance = New clsPoint
    Case Else
      Set CreateNewInstance = Nothing
      Exit Function
  End Select
End Function

Public Sub DoCopy(mClipboard() As clsObject)
  Dim I As Integer, J As Integer, _
      Sum As Integer, Num As Integer
  Dim PointObject As clsPoint, PointObject2 As clsPoint
  
  Dim ShapeMap() As clsShape, ShapeNumMap() As Integer
  Dim ShapeMapCount As Integer
  
  ReDim ShapeMap(10)
  ReDim ShapeNumMap(Count)
  
  For I = 1 To NumObjects
    If mFocus(I) Then
      If ScObjects(I).ObjectType = OT_Point Then
        Set PointObject = ScObjects(I)
        Focus(PointObject.Parent.Point(0).SceneryIndex) = True
        Set PointObject = Nothing
      End If
    End If
  Next I
  
  For I = 1 To NumObjects
    If mFocus(I) Then
      If ScObjects(I).ObjectType = OT_Point Then
        If ScObjects(I).ObjectIndex = 1 Then
          Set PointObject = ScObjects(I)
          With PointObject
            If ShapeMapCount > UBound(ShapeMap) Then ReDim Preserve ShapeMap(ShapeMapCount * 2)
            Set ShapeMap(ShapeMapCount) = CreateNewInstance(.Parent.ShapeType)
            ShapeNumMap(.SceneryIndex) = ShapeMapCount
            
            .Parent.CopyTo ShapeMap(ShapeMapCount)
            For J = 0 To .Parent.NumPoints
              Focus(.Parent.Point(J).SceneryIndex) = True
            Next J
            Sum = Sum + .Parent.NumPoints + 1
          
            ShapeMapCount = ShapeMapCount + 1
          End With
          Set PointObject = Nothing
        End If
      Else
        Sum = Sum + 1
      End If
    End If
  Next I

  ReDim mClipboard(Sum)
  J = 0
  For I = 1 To NumObjects
    If mFocus(I) Then
      J = J + 1
      Set mClipboard(J) = CreateNewInstance(ScObjects(I).ObjectType)
      ScObjects(I).CopyTo mClipboard(J)
      If mClipboard(J).ObjectType = OT_Point Then
        Set PointObject = ScObjects(I)
        Set PointObject2 = mClipboard(J)
        Num = ShapeNumMap(PointObject.Parent.Point(0).SceneryIndex)
        Set PointObject2.Parent = ShapeMap(Num)
        PointObject2.Parent.IntTag = Num
        Set PointObject = Nothing
        Set PointObject2 = Nothing
      End If
    End If
  Next I
  
  For I = 0 To UBound(ShapeMap)
    Set ShapeMap(I) = Nothing
  Next I
  
  RefreshFocus True
End Sub

Public Sub DoDelete(Optional ByVal FromCopy As Boolean)
  Dim I As Integer, J As Integer
  Dim PointObject As clsPoint
  Dim ShapeChanges() As clsShape, ShapeCount As Integer
  Dim DeleteCounts(OT_Point) As Integer
  Dim SelCount As Integer, NumLeft As Integer
  
  Dim BackCount As Integer
  Dim RunwayObject As clsRunway, VORObject As clsVOR, NDBObject As clsNDB
    
  Dim Num As ObjectTypeEnum
  
  ReDim ShapeChanges(10)
  
  For I = 1 To NumObjects
    If ScObjects(I).ObjectType = OT_Point Then
      If ScObjects(I).ObjectIndex = 1 Then
        Set PointObject = ScObjects(I)
        SelCount = 0
        With PointObject.Parent
          If FromCopy Then
            If mFocus(I) Then
              DeleteCounts(.ShapeType) = DeleteCounts(.ShapeType) + 1
              For J = 0 To .NumPoints
                mFocus(.Point(J).SceneryIndex) = True
                Set .Point(J).Parent = Nothing
              Next J
            End If
          Else
            For J = 0 To .NumPoints
              If mFocus(.Point(J).SceneryIndex) Then
                SelCount = SelCount + 1
              End If
            Next J
            NumLeft = .NumPoints - SelCount + 1
            
            If (NumLeft < 2) Or ((NumLeft < 3) And (.ShapeType = OT_Polygon Or .ShapeType = OT_FlatArea)) Then
              For J = 0 To .NumPoints
                mFocus(.Point(J).SceneryIndex) = True
                Set .Point(J).Parent = Nothing
              Next J
              DeleteCounts(.ShapeType) = DeleteCounts(.ShapeType) + 1
            Else
              ' Between 1 and number of points - 1
              If Between(SelCount, 1, .NumPoints) Then
                If ShapeCount > UBound(ShapeChanges) Then ReDim Preserve ShapeChanges(ShapeCount * 2)
                Set ShapeChanges(ShapeCount) = PointObject.Parent
                ShapeCount = ShapeCount + 1
              End If
              
              .ObjectIndex = .ObjectIndex - DeleteCounts(.ShapeType)
            End If
          End If
        End With
      End If
    Else
      Num = ScObjects(I).ObjectType
      If mFocus(I) Then
        If TypeOf ScObjects(I) Is clsRunway Then
          Set RunwayObject = ScObjects(I)
          Set RunwayObject.Near = Nothing
          Set RunwayObject.Far = Nothing
          If RunwayObject.AFDEntry Then
            Scenery.AFDRefresh = True
          End If
        ElseIf TypeOf ScObjects(I) Is clsVOR Then
          Set VORObject = ScObjects(I)
          If ((VORObject.Flags And 128) > 0) Then
            Scenery.AFDRefresh = True
          End If
          Set VORObject = Nothing
        ElseIf TypeOf ScObjects(I) Is clsNDB Then
          Set NDBObject = ScObjects(I)
          If NDBObject.AFDEntry Then
            Scenery.AFDRefresh = True
          End If
          Set NDBObject = Nothing
        End If
        DeleteCounts(Num) = DeleteCounts(Num) + 1
      Else
        ScObjects(I).ObjectIndex = ScObjects(I).ObjectIndex - DeleteCounts(Num)
      End If
      If Num = OT_TowerLocation Then TowerIndex = 0
    End If
  Next I
  
  For I = 0 To ShapeCount - 1
    ShapeChanges(I).DeletePoints
  Next I
  
  For I = 0 To OT_Point
    ObjectCounts(I) = ObjectCounts(I) - DeleteCounts(I)
  Next I
  
  For I = 0 To UBound(ShapeChanges)
    Set ShapeChanges(I) = Nothing
  Next I
    
  For I = 1 To NumObjects
    If mFocus(I) Then
      Set ScObjects(I) = Nothing
      BackCount = BackCount + 1
    Else
      If BackCount > 0 Then
        ScObjects(I).SceneryIndex = I - BackCount
        Set ScObjects(I - BackCount) = ScObjects(I)
      End If
    End If
  Next I
  
  For I = 0 To Count
    mFocus(I) = False
  Next I
  SingleFocus = 0
  
  If ObjectCounts(OT_FlatArea) = 0 Then
    Dim FileTitle As String, FileBase As String, File As String
    FileTitle = GetFileTitle(Scenery.File)
    FileBase = Left$(FileTitle, InStrRev(FileTitle, ".") - 1)
    FileTitle = FileBase & "_A16N.bgl"
    File = AddDir(ExportPath, "Scenery\" & FileTitle)
    If FileExists(File) Then
      If MsgBoxEx(frmMain, Lang.ResolveString(RES_ERR_FlatArea, FileTitle), vbQuestion Or vbYesNo, RES_ERR_FlatArea) = vbYes Then
        On Error Resume Next
        Kill File
        On Error GoTo 0
      End If
    End If
  End If

  Count = Count - BackCount

  frmMain.UpdateObjectList
  Changed = True
  
  Scenery.Draw
  RaiseSingleFocusChangeEvent 0
End Sub

Private Sub DoOldVersionConversion()
  Dim I As Integer
  Const ConversionValue = 1852 / 1856
  If frmConvert.DoDialog = 1 Then
    For I = 1 To Count
      With Scenery(I)
        .PositionX = .PositionX * ConversionValue
        .PositionY = .PositionY * ConversionValue
      End With
    Next I
  End If
End Sub

Public Sub DoPaste(mClipboard() As clsObject)
  Dim I As Integer
  Dim PointObject As clsPoint, PointObject2 As clsPoint
  Dim Start As Integer
  Dim Num As Integer
  
  Dim ShapeMap() As clsShape
  
  ClearFocus False
 
  ReDim ShapeMap(10)

  Start = NumObjects
  Count = Start + UBound(mClipboard)
  For I = 1 To UBound(mClipboard)
    If mClipboard(I).ObjectType = OT_TowerLocation And TowerIndex > 0 Then
      Start = Start - 1
      Count = Count - 1
      MsgBoxEx frmMain, Lang.GetString(RES_ERR_Tower), vbInformation, RES_ERR_Tower
    Else
      If mClipboard(I).ObjectType = OT_TowerLocation Then TowerIndex = I

      Set ScObjects(Start + I) = CreateNewInstance(mClipboard(I).ObjectType)
      With ScObjects(Start + I)
        mFocus(Start + I) = True
        .SceneryIndex = Start + I
        Num = mClipboard(I).ObjectType
        ObjectCounts(Num) = ObjectCounts(Num) + 1
        .ObjectIndex = ObjectCounts(Num)
        mClipboard(I).CopyTo ScObjects(Start + I)
        If .ObjectType = OT_Point Then
          Set PointObject = ScObjects(Start + I)
          If .ObjectIndex = 1 Then
            Set PointObject2 = mClipboard(I)
            Set PointObject.Parent = CreateNewInstance(PointObject2.Parent.ShapeType)
            PointObject2.Parent.CopyTo PointObject.Parent
            
            Num = PointObject.Parent.ShapeType
            ObjectCounts(Num) = ObjectCounts(Num) + 1
            PointObject.Parent.ObjectIndex = ObjectCounts(Num)
            
            If PointObject2.Parent.IntTag > UBound(ShapeMap) Then
              ReDim Preserve ShapeMap(PointObject2.Parent.IntTag * 2)
            End If
            Set ShapeMap(PointObject2.Parent.IntTag) = PointObject.Parent
            Set PointObject2 = Nothing
          Else
            Set PointObject2 = mClipboard(I)
            Set PointObject.Parent = ShapeMap(PointObject2.Parent.IntTag)
          End If
          PointObject.Parent.AddPoint PointObject, PointObject.ObjectIndex
          Set PointObject = Nothing
        End If
      End With
    End If
  Next I

  For I = 0 To UBound(ShapeMap)
    Set ShapeMap(I) = Nothing
  Next I
  
  For I = Start + 1 To NumObjects
    ScObjects(I).UpdateObject
  Next I
  
  frmMain.UpdateObjectList
  Changed = True
  
  Draw
  RefreshFocus True
End Sub

' Calls the Draw function for each object
Public Sub Draw()
  Dim I As Integer
  On Error Resume Next
  
  SetScreenMousePointer vbHourglass

  FocusWidth = picEditor.ScaleX(AbsFocusWidth)

  picEditor.StartDraw
  
  Options.SetGLBackColor
  glCls

  For I = 0 To NumObjects
    ScObjects(I).DrawBottom
  Next I
  For I = 0 To NumObjects
    ScObjects(I).DrawTop
  Next I

  ' Transfer the buffer to the visible picturebox
  RaiseEvent Redraw

  SetScreenMousePointer vbDefault
End Sub

' Edit the properties of an object
Public Sub EditProperties(ByVal Index As Integer)
  If ScObjects(Index).EditProperties Then
    If Index = 0 Then RaiseEvent TitleBarChange
    frmMain.UpdateObjectList Index
    Changed = True
    Draw
  End If
End Sub

' Gets the file name
Public Property Get File() As String
  File = mFile
End Property

' Sets the file name and changes the menu caption
Public Property Let File(ByVal vData As String)
  mFile = vData
  RaiseEvent TitleBarChange
End Property

Public Function FileHeader() As String
  FileHeader = "; " & String(60, "*") & vbCrLf & _
               "; Scenery compiled by FS Scenery Creator " & App.Major & "." & App.Minor & "." & App.Revision & " by Derek Leung" & vbCrLf & _
               "; Scenery compiled on " & Now & vbCrLf & _
               "; Scenery created by " & Header.Author & vbCrLf & _
               "; " & String(60, "*") & vbCrLf
End Function

' Gets the focus status
Public Property Get Focus(ByVal Index As Integer) As Boolean
  Focus = mFocus(Index)
End Property

' Sets a focus and draws it,
' but does not raise a redraw event
Public Property Let Focus(ByVal Index As Integer, ByVal vData As Boolean)
  If vData <> mFocus(Index) Then
    mFocus(Index) = vData
    picEditor.StartDraw
    ScObjects(Index).DrawBottom
    ScObjects(Index).DrawTop
  End If
  If vData Then SingleFocus = Index
End Property

' No frills version of Let Focus
Public Property Let FocusSimple(ByVal Index As Integer, ByVal vData As Boolean)
  mFocus(Index) = vData
End Property

' Gets the header in clsHeader form
Public Function Header() As clsHeader
  Set Header = ScObjects(0)
End Function

' Checks to see if the cursor is inside the focus area
' denoted by the circle, and returns the indices in
' an array
Public Sub InsideFocusArea(ByVal X As Single, ByVal Y As Single, ByVal Topmost As Boolean)
  Dim I As Integer, Cnt As Integer, Radius2 As Single, _
    Low As Integer, High As Integer, myStep As Integer
  Radius2 = FocusWidth * FocusWidth

  ReDim Indices(10) As Integer
  
  Low = IIf(Topmost, Count, 1)
  High = IIf(Topmost, 1, Count)
  myStep = IIf(Topmost, -1, 1)
  
  For I = Low To High Step myStep
    With ScObjects(I)
      If Distance2(X, Y, .PositionX, .PositionY) < Radius2 Then
        If Options.ObjectVisible(ScObjects(I).ObjectType) Then
          Cnt = Cnt + 1
          If Cnt > UBound(Indices) Then ReDim Preserve Indices(Cnt * 2)
          Indices(Cnt) = I
          Focus(I) = True
        Else
          Focus(I) = False
        End If
      Else
        Focus(I) = False
      End If
    End With
  Next I
  ReDim Preserve Indices(Cnt)
  
  RaiseEvent FocusChanged(Indices)
  RaiseEvent Redraw
End Sub

' Gets the Scenery object at Index
' Default Property
Public Property Get Item(ByVal Index As Integer) As clsObject
Attribute Item.VB_UserMemId = 0
  Set Item = ScObjects(Index)
End Property

' Sets the Scenery object at Index
' Default Property
Public Property Set Item(ByVal Index As Integer, NewObject As clsObject)
  Set ScObjects(Index) = NewObject
End Property

' Loads file data (Airport format)
Private Sub LoadAPTFile(ByVal Filename As String)
  Dim FileNum As Integer, FileData As clsAPTData

  Dim I As Integer, J As Integer, Counter As Integer

  Dim NewObject As clsObject, ShapeObject As clsShape, _
    SurfaceAreaObject As clsSurfaceArea, RunwayObject As clsRunway
    
  Dim strTemp As String

'  Dim LinkMap() As clsObject
'
'  Dim ShapeObject As clsShape, NewObject As clsObject, _
'    PointObject As clsPoint
'
'  Dim BuildingObject As clsBuilding, MacroObject As clsMacro
'  Dim LinkFromNum As Integer

  FileNum = FreeFile

  Open Filename For Input As #FileNum
  Set FileData = New clsAPTData
  FileData.FileNum = FileNum
  
  ' Progress bar
  With frmMain
    If .StatusBar.Visible Then
      .barProgress.Value = 0
      .lblStatus.Caption = Lang.GetString(RES_Main_Loading)
      .barProgress.Max = LOF(FileNum)
      .StatusBar.Visible = False
      .picProgress.Visible = True
      .picProgress.Refresh
    End If
  End With
  
  ' Header
  Set ScObjects(0) = CreateNewInstance(OT_Header)
  ScObjects(0).ObjectIndex = 0
  ObjectCounts(OT_Header) = 1
  
  Set ScObjects(1) = CreateNewInstance(OT_Background)
  ScObjects(0).LoadAPTData FileData
  If ScObjects(1).ObjectIndex = 0 Then
    Set ScObjects(1) = Nothing
    I = 1
  Else ' If TypeOf ScObjects(1) Is clsBackground Then
    ObjectCounts(OT_Background) = 1
    I = 2
  End If
  Count = I - 1
  
  Do Until EOF(FileNum)
    With FileData
      .ReadSection
      
      Select Case LCase$(.Section)
        Case "position"
          Set NewObject = CreateNewInstance(OT_MenuEntry)
          NewObject.LoadAPTData FileData
          ObjectCounts(OT_MenuEntry) = ObjectCounts(OT_MenuEntry) + 1
          NewObject.ObjectIndex = ObjectCounts(OT_MenuEntry)
        Case "atis"
          Set NewObject = CreateNewInstance(OT_ATIS)
          NewObject.LoadAPTData FileData
          ObjectCounts(OT_ATIS) = ObjectCounts(OT_ATIS) + 1
          NewObject.ObjectIndex = ObjectCounts(OT_ATIS)
        Case "ndb"
          Set NewObject = CreateNewInstance(OT_NDB)
          NewObject.LoadAPTData FileData
          ObjectCounts(OT_NDB) = ObjectCounts(OT_NDB) + 1
          NewObject.ObjectIndex = ObjectCounts(OT_NDB)
        Case "vor"
          Set NewObject = CreateNewInstance(OT_VOR)
          NewObject.LoadAPTData FileData
          ObjectCounts(OT_VOR) = ObjectCounts(OT_VOR) + 1
          NewObject.ObjectIndex = ObjectCounts(OT_VOR)
        Case "piste"
          Set NewObject = CreateNewInstance(OT_Runway)
          NewObject.LoadAPTData FileData
          ObjectCounts(OT_Runway) = ObjectCounts(OT_Runway) + 1
          NewObject.ObjectIndex = ObjectCounts(OT_Runway)
        Case "surfacetype"
          Set NewObject = CreateNewInstance(OT_SurfaceArea)
          NewObject.LoadAPTData FileData
          ObjectCounts(OT_SurfaceArea) = ObjectCounts(OT_SurfaceArea) + 1
          NewObject.ObjectIndex = ObjectCounts(OT_SurfaceArea)
        Case "polygone"
          Set NewObject = CreateNewInstance(OT_Polygon)
          NewObject.LoadAPTData FileData
          ObjectCounts(OT_Polygon) = ObjectCounts(OT_Polygon) + 1
          NewObject.ObjectIndex = ObjectCounts(OT_Polygon)
        Case "route", "roadtaxi"
          ' Route also covers Taxiway and River
          Set NewObject = CreateNewInstance(OT_Taxiway)
          NewObject.LoadAPTData FileData
          ObjectCounts(NewObject.ObjectType) = ObjectCounts(NewObject.ObjectType) + 1
          NewObject.ObjectIndex = ObjectCounts(NewObject.ObjectType)
        Case "line"
          Set NewObject = CreateNewInstance(OT_Line)
          NewObject.LoadAPTData FileData
          ObjectCounts(OT_Line) = ObjectCounts(OT_Line) + 1
          NewObject.ObjectIndex = ObjectCounts(OT_Line)
        Case "flatten"
          Set NewObject = CreateNewInstance(OT_FlatArea)
          NewObject.LoadAPTData FileData
          ObjectCounts(OT_FlatArea) = ObjectCounts(OT_FlatArea) + 1
          NewObject.ObjectIndex = ObjectCounts(OT_FlatArea)
        Case "batiment"
          Set NewObject = CreateNewInstance(OT_Building)
          NewObject.LoadAPTData FileData
          ObjectCounts(OT_Building) = ObjectCounts(OT_Building) + 1
          NewObject.ObjectIndex = ObjectCounts(OT_Building)
        Case "poly3d", "fuel"
          Set NewObject = CreateNewInstance(OT_Macro)
          NewObject.LoadAPTData FileData
          ObjectCounts(OT_Macro) = ObjectCounts(OT_Macro) + 1
          NewObject.ObjectIndex = ObjectCounts(OT_Macro)
        Case "exclude"
          Set NewObject = CreateNewInstance(OT_Exclusion)
          NewObject.LoadAPTData FileData
          ObjectCounts(OT_Exclusion) = ObjectCounts(OT_Exclusion) + 1
          NewObject.ObjectIndex = ObjectCounts(OT_Exclusion)
        Case "ils"
          strTemp = ReadNext(FileData.GetData("ID"), " ")
          
          For J = 1 To Count
            If TypeOf Scenery(J) Is clsRunway Then
              Set RunwayObject = Scenery(J)
              If StrComp(RunwayObject.Near.ILSID, strTemp, vbTextCompare) = 0 Then
                RunwayObject.Near.ILSBeamWidth = ValEx(.GetData("width"))
                Exit For
              End If
              If StrComp(RunwayObject.Far.ILSID, strTemp, vbTextCompare) = 0 Then
                RunwayObject.Far.ILSBeamWidth = ValEx(.GetData("width"))
                Exit For
              End If
            End If
          Next J
        Case "fs2000"
          Set NewObject = CreateNewInstance(OT_TowerLocation)
          NewObject.SceneryIndex = I
          Count = I
          Set ScObjects(I) = NewObject
          TowerIndex = I
          Header.LoadFS2000APTData FileData
          If TowerIndex > 0 Then
            ObjectCounts(OT_TowerLocation) = ObjectCounts(OT_TowerLocation) + 1
            NewObject.ObjectIndex = ObjectCounts(OT_TowerLocation)
            I = I + 1
          Else
            Set ScObjects(I) = Nothing
            Count = I - 1
          End If
          Set NewObject = Nothing
      End Select
      
      If Not NewObject Is Nothing Then
        If Not TypeOf NewObject Is clsShape Then
          NewObject.SceneryIndex = I
          Count = I
          Set ScObjects(I) = NewObject
          I = I + 1
        Else
          Set ShapeObject = NewObject
          With ShapeObject
            Count = I + .NumPoints
            For J = 0 To .NumPoints
              .Point(J).SceneryIndex = I + J
              Set ScObjects(I + J) = .Point(J)
            Next J
            I = I + .NumPoints + 1
          End With
          If ShapeObject.ShapeType = OT_Polygon Then
            ' Potentially a surfacetype
            Set SurfaceAreaObject = CreateNewInstance(OT_SurfaceArea)
            With SurfaceAreaObject
              .LoadAPTFromPolygon FileData
              If .mDrawable = True Then
                ObjectCounts(OT_SurfaceArea) = ObjectCounts(OT_SurfaceArea) + 1
                .ObjectIndex = ObjectCounts(OT_SurfaceArea)
                .SceneryIndex = I
                Count = I
                Set ScObjects(I) = SurfaceAreaObject
                I = I + 1
              End If
            End With
          End If
          Set ShapeObject = Nothing
        End If
      End If
    End With
    Set NewObject = Nothing
    
    ' Speeds up loading by segmenting the progress indicator
    Counter = Counter + 1
    If (Counter Mod 20) = 0 Then
      If frmMain.picProgress.Visible Then
        frmMain.barProgress.Value = Seek(FileNum)
      End If
    End If
  Loop

'  For I = 0 To NumObjects
'    If ScObjects(I).ObjectType = OT_Point Then
'      Set PointObject = ScObjects(I)
'      If PointObject.ObjectIndex = 1 Then
'        If PointObject.Parent.LinkTo > 0 Then
'          Set PointObject.Parent.LinkObject = LinkMap(PointObject.Parent.LinkTo)
'        End If
'        PointObject.UpdateObject
'      End If
'    End If
'  Next I
'
'  For I = 0 To UBound(LinkMap)
'    Set LinkMap(I) = Nothing
'  Next I
'
'  If VersionData <= 15 Then
'    DoOldVersionConversion
'  End If

  With frmMain
    If .picProgress.Visible Then
      .barProgress.Value = LOF(FileNum)
      .picProgress.Visible = False
      .StatusBar.Visible = True
    End If
  End With

  Close #FileNum

  ResortScenery
End Sub

' Loads file data (FSSC Ascii format)
Private Sub LoadAsciiFile(ByVal Filename As String)
  Dim FileNum As Integer, FileData As clsFileData
  
  Dim VersionData As Integer, ApproxCount As Integer
  Dim I As Integer, J As Integer
  Dim ToolbarRestore As Boolean
  
  Dim LinkMap() As clsObject
  
  Dim ShapeObject As clsShape, NewObject As clsObject, _
    PointObject As clsPoint
    
  Dim BuildingObject As clsBuilding, MacroObject As clsMacro
  Dim LinkFromNum As Integer

  FileNum = FreeFile

  Open Filename For Input As #FileNum
  Set FileData = New clsFileData
  With FileData
    .FileNum = FileNum
    .ReadSection
    If UCase$(.Section) <> "GENERAL" Then Error 1
    .GetData "Version", VersionData, 0
    .GetData "Count", ApproxCount, -1
    .GetData "AFDRefresh", AFDRefresh, True
    .GetData "ExportPath", ExportPath, ""
  End With
  If VersionData > MyFileVer Then Error 2
  ' Can't set max = min.
  If ApproxCount = 0 Then ApproxCount = 1
  
  ' Progress bar
  With frmMain
    If .StatusBar.Visible Then
      .barProgress.Value = 0
      .lblStatus.Caption = Lang.GetString(RES_Main_Loading)
      .barProgress.Max = ApproxCount
      .StatusBar.Visible = False
      .picProgress.Visible = True
      .picProgress.Refresh
      ToolbarRestore = True
    End If
  End With
  
  ReDim LinkMap(20)
  
  I = 0
  Do Until EOF(FileNum)
    With FileData
      .ReadSection
      For J = 0 To OT_Point
        If StrComp(.Section, ObjectNames(J), vbTextCompare) = 0 Then
          Exit For
        End If
      Next J

      
      Set NewObject = CreateNewInstance(J)
      If J < OT_Point Then
        If TypeOf NewObject Is clsShape Then
          Set ShapeObject = NewObject
          ShapeObject.LoadData FileData, VersionData
          LinkFromNum = ShapeObject.LinkFrom
        Else
          NewObject.LoadData FileData, VersionData
        End If
        
        If J = OT_TowerLocation Then TowerIndex = I

        ObjectCounts(J) = ObjectCounts(J) + 1
        NewObject.ObjectIndex = ObjectCounts(J)
      ElseIf J = OT_Point Then
        NewObject.LoadData FileData, VersionData
        Set PointObject = NewObject
        Set PointObject.Parent = ShapeObject
        ShapeObject.AddPoint PointObject
      End If
      
      If Not NewObject Is Nothing Then
        If LinkFromNum > 0 Then
          If LinkFromNum > UBound(LinkMap) Then
            ReDim Preserve LinkMap(LinkFromNum * 2)
          End If
          Set LinkMap(LinkFromNum) = NewObject
          LinkFromNum = 0
        End If
      
        If Not TypeOf NewObject Is clsShape Then
          NewObject.SceneryIndex = I
          Count = I
          Set ScObjects(I) = NewObject
          I = I + 1
        End If
      End If
    End With
    
    ' Speeds up loading by segmenting the progress indicator
    If (I Mod 50) = 0 Then
      If ToolbarRestore Then
        frmMain.barProgress.Value = IIf(ApproxCount < I, ApproxCount, I)
      End If
      DoEvents
    End If
  Loop
  
  For I = 0 To NumObjects
    If ScObjects(I).ObjectType = OT_Point Then
      Set PointObject = ScObjects(I)
      If PointObject.ObjectIndex = 1 Then
        If PointObject.Parent.LinkTo > 0 Then
          Set PointObject.Parent.LinkObject = LinkMap(PointObject.Parent.LinkTo)
        End If
        PointObject.UpdateObject
      End If
    End If
  Next I
  
  For I = 0 To UBound(LinkMap)
    Set LinkMap(I) = Nothing
  Next I

  If VersionData <= 15 Then
    DoOldVersionConversion
  End If

  With frmMain
    If ToolbarRestore Then
      .barProgress.Value = ApproxCount
      .picProgress.Visible = False
      .StatusBar.Visible = True
    End If
  End With

  Close #FileNum
End Sub

' Loads file data (Binary format)
Private Sub LoadBinaryFile(ByVal Filename As String)
  Dim I As Integer, T As Integer, _
    FileNum As Integer, VerMinor As Integer, _
    OType As Byte, Temp As String, _
    Start As Integer, ActualCount As Integer, _
    PointObject As clsPoint, OType2 As ObjectTypeEnum
  Dim NewObject As clsObject
  Dim TempInt As Integer
  Dim ToolbarRestore As Boolean

  Dim PolygonMap() As clsShape

  ', CityObjectWarned As Boolean

  FileNum = FreeFile

  Open Filename For Binary As #FileNum
  
  ' Get the File Header information
  ' Identifier, EOF to prevent text in DOS type, Version Number
  Temp = Space$(24)
  Get #FileNum, , Temp
'  If Temp <> "FS Scenery Creator File" & Chr$(26) Then Error 1
  Get #FileNum, , T
'  If T > App.Major Then Error 2
  Get #FileNum, , VerMinor
  If VerMinor > MyFileVer Then Error 2
  Get #FileNum, , T
  
  ' Take care of synthetic scenery which used to  be a
  ' separate object.
'  If VerMinor < 3 Then T = T - 1
  If VerMinor < 3 Then
    ' Version 1.0.x not supported anymore
    Err.Raise 1
  End If
  ' Take care of background image whic used to be part
  ' of the header.
  If VerMinor < 8 Then T = T + 1
  Count = T

  If T = 0 Then T = 1
  ' Progress bar
  With frmMain
    If .StatusBar.Visible Then
      .barProgress.Value = 0
      .lblStatus.Caption = Lang.GetString(RES_Main_Loading)
      .barProgress.Max = T
      .StatusBar.Visible = False
      .picProgress.Visible = True
      .picProgress.Refresh
      ToolbarRestore = True
    End If
  End With

  ' Header
  Get #FileNum, , OType
  Set ScObjects(0) = CreateNewInstance(OT_Header)
  ScObjects(0).ObjectIndex = 0
  ObjectCounts(OT_Header) = 1
  
  If VerMinor < 8 Then
    Set ScObjects(1) = CreateNewInstance(OT_Background)
    ScObjects(0).LoadBinaryData FileNum, VerMinor, 0
    If ScObjects(1).ObjectIndex = 0 Then
      Count = Count - 1
      Start = 0
      Set ScObjects(1) = Nothing
    ElseIf TypeOf ScObjects(1) Is clsBackground Then
      ObjectCounts(OT_Background) = 1
      Start = 1
    End If
  Else
    Start = 0
    ScObjects(0).LoadBinaryData FileNum, VerMinor, 0
  End If
  ActualCount = Start
  ReDim PolygonMap(T)

  ' Get the data records from the file
  For I = 1 + Start To Count
    Get #FileNum, , OType
    Select Case OType
      Case 0 ' Header
        Error 1
      Case 1 ' Synthetic
        Error 1
      Case 2 ' ATIS
        Set NewObject = CreateNewInstance(OT_ATIS)
        NewObject.LoadBinaryData FileNum, VerMinor, 0
        ObjectCounts(OT_ATIS) = ObjectCounts(OT_ATIS) + 1
        NewObject.ObjectIndex = ObjectCounts(OT_ATIS)
      Case 3 ' VOR
        Set NewObject = CreateNewInstance(OT_VOR)
        NewObject.LoadBinaryData FileNum, VerMinor, 0
        ObjectCounts(OT_VOR) = ObjectCounts(OT_VOR) + 1
        NewObject.ObjectIndex = ObjectCounts(OT_VOR)
      Case 4 ' NDB
        Set NewObject = CreateNewInstance(OT_NDB)
        NewObject.LoadBinaryData FileNum, VerMinor, 0
        ObjectCounts(OT_NDB) = ObjectCounts(OT_NDB) + 1
        NewObject.ObjectIndex = ObjectCounts(OT_NDB)
      Case 5 ' ILS
        Error 1
      Case 6 ' Runway
        Set NewObject = CreateNewInstance(OT_Runway)
        NewObject.LoadBinaryData FileNum, VerMinor, 0
        ObjectCounts(OT_Runway) = ObjectCounts(OT_Runway) + 1
        NewObject.ObjectIndex = ObjectCounts(OT_Runway)
      Case 7 ' MenuEntry
        Set NewObject = CreateNewInstance(OT_MenuEntry)
        NewObject.LoadBinaryData FileNum, VerMinor, 0
        ObjectCounts(OT_MenuEntry) = ObjectCounts(OT_MenuEntry) + 1
        NewObject.ObjectIndex = ObjectCounts(OT_MenuEntry)
      Case 8 ' Polygon
        Set NewObject = CreateNewInstance(OT_Polygon)
        NewObject.LoadBinaryData FileNum, VerMinor, 0
        Set PolygonMap(I) = NewObject
        ObjectCounts(OT_Polygon) = ObjectCounts(OT_Polygon) + 1
        NewObject.ObjectIndex = ObjectCounts(OT_Polygon)
      Case 9 ' Taxiway
        Set NewObject = CreateNewInstance(OT_Taxiway)
        NewObject.LoadBinaryData FileNum, VerMinor, 0
        Set PolygonMap(I) = NewObject
        ObjectCounts(OT_Taxiway) = ObjectCounts(OT_Taxiway) + 1
        NewObject.ObjectIndex = ObjectCounts(OT_Taxiway)
      Case 10 ' Road
        Set NewObject = CreateNewInstance(OT_Road)
        NewObject.LoadBinaryData FileNum, VerMinor, 0
        Set PolygonMap(I) = NewObject
        ObjectCounts(OT_Road) = ObjectCounts(OT_Road) + 1
        NewObject.ObjectIndex = ObjectCounts(OT_Road)
      Case 11 ' River
        Set NewObject = CreateNewInstance(OT_River)
        NewObject.LoadBinaryData FileNum, VerMinor, 0
        Set PolygonMap(I) = NewObject
        ObjectCounts(OT_River) = ObjectCounts(OT_River) + 1
        NewObject.ObjectIndex = ObjectCounts(OT_River)
      Case 12 ' Line
        Set NewObject = CreateNewInstance(OT_Line)
        NewObject.LoadBinaryData FileNum, VerMinor, 0
        Set PolygonMap(I) = NewObject
        ObjectCounts(OT_Line) = ObjectCounts(OT_Line) + 1
        NewObject.ObjectIndex = ObjectCounts(OT_Line)
      Case 13 ' Building
        Set NewObject = CreateNewInstance(OT_Building)
        NewObject.LoadBinaryData FileNum, VerMinor, 0
        ObjectCounts(OT_Building) = ObjectCounts(OT_Building) + 1
        NewObject.ObjectIndex = ObjectCounts(OT_Building)
      Case 14 ' City
        Set NewObject = CreateNewInstance(OT_Macro)
        NewObject.LoadBinaryData FileNum, VerMinor, 1
        ObjectCounts(OT_Macro) = ObjectCounts(OT_Macro) + 1
        NewObject.ObjectIndex = ObjectCounts(OT_Macro)
      Case 15 ' FuelArea
        Set NewObject = CreateNewInstance(OT_Macro)
        NewObject.LoadBinaryData FileNum, VerMinor, 2
        ObjectCounts(OT_Macro) = ObjectCounts(OT_Macro) + 1
        NewObject.ObjectIndex = ObjectCounts(OT_Macro)
      Case 16 ' Macro
        Set NewObject = CreateNewInstance(OT_Macro)
        NewObject.LoadBinaryData FileNum, VerMinor, 0
        ObjectCounts(OT_Macro) = ObjectCounts(OT_Macro) + 1
        NewObject.ObjectIndex = ObjectCounts(OT_Macro)
      Case 17 ' Building2K
        Set NewObject = CreateNewInstance(OT_Building)
        NewObject.LoadBinaryData FileNum, VerMinor, 1
        ObjectCounts(OT_Building) = ObjectCounts(OT_Building) + 1
        NewObject.ObjectIndex = ObjectCounts(OT_Building)
      Case 18 ' Background
        Set NewObject = CreateNewInstance(OT_Background)
        NewObject.LoadBinaryData FileNum, VerMinor, 0
        ObjectCounts(OT_Background) = ObjectCounts(OT_Background) + 1
        NewObject.ObjectIndex = ObjectCounts(OT_Background)
      Case 19 ' FlatArea
        Set NewObject = CreateNewInstance(OT_FlatArea)
        NewObject.LoadBinaryData FileNum, VerMinor, 0
        Set PolygonMap(I) = NewObject
        ObjectCounts(OT_FlatArea) = ObjectCounts(OT_FlatArea) + 1
        NewObject.ObjectIndex = ObjectCounts(OT_FlatArea)
      Case 30 ' Point
        Set NewObject = CreateNewInstance(OT_Point)
        NewObject.LoadBinaryData FileNum, VerMinor, 0
        TempInt = NewObject.SceneryIndex
        If PolygonMap(TempInt) Is Nothing Then
          TempInt = TempInt + 1
          Do
            If Not PolygonMap(TempInt) Is Nothing Then
              Exit Do
            End If
            TempInt = TempInt - 1
            If TempInt = 0 Then Error 1
          Loop
        End If
        Set PointObject = NewObject
        Set PointObject.Parent = PolygonMap(TempInt)
        Set PolygonMap(TempInt).Point(PointObject.ObjectIndex - 1) = PointObject
        Set PointObject = Nothing
      Case Else
        Error 1
    End Select
    
    If Not TypeOf NewObject Is clsShape Then
      ActualCount = ActualCount + 1
      NewObject.SceneryIndex = ActualCount
      Set ScObjects(ActualCount) = NewObject
    End If
    
    ' Speeds up loading by segmenting the progress indicator
    If (I Mod 50) = 0 Then
      If ToolbarRestore Then
        frmMain.barProgress.Value = I
      End If
'      DoEvents
    End If
  Next I

  Count = ActualCount

  For I = 0 To NumObjects
    If ScObjects(I).ObjectType = OT_Point Then
      Set PointObject = ScObjects(I)
      If PointObject.ObjectIndex = 1 Then
        PointObject.UpdateObject
      End If
    End If
  Next I
  
  DoOldVersionConversion

  With frmMain
    If ToolbarRestore Then
      .barProgress.Value = Count
      .picProgress.Visible = False
      .StatusBar.Visible = True
    End If
  End With

  Close #FileNum
 
  ResortScenery
End Sub

' Loads file data
Public Sub LoadFile(ByVal Filename As String)
  Dim FileNum As Integer, Temp As String
  On Error GoTo ErrorH:
  
  SetScreenMousePointer vbHourglass
  frmMain.Refresh
  If Not FileExists(Filename) Then Error 53

  ChangeDir GetDir(Filename)
  Clear
  
  If StrComp(Right$(Filename, 4), ".apt", vbTextCompare) = 0 Then
    LoadAPTFile Filename
  ElseIf StrComp(Right$(Filename, 4), ".sca", vbTextCompare) = 0 Then
    LoadSCASMFile Filename
  Else
    FileNum = FreeFile
    Open Filename For Binary As #FileNum
    ' Get the File Header information
    ' Identifier, EOF to prevent text in DOS type, Version Number
    Temp = Space$(24)
    Get #FileNum, , Temp
    Close #FileNum
    If Temp = "FS Scenery Creator File" & Chr$(26) Then
      LoadBinaryFile Filename
    Else
      LoadAsciiFile Filename
    End If
  End If

  ' Clean up
  File = Filename
  Changed = False
  ' ClearValues will initiate a draw
  RaiseEvent ClearValues
  RaiseEvent TitleBarChange
  PurgeAutoSaves
  NextAutoSave = (Timer + Options.AutoSave * 60) Mod 86400
  SetScreenMousePointer vbDefault
EndErr:
  frmMain.UpdateObjectList
  Close
  Exit Sub
ErrorH:
  SetScreenMousePointer vbDefault
  Select Case Err.Number
    Case 1
      ' Invalid file format
      MsgBoxEx frmMain, Lang.GetString(RES_ERR_InvalidFormat), vbCritical Or vbOKOnly, RES_ERR_InvalidFormat
    Case 2
      ' FSSC Version too high
      MsgBoxEx frmMain, Lang.GetString(RES_ERR_HigherFSSCVersion), vbCritical Or vbOKOnly, RES_ERR_HigherFSSCVersion
    Case 3
      ' Canceled Switch FS Version
    Case Else
      MsgBoxEx frmMain, Error$, vbCritical Or vbOKOnly, IDH_ERR_OPENFILE
  End Select
  
  With frmMain
    If .picProgress.Visible Then
      .picProgress.Visible = False
      .StatusBar.Visible = True
    End If
  End With
  NewFile
  Resume EndErr:
End Sub

' Loads file data (scasm code)
Private Sub LoadSCASMFile(ByVal Filename As String)
  MsgBoxEx frmMain, "Sorry, SCASM code file import not implemented yet. Thank you for your patience.", vbInformation, 0
End Sub

' Checks to see if the cursor is inside the focus area
' denoted by the circle, and returns the indices in
' an array
Public Function NearestObject(ByVal X As Single, ByVal Y As Single, ByVal Topmost As Boolean) As Integer
  Dim I As Integer, Radius2 As Single, _
    Low As Integer, High As Integer, myStep As Integer
  Radius2 = FocusWidth * FocusWidth

  Low = IIf(Topmost, Count, 1)
  High = IIf(Topmost, 1, Count)
  myStep = IIf(Topmost, -1, 1)

  For I = Low To High Step myStep
    With ScObjects(I)
      If Distance2(X, Y, .PositionX, .PositionY) < Radius2 Then
        NearestObject = I
        Exit For
      End If
    End With
  Next I
End Function

' Checks whether we can snap a point to a line
Public Function NearestLine(ByRef X As Single, ByRef Y As Single) As Boolean
  Dim I As Integer, Temp As clsPoint
  
  For I = NumObjects To 1 Step -1
    If TypeOf ScObjects(I) Is clsPoint Then
      If ScObjects(I).ObjectIndex = 1 Then
        Set Temp = ScObjects(I)
        ' Test the new point
        If Temp.Parent.TestPoint2(X, Y) Then
          NearestLine = True
          Set Temp = Nothing
          Exit For
        End If
        Set Temp = Nothing
      End If
    End If
  Next I
End Function

' Checks whether we can snap a point to a line
Public Function NearestPoint(ByRef X As Single, ByRef Y As Single) As Boolean
  Dim I As Integer, Threshold As Single
  
  Threshold = picEditor.ScaleX(Point_Threshold)
  Threshold = Threshold * Threshold

  For I = NumObjects To 1 Step -1
    If TypeOf ScObjects(I) Is clsPoint Then
      With ScObjects(I)
        If Not Focus(I) Then
          If Distance2(.PositionX, .PositionY, X, Y) <= Threshold Then
            X = .PositionX
            Y = .PositionY
            NearestPoint = True
            Exit For
          End If
        End If
      End With
    End If
  Next I
End Function

' Prepares for new file
Public Sub NewFile()
  Clear
  
  RaiseEvent ClearValues
  RaiseSingleFocusChangeEvent 0
  PurgeAutoSaves
  NextAutoSave = (Timer + Options.AutoSave * 60) Mod 86400
  
  frmMain.UpdateObjectList
End Sub

' Checks if the user clicked on a clsShape line
' If so, creates a new point and returns true
Public Function OnPolyLine(ByVal X As Single, ByVal Y As Single) As Boolean
  Dim I As Integer, Temp As clsPoint

  For I = NumObjects To 1 Step -1
    If TypeOf ScObjects(I) Is clsPoint Then
      If ScObjects(I).ObjectIndex = 1 Then
        Set Temp = ScObjects(I)
        ' Test the new point
        If Temp.Parent.TestPoint(X, Y) Then
          OnPolyLine = True
          Set Temp = Nothing
          Exit For
        End If
        Set Temp = Nothing
      End If
    End If
  Next I
End Function

' Sort a list
Private Sub QuickSort(SortList() As Integer, ByVal First As Integer, ByVal Last As Integer)
  Dim Low As Integer, High As Integer, _
    Temp As Integer, TestElement As Integer
  If First > Last Then Exit Sub
  Low = First
  High = Last
  TestElement = SortList((First + Last) / 2)  'Select an element from the middle.
  Do
    Do While ScObjects(SortList(Low)).SortIndex < ScObjects(TestElement).SortIndex     'Find lowest element that is >= TestElement.
      Low = Low + 1
    Loop
    Do While ScObjects(SortList(High)).SortIndex > ScObjects(TestElement).SortIndex    'Find highest element that is <= TestElement.
      High = High - 1
    Loop
    If (Low <= High) Then             'If not done,
      Temp = SortList(Low)            ' Swap the elements.
      SortList(Low) = SortList(High)
      SortList(High) = Temp
      Low = Low + 1
      High = High - 1
    End If
  Loop While (Low <= High)
  If (First < High) Then QuickSort SortList, First, High
  If (Low < Last) Then QuickSort SortList, Low, Last
End Sub

' Raise the FocusChanged event with a single element
Public Sub RaiseSingleFocusChangeEvent(ByVal Index As Integer)
  If Index = 0 Then
    ReDim Indices(0) As Integer
  Else
    ReDim Indices(1) As Integer
    Indices(1) = Index
  End If
  RaiseEvent FocusChanged(Indices)
End Sub

' Call the FocusChanged event
Public Sub RefreshFocus(ByVal Topmost As Boolean)
  Dim I As Integer, Cnt As Integer, _
    Low As Integer, High As Integer, myStep As Integer

  ReDim Indices(10) As Integer
  
  Low = IIf(Topmost, Count, 1)
  High = IIf(Topmost, 1, Count)
  myStep = IIf(Topmost, -1, 1)

  For I = Low To High Step myStep
    If mFocus(I) Then
      Cnt = Cnt + 1
      If Cnt > UBound(Indices) Then ReDim Preserve Indices(Cnt * 2)
      Indices(Cnt) = I
    End If
  Next I
  ReDim Preserve Indices(Cnt)
  
  RaiseEvent FocusChanged(Indices)
  RaiseEvent Redraw
End Sub

' Sorts scenery by the default order
Public Sub ResortScenery()
  Dim I As Integer
  Dim Objs() As Integer, NewObjects() As clsObject
  
  ReDim Objs(NumObjects)
  ReDim NewObjects(NumObjects)
  For I = 0 To NumObjects
    Objs(I) = I
    Set NewObjects(I) = ScObjects(I)
  Next I
  QuickSort Objs, 0, NumObjects
  For I = 0 To NumObjects
    Set ScObjects(I) = NewObjects(Objs(I))
    ScObjects(I).SceneryIndex = I
    Set NewObjects(Objs(I)) = Nothing
  Next I
  Draw
End Sub

' Saves the scenery to a file
Public Sub SaveFile(ByVal Filename As String, Optional ByVal AutoSave As Boolean = False)
  Dim I As Integer, PointObject As clsPoint
  Dim FileNum As Integer, FileData As clsFileData
  
  ' Progress bar
  With frmMain
    If .StatusBar.Visible And Count > 0 Then
      .barProgress.Value = 0
      .lblStatus.Caption = Lang.GetString(IIf(AutoSave, RES_Main_Autosaving, RES_Main_Saving))
      .barProgress.Max = Count
      .StatusBar.Visible = False
      .picProgress.Visible = True
      .picProgress.Refresh
    End If
  End With
  
  LinkCounter = 1
  For I = 1 To Count
    If ScObjects(I).ObjectType = OT_Point Then
      Set PointObject = ScObjects(I)
      If PointObject.ObjectIndex = 1 Then
        PointObject.Parent.PrepareLink
      End If
      Set PointObject = Nothing
    End If
  Next I
  
  FileNum = FreeFile
  Open Filename For Output As #FileNum
  Set FileData = New clsFileData

  With FileData
    .FileNum = FileNum
    .InitializeForWrite "General", "General"
    .SaveData "Version", MyFileVer, 0
    .SaveData "Count", NumObjects, -1
    .SaveData "AFDRefresh", AFDRefresh, True
    .SaveData "ExportPath", ExportPath, ""
    .WriteSection
  End With

  For I = 0 To Count
    ScObjects(I).SaveData FileData
    If (I Mod 50) = 0 Then
      If frmMain.picProgress.Visible Then
        frmMain.barProgress.Value = I
      End If
    End If
  Next I
  Close #FileNum

  If Not AutoSave Then
    File = Filename
    Changed = False
  End If
  NextAutoSave = (Timer + Options.AutoSave * 60) Mod 86400

  ' Progress bar
  With frmMain
    If .picProgress.Visible And Count > 0 Then
      .barProgress.Value = Count
      .picProgress.Visible = False
      .StatusBar.Visible = True
    End If
  End With
End Sub

' Checks to see what has the focus within the bounded
' rectangle
Public Sub SelectRectFocus(ByVal Left As Single, ByVal Top As Single, ByVal Right As Single, ByVal Bottom As Single)
  Dim I As Integer, FocusRgn As Long, _
    Temp As Single, Cnt As Integer

  ReDim Indices(10) As Integer

  ' Change the coordinates from min to max
  If Right < Left Then
    Temp = Right
    Right = Left
    Left = Temp
  End If
  If Bottom < Top Then
    Temp = Bottom
    Bottom = Top
    Top = Temp
  End If
  
  ' Prepare Rect for RectInRegion
  FocusRgn = CreateRectRgn(Left, Top, Right, Bottom)

  ' Loop to see what is inside the rectangle
  For I = NumObjects To 1 Step -1
    If (PtInRegion(FocusRgn, ScObjects(I).PositionX, ScObjects(I).PositionY) <> 0) Then
      If Options.ObjectVisible(ScObjects(I).ObjectType) Then
        Cnt = Cnt + 1
        If Cnt > UBound(Indices) Then ReDim Preserve Indices(Cnt * 2)
        Indices(Cnt) = I
        Focus(I) = True
      Else
        Focus(I) = False
      End If
    Else
      Focus(I) = False
    End If
  Next I
  DeleteRegion FocusRgn
  ReDim Preserve Indices(Cnt)
  
  RaiseEvent FocusChanged(Indices)
  RaiseEvent Redraw
End Sub

Public Sub SendToBack()
  Dim I As Integer, UpTo As Integer
  Dim MyObject As clsObject, _
    PointObject As clsPoint, PointObject2 As clsPoint

  Set MyObject = Scenery(SingleFocus)
  If MyObject.ObjectType = OT_Point Then
    Set PointObject = MyObject
    Set MyObject = PointObject.Parent.Point(0)
    UpTo = MyObject.SceneryIndex
    PointObject.Parent.ObjectIndex = 1
  Else
    UpTo = SingleFocus
  End If
  
  ClearFocus True
  
  For I = UpTo To 2 Step -1
    Set ScObjects(I) = ScObjects(I - 1)
    ScObjects(I).SceneryIndex = I
    If MyObject.ObjectType <> OT_Point Then
      If ScObjects(I).ObjectType = MyObject.ObjectType Then
        ScObjects(I).ObjectIndex = ScObjects(I).ObjectIndex + 1
      End If
    Else
      If ScObjects(I).ObjectType = OT_Point Then
        If ScObjects(I).ObjectIndex = 1 Then
          Set PointObject2 = ScObjects(I)
          If PointObject.Parent.ShapeType = PointObject2.Parent.ShapeType Then
            PointObject2.Parent.ObjectIndex = PointObject2.Parent.ObjectIndex + 1
          End If
        End If
      End If
    End If
  Next I
  MyObject.SceneryIndex = 1
  MyObject.ObjectIndex = 1
  Set ScObjects(1) = MyObject

  frmMain.UpdateObjectList
  Changed = True

  Scenery.Draw
End Sub

' Clears the focus and sets the focus to Index
Public Sub SetSingleFocus(ByVal Index As Integer)
  ClearFocus False
  Focus(Index) = True
  RaiseEvent Redraw
  RaiseSingleFocusChangeEvent Index
End Sub

' Transform the scenery objects about the origin and change
' heading properties accordingly
Public Sub TransformScenery(ByVal X As Single, ByVal Y As Single, ByVal Rotation As Single, ByVal TransformAll As Boolean, ByVal ChangeRotation As Boolean)
  On Error Resume Next
  
  Dim I As Integer, myPos As PointType
  ' Going backwards ensures that the points are processed
  ' before the shape
  For I = Count To 1 Step -1
    If TransformAll Or mFocus(I) Then
      With ScObjects(I)
        myPos = MakePoint(.PositionX, .PositionY)
        Rotate myPos, Rotation
        If ChangeRotation Then
          .Rotation = EnsureRotation(.Rotation + Rotation)
        End If
        .PositionX = Round(myPos.X + X, 2)
        .PositionY = Round(myPos.Y + Y, 2)
        .UpdateObject
      End With
    End If
  Next I
  Changed = True
  ClearFocus True
  Draw
End Sub

Public Sub Update(ByVal NewVersion As Integer)
  Dim I As Long
  For I = 0 To NumObjects
    With ScObjects(I)
      .VersionChanged NewVersion
      .UpdateObject
    End With
  Next I
End Sub

Private Sub WriteSceneryCFGFile(ByVal File As String)
    Dim Section As String, Res As String, _
    Res2 As String, myPath As String, _
    I As Integer, J As Integer, K As Integer, _
    L As Integer, M As Integer, _
    PointObject As clsPoint, ShapeObject As clsShape
  Dim Attr As VbFileAttribute
  
  If Options.FSVersion >= Version_FS2K Then
    I = 1
    Do
      Section = "Area." & Format$(I, "000")
      Res = GetINIString(Section, "Title", "", File)
      Res2 = GetINIString(Section, "Local", "", File)
      If Mid$(Res2, 2, 1) <> ":" Then Res2 = AddDir(Options.FSPath, Res2)
      If Res = "" Or (StrComp(Res, Header.Name, vbTextCompare) = 0) Or (StrComp(Res2, ExportPath, vbTextCompare) = 0) Then
        ' Found a free layer
        If StrComp(Left$(ExportPath, Len(Options.FSPath)), Options.FSPath, vbTextCompare) = 0 Then
          myPath = Mid$(ExportPath, Len(Options.FSPath) + 2)
        Else
          myPath = ExportPath
        End If
        If StrComp(Right$(myPath, 8), "\scenery", vbTextCompare) = 0 Then
          myPath = Left$(myPath, Len(myPath) - 8)
        End If
        If myPath = "" Then myPath = ExportPath
        
        Res = "Title=" & Header.Name & vbNullChar & _
              "Active=TRUE" & vbNullChar & _
              "Layer=" & Format$(I) & vbNullChar & _
              "Local=" & myPath & vbNullChar
         
        ' Flatten
        For J = 1 To Scenery.Count
          If Scenery(J).ObjectType = OT_Point Then
            Set PointObject = Scenery(J)
            If PointObject.ObjectIndex = 1 And PointObject.Parent.ShapeType = OT_FlatArea Then
              Set ShapeObject = PointObject.Parent
              With ShapeObject
                If ShapeObject.Extra3 = 0 Then
                  K = 0
                  L = 2
                  Do
                    M = L + 1
                    If M > .NumPoints Then M = L
                    Res = Res & "Flatten." & K & "=" & FloatFormat(Header.Altitude * MToFt, "####0.0#") & ", " & _
                                                       ReturnPoint(.Point(0).X, .Point(0).Y).LatLonConfig & ", " & _
                                                       ReturnPoint(.Point(L - 1).X, .Point(L - 1).Y).LatLonConfig & ", " & _
                                                       ReturnPoint(.Point(L).X, .Point(L).Y).LatLonConfig & ", " & _
                                                       ReturnPoint(.Point(M).X, .Point(M).Y).LatLonConfig & vbNullChar
                    L = L + 2
                    K = K + 1
                  Loop Until (L > .NumPoints) Or (K = 10)
                  Exit For
                End If
              End With
              Set ShapeObject = Nothing
            End If
            Set PointObject = Nothing
          End If
        Next J
        
        Res = Res & vbNullChar
        
        Attr = GetAttr(File)
        If Attr And vbReadOnly Then SetAttr File, Attr And Not vbReadOnly
        WritePrivateProfileSection Section, Res, File
        
        Exit Do
      End If
      I = I + 1
    Loop
  End If
End Sub

Private Sub Class_Initialize()
  NumObjects = -1
End Sub

Private Sub Class_Terminate()
  Dim I As Integer, PointObject As clsPoint, RunwayObject As clsRunway
  For I = 0 To NumObjects
    If Not ScObjects(I) Is Nothing Then
      If TypeOf ScObjects(I) Is clsPoint Then
        Set PointObject = ScObjects(I)
        Set PointObject.Parent = Nothing
        Set PointObject = Nothing
      ElseIf TypeOf ScObjects(I) Is clsRunway Then
        Set RunwayObject = ScObjects(I)
        Set RunwayObject.Near = Nothing
        Set RunwayObject.Far = Nothing
        Set RunwayObject = Nothing
      End If
      Set ScObjects(I) = Nothing
    End If
  Next I
  
  Set RunwayObject = Defaults(OT_Runway)
  If Not RunwayObject Is Nothing Then
    Set RunwayObject.Near = Nothing
    Set RunwayObject.Far = Nothing
  End If
  For I = 0 To MAX_OBJ
    Set Defaults(I) = Nothing
  Next I
End Sub
