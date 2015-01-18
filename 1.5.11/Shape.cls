VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Shape Class
' Provides functions for using Shapes in the program

Option Explicit
Implements clsObject

' Common properties
Public ObjectIndex As Integer, _
       Name As String, _
       Locked As Boolean, _
       Complexity As Byte

' General properties
Public Texture As String, _
       Color As Long, _
       Layer As Byte, _
       MScale As Single, _
       DotSpacing As Integer, _
       DotColor As Long, _
       LineWidth As Single, _
       LineColor As Long, _
       Extra1 As Single, _
       Extra2 As Boolean, _
       Extra3 As Byte, _
       LinkObject As clsObject, _
       Visibility As Long
       
Public LinkTo As Integer

Public LinkFrom As Integer, LinkFromObj As clsShape
       
Public ShapeType As ObjectTypeEnum

Private myPoints() As clsPoint
Private myNumPoints As Integer

Public mDrawable As Boolean
Private glBitmapID As Integer

Private Pts() As PointType
Private TessPts() As TessType
Private TexCoord() As TessType
Private SumX As Single, SumY As Single
Private AvgX As Single, AvgY As Single

Public BoolTag As Boolean, IntTag As Integer

Private Sub Class_Initialize()
  myNumPoints = -1
  ReDim myPoints(5)
  glBitmapID = -1
End Sub

Private Sub Class_Terminate()
  Dim I As Integer
  Dim ShapeObject As clsShape

  For I = 0 To myNumPoints
    Set myPoints(I) = Nothing
  Next I
  If Not LinkFromObj Is Nothing Then
    Set LinkFromObj.LinkObject = Nothing
    Set LinkFromObj = Nothing
  End If
  
  If Not LinkObject Is Nothing Then
    If LinkObject.ObjectType = OT_Polygon Then
      Set ShapeObject = LinkObject
      Set ShapeObject.LinkFromObj = Nothing
    End If
    Set LinkObject = Nothing
  End If
  Set LinkObject = Nothing
  If Closing Then Error 1
  If glBitmapID > -1 Then picEditor.DeleteBitmap glBitmapID
End Sub

Public Sub AddPoint(NewPoint As clsPoint, Optional ByVal ObjIndex As Integer)
  If ObjIndex > 0 Then
    If ObjIndex > NumPoints Then NumPoints = ObjIndex - 1
    Set myPoints(ObjIndex - 1) = NewPoint
  Else
    NumPoints = NumPoints + 1
    Set myPoints(NumPoints) = NewPoint
    NewPoint.ObjectIndex = NumPoints + 1
  End If
End Sub

Public Sub AddSuccess(NewPoint As clsPoint)
  Dim PointX As Single, PointY As Single, _
    Length As Single, SqrtThree As Single
  
  mDrawable = True
  
  SqrtThree = Sqr(3)
  Length = Round(picEditor.ScaleX(40), 2)
  
  PointX = NewPoint.X
  PointY = NewPoint.Y
  Set myPoints(0) = NewPoint
  Set myPoints(1) = New clsPoint
  Select Case ShapeType
    Case OT_Polygon, OT_FlatArea
      Set myPoints(2) = New clsPoint
      myPoints(0).SetAsNewPoint Me, PointX - Length, PointY - Round(Length / SqrtThree, 2), 1
      myPoints(1).SetAsNewPoint Me, PointX + Length, PointY - Round(Length / SqrtThree, 2), 2
      myPoints(2).SetAsNewPoint Me, PointX, PointY + Round(Length * 2 / SqrtThree, 2), 3
      myNumPoints = 2
    Case Else
      myPoints(0).SetAsNewPoint Me, PointX - Length, PointY, 1
      myPoints(1).SetAsNewPoint Me, PointX + Length, PointY, 2
      myNumPoints = 1
  End Select

  CopyTo Defaults(ShapeType)
End Sub

Public Property Get Caption(Optional ByVal Default As Boolean = False)
  If Name = "" Or Default Then
    Caption = Lang.GetString(ShapeType + RES_Obj_Header) & " " & ObjectIndex
  Else
    Caption = Name
  End If
End Property

Public Sub Compile(ByVal FileNum As Integer)
  Dim Pos As String, Nums As String, _
    I As Integer, J As Integer, _
    intX As Long, intY As Long, _
    Dist As Double, V1 As Long, V2 As Long, _
    UseScale As Single, Flags As String, _
    BaseColor As String, Palette As String, _
    ColorString As String, _
    MoveToStr As String, DrawToStr As String, _
    strBuffer As String, StrBuffer2 As String, _
    Range As Integer, _
    Flags2 As String

  Dim TempPts() As PointType
  Dim APIPoints() As POINTAPI

  Dim deltaX As Single, deltaY As Single, _
    NewX As Single, NewY As Single, _
    DotTestValue As Integer, numberOfDots As Single, _
    OldX As Single, OldY As Single, OldCaption As String
  Dim Pt As PointType
  
  Dim PtRes1 As PointType, PtRes2 As PointType
  
  ReDim TempPts(UBound(Pts))
  ReDim APIPoints(UBound(Pts))
  Dim Rgn As Long

  Dim MaxPoint As Single, TempTest As Single

  If ShapeType = OT_FlatArea Then
    If Extra3 = 0 Then
      ' Flatten
    ElseIf Extra3 = 1 Then
      ' Area16N
      Dim Head1 As Single, Head2 As Single
      strBuffer = "Area16N" & vbCrLf & _
                  "    Elevation( " & FloatFormat(Extra1, "####0.0#") & vbCrLf
      RectToPolar Pts(0).X, Pts(0).Y, 0, Head1
      RectToPolar Pts(1).X, Pts(1).Y, 0, Head2

      If Head1 < 0 Then Head1 = Head1 + 360
      If Head2 < 0 Then Head2 = Head2 + 360

      If Head1 < Head2 Then
        For I = 0 To UBound(Pts)
          strBuffer = strBuffer & "      " & ReturnPoint(Pts(I).X, Pts(I).Y).LatLon & vbCrLf
        Next I
      Else
        For I = UBound(Pts) To 0 Step -1
          strBuffer = strBuffer & "      " & ReturnPoint(Pts(I).X, Pts(I).Y).LatLon & vbCrLf
        Next I
      End If

      strBuffer = strBuffer & "    )" & vbCrLf & _
                  "End16" & vbCrLf & vbCrLf

      Scenery.A16NBuffer = Scenery.A16NBuffer & strBuffer
    End If
    Exit Sub
  End If

  If ShapeType = OT_Line And Not LinkObject Is Nothing Then
    OldX = LinkObject.PositionX
    OldY = LinkObject.PositionY
    OldCaption = LinkObject.Name
    LinkObject.Name = LinkObject.Caption & " (" & Caption & ")"

    For I = 0 To myNumPoints - 1
      If myPoints(I).PtType = FSNormal Then
        ' Calculate the number of dots needed to draw
        ' and then subdivide the line
        If DotSpacing > 0 Then
          numberOfDots = Distance(Pts(I).X, Pts(I).Y, Pts(I + 1).X, Pts(I + 1).Y) / DotSpacing
          If numberOfDots > 0 Then
            deltaX = (Pts(I + 1).X - Pts(I).X) / numberOfDots
            deltaY = (Pts(I + 1).Y - Pts(I).Y) / numberOfDots
            ' Add a 1/4 spacing margin
            ' Nevermind, bad idea
            NewX = Pts(I).X ' + deltaX / 4
            NewY = Pts(I).Y ' + deltaY / 4
            For J = 1 To Int(numberOfDots) + 1
              LinkObject.PositionX = NewX
              LinkObject.PositionY = NewY
              LinkObject.Compile FileNum
              NewX = NewX + deltaX
              NewY = NewY + deltaY
            Next J
          End If
        End If
      End If
    Next I

    LinkObject.PositionX = OldX
    LinkObject.PositionY = OldY
    LinkObject.Name = OldCaption

    Exit Sub
  End If

  ' Determine the average for all points for reference center
  SumX = 0
  SumY = 0
  For I = 0 To myNumPoints
    TempPts(I).X = myPoints(I).X
    TempPts(I).Y = myPoints(I).Y
    Rotate TempPts(I), Scenery.Header.Rotation
    SumX = SumX + TempPts(I).X
    SumY = SumY + TempPts(I).Y
    Nums = Nums & Str$(I)
  Next I
  AvgX = SumX / (myNumPoints + 1)
  AvgY = SumY / (myNumPoints + 1)
  intX = CLng(AvgX)
  intY = CLng(AvgY)

  ' Find out longest distance
  For I = 0 To UBound(TempPts)
    Dist = Distance(AvgX, AvgY, TempPts(I).X, TempPts(I).Y)
    TempTest = Abs(TempPts(I).X - AvgX)
    If TempTest > MaxPoint Then
      MaxPoint = TempTest
    End If

    TempTest = Abs(TempPts(I).Y - AvgY)
    If TempTest > MaxPoint Then
      MaxPoint = TempTest
    End If

    If Dist > V2 Then V2 = Dist
  Next I

  V2 = V2 * 2
  If V2 < 300 Then V2 = 300

  If ShapeType = OT_TaxiwayLine Then
    UseScale = 0.3048
    Layer = 40
  Else
    If Texture = "" Then
      UseScale = MaxPoint / 32000
      Select Case UseScale
        Case Is <= 0.01
          UseScale = 0.01
        Case Is <= 0.05
          UseScale = 0.05
        Case Is <= 0.1
          UseScale = 0.1
        Case Is < 0.5
          UseScale = 0.5
        Case Else
          UseScale = 1
      End Select
    Else
      UseScale = MScale
    End If
  End If

  If Visibility = 0 Then
    V1 = V2 * 10
    If V1 < 15000 Then V1 = 15000
  Else
    V1 = Visibility
  End If
  
  V2 = CLng(V2 / UseScale / 10) * 10 ' Scale units

  If V2 > 60000 Then V2 = 60000

  Pt = MakePoint(intX, intY)
  Rotate Pt, -Scenery.Header.Rotation
  Pos = ReturnPoint(Pt.X, Pt.Y).LatLon

  If Not LinkFromObj Is Nothing Then
    ' Polygon
    MsgBoxEx frmMain, "Sorry, linked polygons have not been implemented yet. Thank you for your patience.", vbInformation, 0
    Exit Sub
  End If

  Range = Int(V1 / 1000 + 0.99)
  If Range > 80 Then Range = 80
  
  Print #FileNum, "; " & Caption
  Print #FileNum, "Area( " & IIf(Range <= 20, "5 ", "B ") & Pos & " " & Range & " )"

  If Layer > 0 Then
    Print #FileNum, "    LayerCall( :lblLayer " & Layer & " )"
    Print #FileNum, "    Jump( : )"
    Print #FileNum, ":lblLayer"
  End If

  If Complexity > 0 Then Print #FileNum, "    IfVarRange( :lblEnd 346 " & Complexity & " 32767 )"

  Print #FileNum, "    RefPoint( rel :lblEnd " & FloatFormat(UseScale, "#0.0#######") & " " & Pos & " v1= " & V1 & " v2= " & V2 & " )"

  If Options.FSVersion >= Version_FS2K Then
    ColorToSCASM Color, BaseColor, Palette
    If Palette = "F0" Then
      ColorString = "F0 " & BaseColor & " 0 0"
    Else
      ColorString = Palette & " " & BaseColor
    End If

    If Len(Texture) > 0 Then
      Flags = Left$(Texture, InStr(Texture, " ") - 1)
      Print #FileNum, "    LoadBitmap( 0 " & Mid$(Flags, 2) & "5 " & ColorString & " " & QuoteString(GetFileTitle(Texture)) & " )"
      Print #FileNum, "    Smoothing( 1 )"
    Else
      If Color > 0 Then
        Print #FileNum, "    TextureEnable( 0 )"
        If ShapeType = OT_Line And Extra1 < 0.2 Then
          Print #FileNum, "    RGBLColor( " & ColorString & " )"
        Else
          Print #FileNum, "    RGBSColor( " & ColorString & " )"
        End If
      End If
    End If
  Else
    ColorToSCASM Color, BaseColor, Palette
    If Color > 0 Then
      If Len(Texture) = 0 Then
        Print #FileNum, "    TextureEnable( 0 )"
      End If
      If ShapeType = OT_Line And Extra1 < 0.2 Then
        Print #FileNum, "    LineColor( " & Hex$(BaseColor) & " F0 )"
      Else
        Print #FileNum, "    SurfaceColor( " & Hex$(BaseColor) & " F0 )"
        If Len(Texture) > 0 Then Print #FileNum, "    IfVarAnd( :NoTextured 340 FFFF )"
      End If
    End If
    If Len(Texture) > 0 Then
      Print #FileNum, "    Bitmap( " & GetFileTitle(Texture) & " 0 0 0 0 )"
      Print #FileNum, "    Smoothing( 1 )"
      If Color > 0 Then Print #FileNum, ":NoTextured"
    End If
  End If

  If ShapeType = OT_Polygon Then
    Print #FileNum, "    Points( 0"
    For I = 0 To UBound(TempPts)
      Print #FileNum, "            " & CLng((TempPts(I).X - intX) / UseScale) & " " & CLng((Extra1 + myPoints(I).Z) / UseScale) & " " & CLng((TempPts(I).Y - intY) / UseScale)
      ' Find if this is a concave polygon
      APIPoints(I) = MakeAPIPoint(TempPts(I).X, TempPts(I).Y)
    Next I
    Print #FileNum, "    )"
    Rgn = CreatePolygonRgn(APIPoints(0), UBound(TempPts) + 1, WINDING)

    For I = 0 To UBound(TempPts) - 2
      For J = I + 2 To UBound(TempPts)
        If I = 0 And J = UBound(TempPts) Then Exit For
        If PtInRegion(Rgn, (TempPts(I).X + TempPts(J).X) / 2, (TempPts(I).Y + TempPts(J).Y) / 2) = 0 Then
          ' Concave
          Print #FileNum, "    ConcavePoly"
          GoTo ConcaveFinished:
        End If
      Next J
    Next I
    DeleteRegion Rgn
ConcaveFinished:
    If Color And NightMask Then
      Print #FileNum, "    ShadedPoly( au" & Nums & " )"
    Else
      Print #FileNum, "    Poly( au" & Nums & " )"
    End If
  ElseIf ShapeType = OT_Line And DotSpacing > 0 Then

  ElseIf ShapeType = OT_Line And Extra1 < 0.2 Then
    If Extra2 Then Print #FileNum, "    IfVarRange( :lblEnd 28C 2 4 )"
    Print #FileNum, "    MoveTo( " & CLng((TempPts(0).X - intX) / UseScale) & " " & CLng(myPoints(0).Z / UseScale) & " " & CLng((TempPts(0).Y - intY) / UseScale) & " )"
    For I = 1 To UBound(TempPts)
      Print #FileNum, "    " & IIf(myPoints(I - 1).PtType = FSNormal, "DrawTo", "MoveTo") & "( " & CLng((TempPts(I).X - intX) / UseScale) & " " & CLng(myPoints(I).Z / UseScale) & " " & CLng((TempPts(I).Y - intY) / UseScale) & " )"
    Next I
  ElseIf ShapeType = OT_TaxiwayLine Then
    Print #FileNum, "    TaxiMarkings("
    Flags = Choose(Extra3 + 1, "SS", "DD", "SD", "C", "HILS", "HRWY", "HTAXI")
    If Extra3 <= 3 Then
      Flags2 = Choose(Extra3 + 1, "ISS", "IDD", "IDS", "IC")
      If Extra3 = 2 And Extra2 Then
        Flags = "DSL"
      ElseIf Extra2 Then
        Flags = Flags & "L"
      End If
      Print #FileNum, "    " & Flags & " " & FloatFormat((TempPts(0).X - intX) / UseScale, "####0.0#########") & " " & FloatFormat((TempPts(0).Y - intY) / UseScale, "####0.0#########")
      For I = 1 To UBound(TempPts) - 1
        If Extra1 = 0 Then
          ' No arc
          Print #FileNum, "    " & IIf(Point(I - 1).PtType = FSNormal, Flags, Flags2) & " " & FloatFormat((TempPts(I).X - intX) / UseScale, "####0.0#########") & " " & FloatFormat((TempPts(I).Y - intY) / UseScale, "####0.0#########")
        Else
          If Point(I - 1).PtType = FSNormal Then
            If FindArcPoints(Extra1, TempPts(I - 1), TempPts(I), TempPts(I + 1), PtRes1, PtRes2) Then
              Print #FileNum, "    " & Flags & " " & FloatFormat((PtRes1.X - intX) / UseScale, "####0.0#########") & " " & FloatFormat((PtRes1.Y - intY) / UseScale, "####0.0#########")
              If Extra2 Then
                Print #FileNum, "    ARCL"
              Else
                Print #FileNum, "    ARC"
              End If
              Print #FileNum, "    " & Flags & " " & FloatFormat((PtRes2.X - intX) / UseScale, "####0.0#########") & " " & FloatFormat((PtRes2.Y - intY) / UseScale, "####0.0#########")
            Else
              ' Couldn't fit an arc, just draw the point
              Print #FileNum, "    " & Flags & " " & FloatFormat((TempPts(I).X - intX) / UseScale, "####0.0#########") & " " & FloatFormat((TempPts(I).Y - intY) / UseScale, "####0.0#########")
            End If
          Else
            Print #FileNum, "    " & Flags2 & " " & FloatFormat((TempPts(I).X - intX) / UseScale, "####0.0#########") & " " & FloatFormat((TempPts(I).Y - intY) / UseScale, "####0.0#########")
            Print #FileNum, "    " & Flags2 & " " & FloatFormat((TempPts(I).X - intX) / UseScale, "####0.0#########") & " " & FloatFormat((TempPts(I).Y - intY) / UseScale, "####0.0#########")
          End If
        End If
      Next I
      I = UBound(TempPts)
      Print #FileNum, "    " & Flags & " " & FloatFormat((TempPts(I).X - intX) / UseScale, "####0.0#########") & " " & FloatFormat((TempPts(I).Y - intY) / UseScale, "####0.0#########")
    Else
      For I = 0 To UBound(TempPts) - 1
        If Point(I).PtType = FSNormal Then
          Print #FileNum, "    " & Flags & " " & FloatFormat((TempPts(I).X - intX) / UseScale, "####0.0#########") & " " & FloatFormat((TempPts(I).Y - intY) / UseScale, "####0.0#########") & " " & FloatFormat((TempPts(I + 1).X - intX) / UseScale, "####0.0#########") & " " & FloatFormat((TempPts(I + 1).Y - intY) / UseScale, "####0.0#########")
        End If
      Next I
    End If
    Print #FileNum, "    BREAK )"
  Else
    Select Case ShapeType
      Case OT_Taxiway
        If Options.FSVersion >= Version_FS2K Then
          MoveToStr = "TaxiwayStart"
          DrawToStr = "TaxiwayCont"
        Else
          MoveToStr = "TaxiMoveTo"
          DrawToStr = "TaxiLineTo"
        End If
        Flags = ""
      Case OT_Road
        Select Case Extra3
          Case 0
            If Options.FSVersion >= Version_FS2K Then
              MoveToStr = "RRStart"
              DrawToStr = "RoadCont"
              Flags = IIf(Texture = "", "0 ", "1 ")
            Else
              MoveToStr = "RoadMoveTo"
              DrawToStr = "RoadLineTo"
            End If
          Case 1
            MoveToStr = "RRStart"
            DrawToStr = "RoadCont"
            Flags = "2 "
          Case 2
            MoveToStr = "RRStart"
            DrawToStr = "RoadCont"
            Flags = "3 "
          Case 3
            MoveToStr = "RRStart"
            DrawToStr = "RoadCont"
            Flags = "4 "
        End Select
      Case OT_River
        If Options.FSVersion >= Version_FS2K Then
          MoveToStr = "RiverStart"
          DrawToStr = "RiverCont"
        Else
          MoveToStr = "RoadMoveTo2"
          DrawToStr = "RoadLineTo2"
        End If
        Flags = ""
      Case OT_Line
        If Extra2 Then Print #FileNum, "    IfVarRange( :lblEnd 28C 2 4 )"
        If Options.FSVersion >= Version_FS2K Then
          MoveToStr = "RRStart"
          DrawToStr = "RoadCont"
          Flags = IIf(Texture = "", "0 ", "1 ")
        Else
          MoveToStr = "RoadMoveTo"
          DrawToStr = "RoadLineTo"
        End If
    End Select

    Print #FileNum, "    " & MoveToStr & "( " & Flags & CLng(Extra1 / 2 / UseScale * IIf(Extra2 And ShapeType <> OT_Line, 1, -1)) & " " & CLng((TempPts(0).X - intX) / UseScale) & " " & CLng(myPoints(0).Z / UseScale) & " " & CLng((TempPts(0).Y - intY) / UseScale) & " )"
    For I = 1 To UBound(TempPts)
      If myPoints(I - 1).PtType = FSNormal Then
        Print #FileNum, "    " & DrawToStr & "( " & CLng((TempPts(I).X - intX) / UseScale) & " " & CLng(myPoints(I).Z / UseScale) & " " & CLng((TempPts(I).Y - intY) / UseScale) & " )"
      Else
        Print #FileNum, "    " & MoveToStr & "( " & Flags & CLng(Extra1 / 2 / UseScale * IIf(Extra2 And ShapeType <> OT_Line, 1, -1)) & " " & CLng((TempPts(I).X - intX) / UseScale) & " " & CLng(myPoints(I).Z / UseScale) & " " & CLng((TempPts(I).Y - intY) / UseScale) & " )"
      End If
    Next I
    If ShapeType = OT_Taxiway And Options.FSVersion < Version_FS2K Then
      Print #FileNum, "    StartSurface"
      Print #FileNum, "    MoveTo( 0 0 0 )"
      Print #FileNum, "    EndSurface"
    End If
  End If
  If Len(Texture) > 0 Then Print #FileNum, "    Smoothing( 0 )"

  V2 = V2 * UseScale

  UseScale = MaxPoint / 32000
  Select Case UseScale
    Case Is <= 0.01
      UseScale = 0.01
    Case Is <= 0.05
      UseScale = 0.05
    Case Is <= 0.1
      UseScale = 0.1
    Case Is < 0.5
      UseScale = 0.5
    Case Else
      UseScale = 1
  End Select

  V2 = CLng(V2 / UseScale / 10) * 10 ' Scale units

  If DotSpacing > 0 And DotColor > 0 Then
    strBuffer = ""
    Select Case ShapeType
      Case OT_Polygon
        DotTestValue = FSMoveTaxi
        If myPoints(myNumPoints).PtType = DotTestValue Then
          strBuffer = strBuffer & "    DotLine( " & CLng((TempPts(myNumPoints).X - intX) / UseScale) & " " & CLng(myPoints(myNumPoints).Z / UseScale + 1) & " " & CLng((TempPts(myNumPoints).Y - intY) / UseScale) & " " & CLng((TempPts(0).X - intX) / UseScale) & " " & CLng(myPoints(0).Z / UseScale + 1) & " " & CLng((TempPts(0).Y - intY) / UseScale) & " " & CLng(Distance(TempPts(myNumPoints).X, TempPts(myNumPoints).Y, TempPts(0).X, TempPts(0).Y) / DotSpacing + 1) & " )" & vbCrLf
        End If
      Case OT_Taxiway
        DotTestValue = FSNormal
      Case OT_Line
        DotTestValue = FSNormal
    End Select

    For I = 0 To myNumPoints - 1
      If myPoints(I).PtType = DotTestValue Then
        strBuffer = strBuffer & "    DotLine( " & CLng((TempPts(I).X - intX) / UseScale) & " " & CLng(myPoints(I).Z / UseScale + 1) & " " & CLng((TempPts(I).Y - intY) / UseScale) & " " & CLng((TempPts(I + 1).X - intX) / UseScale) & " " & CLng(myPoints(I + 1).Z / UseScale + 1) & " " & CLng((TempPts(I + 1).Y - intY) / UseScale) & " " & CLng(Distance(TempPts(I).X, TempPts(I).Y, TempPts(I + 1).X, TempPts(I + 1).Y) / DotSpacing + 1) & " )" & vbCrLf
      End If
    Next I

    If strBuffer <> "" Then
      Print #FileNum, vbCrLf & ";" & String$(20, "-") & vbCrLf & "; " & Caption & " (Dots)" & vbCrLf
      If ShapeType <> OT_Line Or Extra2 Then Print #FileNum, "    IfVarRange( :lblEnd 28C 2 4 )"
      Print #FileNum, "    RefPoint( rel :lblEnd " & FloatFormat(UseScale, "#0.0#######") & " " & Pos & " v1= " & V1 & " v2= " & V2 & " )"

      If Options.FSVersion >= Version_FS2K Then
        ColorToSCASM DotColor, BaseColor, Palette
        If Palette = "F0" Then
          ColorString = "F0 " & BaseColor & " 0 0"
        Else
          ColorString = Palette & " " & BaseColor
        End If
        Print #FileNum, "    RGBLColor( " & ColorString & " )"
      Else
        ColorToSCASM DotColor, BaseColor, Palette
        Print #FileNum, "    LineColor( " & Hex$(BaseColor) & " F0 )"
      End If
      Print #FileNum, strBuffer
    End If
  End If

  Print #FileNum, ":lblEnd"
  If Layer > 0 Then
    Print #FileNum, "    Return"
  End If
  Print #FileNum, "EndA" & vbCrLf


  If LineWidth > 0 And LineColor > 0 Then
    Print #FileNum, "; " & Caption & " (Lines)" & vbCrLf
    Print #FileNum, "Area( " & IIf(Range <= 20, "5 ", "B ") & Pos & " " & Range & " )"
  
    Print #FileNum, "    LayerCall( :lblLayer " & (Layer + 4) & " )"
    Print #FileNum, "    Jump( : )"
    Print #FileNum, ":lblLayer"
  
    If Complexity > 0 Then Print #FileNum, "    IfVarRange( :lblEnd 346 " & Complexity & " 32767 )"

    Print #FileNum, "    RefPoint( rel :lblEnd " & FloatFormat(UseScale, "#0.0#######") & " " & Pos & " v1= " & V1 & " v2= " & V2 & " )"

    If LineWidth < 0.2 Then
      ColorToSCASM LineColor, BaseColor, Palette
      If Options.FSVersion >= Version_FS2K Then
        If Palette = "F0" Then
          ColorString = "F0 " & BaseColor & " 0 0"
        Else
          ColorString = Palette & " " & Hex$(BaseColor)
        End If
        Print #FileNum, "    RGBLColor( " & ColorString & " )"
      Else
        Print #FileNum, "    LineColor( " & Hex$(BaseColor) & " F0 )"
      End If

      Print #FileNum, "    MoveTo( " & CLng((TempPts(0).X - intX) / UseScale) & " " & CLng(myPoints(0).Z / UseScale) & " " & CLng((TempPts(0).Y - intY) / UseScale) & " )"
      For I = 1 To UBound(TempPts)
        Print #FileNum, "    " & IIf(myPoints(I - 1).PtType = FSNormal, "DrawTo", "MoveTo") & "( " & CLng((TempPts(I).X - intX) / UseScale) & " " & CLng(myPoints(I).Z / UseScale) & " " & CLng((TempPts(I).Y - intY) / UseScale) & " )"
      Next I
    Else
      ColorToSCASM LineColor, BaseColor, Palette
      If Options.FSVersion >= Version_FS2K Then
        If Palette = "F0" Then
          ColorString = "F0 " & BaseColor & " 0 0"
        Else
          ColorString = Palette & " " & BaseColor
        End If
        Print #FileNum, "    RGBSColor( " & ColorString & " )"

        MoveToStr = "RRStart"
        DrawToStr = "RoadCont"
        Flags = "0 "
      Else
        Print #FileNum, "    SurfaceColor( " & Hex$(BaseColor) & " F0 )"
        MoveToStr = "RoadMoveTo"
        DrawToStr = "RoadLineTo"
        Flags = ""
      End If

      Print #FileNum, "    TextureEnable( 0 )"
      Print #FileNum, "    " & MoveToStr & "( " & Flags & CLng(LineWidth / 2 / UseScale * -1) & " " & CLng((TempPts(0).X - intX) / UseScale) & " " & CLng(myPoints(0).Z / UseScale) & " " & CLng((TempPts(0).Y - intY) / UseScale) & " )"
      For I = 1 To UBound(TempPts)
        If myPoints(I - 1).PtType = FSNormal Then
          Print #FileNum, "    " & DrawToStr & "( " & CLng((TempPts(I).X - intX) / UseScale) & " " & CLng(myPoints(I).Z / UseScale) & " " & CLng((TempPts(I).Y - intY) / UseScale) & " )"
        Else
          Print #FileNum, "    " & MoveToStr & "( " & Flags & CLng(LineWidth / 2 / UseScale * -1) & " " & CLng((TempPts(I).X - intX) / UseScale) & " " & CLng(myPoints(I).Z / UseScale) & " " & CLng((TempPts(I).Y - intY) / UseScale) & " )"
        End If
      Next I
    End If
    Print #FileNum, ":lblEnd"
    Print #FileNum, "    Return"
    Print #FileNum, "EndA" & vbCrLf
  End If
End Sub

Public Sub CopyTo(Destination As clsObject)
  Dim Temp As clsShape
  
  If TypeOf Destination Is clsShape Then
    Set Temp = Destination
    With Temp
      .Name = Name
      .Locked = Locked
      .Complexity = Complexity
      .Texture = Texture
      .Color = Color
      .Layer = Layer
      .MScale = MScale
      .DotSpacing = DotSpacing
      .DotColor = DotColor
      .LineWidth = LineWidth
      .LineColor = LineColor
      .Extra1 = Extra1
      .Extra2 = Extra2
      .Extra3 = Extra3
      .ShapeType = ShapeType
      .Visibility = Visibility
    End With
  End If
End Sub

Public Sub DeletePoints()
  Dim I As Integer, J As Integer, K As Integer, _
    Min As Integer, Index As Integer
  Dim NewPoints() As clsPoint
  
  ReDim NewPoints(myNumPoints)
  
  Min = Scenery.Count
  For I = 0 To myNumPoints
    If Not Scenery.Focus(myPoints(I).SceneryIndex) Then
      With myPoints(I)
        'If .SceneryIndex < Min Then
        '  Min = .SceneryIndex
        '  Index = J
        'End If
        Set NewPoints(J) = myPoints(I)
        J = J + 1
      End With
    Else
      Set myPoints(I).Parent = Nothing
    End If
    Set myPoints(I) = Nothing
  Next I
  
  NumPoints = J - 1

  For I = 0 To J - 1
    Set myPoints(I) = NewPoints(I)
    myPoints(I).ObjectIndex = I + 1
    Set NewPoints(I) = Nothing
  Next I
End Sub

Public Sub DrawBottom(Optional ByVal DrawX As Single = -999999#, Optional ByVal DrawY As Single)
  Dim myColor As Long, I As Integer, _
    Width As Integer, Height As Integer
  
  Static PreviousColor As Long

  If Not mDrawable Or Not Options.ObjectVisible(ShapeType) Then Exit Sub

  glPushMatrix
  If DrawX = -999999 Then
    ' No drag
    myColor = ShapeType
    If myColor <> PreviousColor Then
      UpdatePoints myColor
    ElseIf myNumPoints <> UBound(Pts) Then
      UpdatePoints myColor
    End If
  Else
    ' Drag
    myColor = COLOR_DragColor
    UpdatePoints myColor
  End If
  
  PreviousColor = myColor
  
  Options.SetGLColor myColor
 
  If ShapeType = OT_FlatArea Or (ShapeType = OT_Polygon And (Not Options.FillPolygons Or myColor = COLOR_DragColor Or (Color = 0 And glBitmapID = -1))) Then
    ' Basic 1-pixel width lines (closed figure)
  
    ' Nothing here
  ElseIf ShapeType = OT_Polygon Then ' (and Options.FillObject)
    ' Filled polygon
    If glBitmapID > -1 Then
      picEditor.SelectBitmap glBitmapID
      picEditor.GetBitmapDimensions glBitmapID, Width, Height
      glDrawTesselatedTextured TessPts, 1 / Width / MScale, AvgX, 1 / Height / MScale, AvgY
    ElseIf Color > 0 Then
      glForeColor ExtractColor(Color)
      glDrawTesselated TessPts
    End If
  ElseIf (ShapeType = OT_Line And (DotSpacing > 0 Or Extra1 < 0.2)) Or _
          myColor = COLOR_DragColor Or Not Options.ThickLines Or (Color = 0 And glBitmapID = -1) Then
    ' 1-pixel width lines and dashed lines
    
    ' Nothing here
  Else
    ' Thick lines
    If Options.FSVersion <= Version_CFS1 Then
      If glBitmapID > -1 Then
        picEditor.SelectBitmap glBitmapID
        picEditor.GetBitmapDimensions glBitmapID, Width, Height
        glDrawTesselatedTextured TessPts, 1 / Width / MScale, AvgX, 1 / Height / MScale, AvgY
      Else
        glForeColor ExtractColor(Color)
        glDrawTesselated TessPts
      End If
    Else
      If glBitmapID > -1 Then
        picEditor.SelectBitmap glBitmapID
        picEditor.GetBitmapDimensions glBitmapID, Width, Height
        For I = 0 To UBound(TessPts)
          glPaintTexturedRegion2 TessPts(I).Pts, TexCoord(I).Pts
        Next I
      Else
        glForeColor ExtractColor(Color)
        glDrawTesselated TessPts
      End If
    End If
    ' Stuff here
  End If
 
  glPopMatrix
End Sub

Public Sub DrawTop(Optional ByVal DrawX As Single = -999999#, Optional ByVal DrawY As Single)
  Dim myColor As Long, _
    I As Integer, J As Integer, _
    K As Integer, GoUpTo As Integer, _
    deltaX As Single, deltaY As Single, _
    NewX As Single, NewY As Single, _
    DotTestValue As Integer, numberOfDots As Single

  If Not mDrawable Or Not Options.ObjectVisible(ShapeType) Then Exit Sub

  glPushMatrix
  If DrawX = -999999 Then
    ' No drag
    myColor = ShapeType
  Else
    ' Drag
    myColor = COLOR_DragColor
  End If
  
  Options.SetGLColor myColor
  
  DotTestValue = -1
 
  If ShapeType = OT_FlatArea Or (ShapeType = OT_Polygon And (Not Options.FillPolygons Or myColor = COLOR_DragColor Or (Color = 0 And glBitmapID = -1))) Then
    ' Basic 1-pixel width lines (closed figure)
    glDrawPolygon Pts
    If ShapeType = OT_Polygon Then DotTestValue = FSMoveTaxi
  ElseIf ShapeType = OT_Polygon Then ' (and Options.FillObject)
    ' Filled polygon
    DotTestValue = FSMoveTaxi
  ElseIf (ShapeType = OT_Line And (DotSpacing > 0 Or Extra1 < 0.2)) Or _
          myColor = COLOR_DragColor Or Not Options.ThickLines Or (Color = 0 And glBitmapID = -1) Then
    ' 1-pixel width lines and dashed lines
    glBeginLines
    For I = 0 To myNumPoints - 1
      If myPoints(I).PtType = FSNormal Then
        glDrawLine Pts(I).X, Pts(I).Y, Pts(I + 1).X, Pts(I + 1).Y
      End If
    Next I
    glEnd
    
    glDashLine
    glBeginLines
    For I = 0 To myNumPoints - 1
      If myPoints(I).PtType = FSMoveTaxi Then
        glDrawLine Pts(I).X, Pts(I).Y, Pts(I + 1).X, Pts(I + 1).Y
      End If
    Next I
    glEnd
    glSmoothLine
    
    If ShapeType = OT_Line Then DotTestValue = FSNormal
  Else
    ' Thick lines

    If ShapeType = OT_Line Then DotTestValue = FSNormal
  End If
 
  If DotTestValue > -1 And DotSpacing > 0 And myColor <> COLOR_DragColor Then
    ' We need to draw dots
    Select Case ShapeType
      Case OT_Polygon, OT_FlatArea
        GoUpTo = myNumPoints
      Case Else
        GoUpTo = myNumPoints - 1
    End Select
    
    glPointSize 3
    glBeginPoints
    glForeColor ExtractColor(DotColor)
    
    For I = 0 To GoUpTo
      J = (I + 1) Mod (myNumPoints + 1)
      If myPoints(I).PtType = DotTestValue Then
        ' Calculate the number of dots needed to draw
        ' and then subdivide the line
        numberOfDots = Distance(Pts(I).X, Pts(I).Y, Pts(J).X, Pts(J).Y) / DotSpacing
        If numberOfDots > 0 Then
          deltaX = (Pts(J).X - Pts(I).X) / numberOfDots
          deltaY = (Pts(J).Y - Pts(I).Y) / numberOfDots
          ' Add a 1/4 spacing margin
          NewX = Pts(I).X ' + deltaX / 4
          NewY = Pts(I).Y ' + deltaY / 4
          For K = 1 To Int(numberOfDots) + 1
            glDrawPoint NewX, NewY
            NewX = NewX + deltaX
            NewY = NewY + deltaY
          Next K
        End If
      End If
    Next I
    glEnd
  End If
 
  glPopMatrix
End Sub

Public Sub FillDefaults()
  If Defaults(ShapeType) Is Nothing Or Not Options.Remember Then
    ' Default values:
    MScale = 1
    Select Case ShapeType
      Case OT_Polygon
        DotColor = &H50000011
        DotSpacing = 60
      Case OT_Taxiway
        Texture = "N " & AddDir(Options.FSPath, "Texture\concrete.r8")
        Color = &H10000003
        DotColor = &H50000014
        LineColor = &H1000002D
        Extra1 = 30
        Extra2 = True
        Layer = 16
      Case OT_Road
        Texture = "N " & AddDir(Options.FSPath, "Texture\asphalt.r8")
        Color = &H10000002
        Extra1 = 26
        Extra2 = True
        Layer = 16
      Case OT_River
        If Options.FSVersion >= Version_FS2K Then
          Texture = "N " & AddDir(Options.FSPath, "Texture\256b2su1.bmp")
          Color = &H2FB4AD8A
        Else
          Texture = "N " & AddDir(Options.FSPath, "Texture\water.r8")
          Color = &H1000000E
        End If
        Extra1 = 60
        Extra2 = True
        Layer = 12
      Case OT_Line
        Extra1 = 0.1
        Layer = 20
      Case OT_FlatArea
        Extra1 = Scenery.Header.Altitude
        Extra3 = 1
      Case OT_TaxiwayLine
        Layer = 40
    End Select
  Else
    Defaults(ShapeType).CopyTo Me
  End If
End Sub

' Inserts a new point at X, Y, position Index
Private Sub InsertNewPoint(ByVal Index As Integer, ByVal X As Single, ByVal Y As Single)
  Dim I As Integer

  NumPoints = NumPoints + 1
  
  For I = NumPoints - 1 To Index + 1 Step -1
    Set myPoints(I + 1) = myPoints(I)
    myPoints(I + 1).ObjectIndex = (I + 1) + 1
  Next I
  
  ' Insert new point
  Set myPoints(Index + 1) = New clsPoint
  myPoints(Index + 1).SetAsNewPoint Me, X, Y, (Index + 1) + 1
  myPoints(Index + 1).PtType = myPoints(Index).PtType

  frmMain.UpdateObjectList
  Scenery.SetSingleFocus myPoints(Index + 1).SceneryIndex
  DrawBottom 0, 0
  DrawTop 0, 0
End Sub

Public Sub LoadAPTData(FileData As clsAPTData)
  Dim Temp As String, I As Integer, longTemp As Long, bReadColor As Boolean, bReadTexture As Boolean
  
  With FileData
    Select Case ShapeType
      Case OT_Polygon
        Color = ValEx(Format$("&H" & ReadNext(.GetData("Couleur"), " "))) Or PalMask
        Temp = .GetData("Texture")
        Texture = ReadNext(Temp, " ")
        If Temp <> "" Then
          MScale = ValEx(Temp)
          If InStr(Texture, "\") = 0 Then Texture = AddDir(Options.TexturePath, Texture)
          Texture = GetRealName(Texture)
        Else
          MScale = 1
        End If
        Visibility = Meter(.GetData("Range"))
        Temp = .GetData("CLTaxi")
        ReadNext Temp, " "
        If Temp <> "" Then DotColor = ValEx(Format$("&H" & ReadNext(Temp, " "))) Or PalMask
        Extra1 = Meter(.GetData("GNDAltitude"))
        ' Surface type
        Layer = ValEx(.GetData("Layer"))
        Complexity = .GetComplexityData()
        Name = .GetData("Comment")
        NumPoints = .GetData("Max") - 1
        DotSpacing = Meter(.GetData("Spacing"))
        
        Temp = .GetData("SeasonTexture")
        If Temp <> "" Then
          Texture = "N" & Mid$(ReadNext(Temp, " "), 2) & " " & Texture
          Color = RGBMask Or RGB(Format$("&H" & ReadNext(Temp, " ")), Format$("&H" & ReadNext(Temp, " ")), Format$("&H" & ReadNext(Temp, " "))) Or &H1000000 * Format$("&H" & Mid$(ReadNext(Temp, " "), 2, 1))
        ElseIf Texture <> "" Then
          Texture = "N " & Texture
        End If
      Case OT_Line
        Color = ValEx(Format$("&H" & ReadNext(.GetData("Couleur"), " "))) Or PalMask
        Extra1 = Meter(.GetData("Largeur"))
        If Extra1 = 0 Then Extra1 = 0.1
        Temp = .GetData("Poly3D")
        If Temp <> "" Then
        Else
          Temp = .GetData("Espacement")
          Temp = ReadNext(Temp, " ") & " " & ReadNext(Temp, " ")
          DotSpacing = Meter(Temp)
          If DotSpacing > 0 Then
            DotColor = Color
            Color = 0
          End If
        End If
        Visibility = Meter(.GetData("Range"))
        Layer = ValEx(.GetData("Layer"))
        Complexity = .GetComplexityData()
        Name = .GetData("Comment")
        NumPoints = .GetData("Max") - 1
        MScale = 1
      Case OT_FlatArea
        Extra1 = Meter(.GetData("Altitude")) + Scenery.Header.Altitude
        Extra3 = 1
        Name = .GetData("Comment")
        NumPoints = .GetData("Max") - 1
      Case Else
        Temp = .GetData("RoadTaxi")
        Select Case LCase$(Temp)
          Case "road"
            ShapeType = OT_Road
          Case "taxi"
            ShapeType = OT_Taxiway
        End Select
        Extra1 = Meter(.GetData("Largeur"))
        If Extra1 = 0 Then
          ShapeType = OT_Line
          Extra1 = 0.1
        ElseIf Extra1 < 0 Then
          Extra1 = -Extra1
        Else
          Extra2 = True
        End If
        Temp = .GetData("Type")
        Select Case LCase$(Temp)
          Case "route", "routenuit"
            ' Routenuite supported in FSSC, just do a regular road
            Texture = "N " & AddDir(Options.FSPath, "Texture\asphalt.r8")
          Case "fleuve", "fs2kriver"
            ShapeType = OT_River
          
            If Options.FSVersion >= Version_FS2K Then
              Texture = "N " & AddDir(Options.FSPath, "Texture\256b2su1.bmp")
              Color = &H2FB4AD8A
            Else
              Texture = "N " & AddDir(Options.FSPath, "Texture\water.r8")
              Color = &H1000000E
            End If
          Case "fs2kmajor"
            Extra3 = 1
            Texture = "N " & AddDir(Options.FSPath, "Texture\v_road_major.bmp")
          Case "fs2kminor"
            Extra3 = 2
            Texture = "N " & AddDir(Options.FSPath, "Texture\v_road_minor.bmp")
          Case "fs2krail"
            Extra3 = 3
            Texture = "N " & AddDir(Options.FSPath, "Texture\v_railroad.bmp")
          Case "color"
            bReadColor = True
          Case "texture"
            bReadTexture = True
        End Select
        Temp = .GetData("RouteColor")
        ReadNext Temp, " "
        ReadNext Temp, " "
        Temp = ReadNext(Temp, " ")
        If bReadColor Then
          Color = ValEx(Format$("&H" & Temp)) Or PalMask
        End If
        MScale = ValEx(.GetData("Resolution"))
        If bReadTexture Then
          Texture = .GetData("Texture")
          If Texture = "" Then
            MScale = 1
          Else
            If InStr(Texture, "\") = 0 Then Texture = AddDir(Options.TexturePath, Texture)
          End If
          Texture = GetRealName(Texture)
        
          Temp = .GetData("SeasonTexture")
          If Temp <> "" Then
            Texture = "N" & Mid$(ReadNext(Temp, " "), 2) & " " & Texture
            Color = RGBMask Or RGB(Format$("&H" & ReadNext(Temp, " ")), Format$("&H" & ReadNext(Temp, " ")), Format$("&H" & ReadNext(Temp, " "))) Or &H1000000 * Format$("&H" & Right$(ReadNext(Temp, " "), 1))
          ElseIf Texture <> "" Then
            Texture = "N " & Texture
          End If
        End If
  
        If ShapeType = OT_Taxiway Then
          DotSpacing = Meter(.GetData("Spacing"))
          Temp = .GetData("CenterColor")
          If ReadNext(Temp, " ") = "0" Then
            DotSpacing = 0
          Else
            DotColor = ValEx(Format$("&H" & ReadNext(Temp, " "))) Or PalMask
          End If
        
          Temp = .GetData("CenterStripe")
          If ReadNext(Temp, " ") <> "0" Then
            LineWidth = 0.2
            LineColor = ValEx(Format$("&H" & ReadNext(Temp, " "))) Or PalMask
          End If
        End If
        
        Visibility = Meter(.GetData("Range"))
      
        Layer = Val(.GetData("Layer"))
        Complexity = .GetComplexityData
        Name = .GetData("Comment")
        NumPoints = .GetData("Max") - 1
    End Select
    
    If MScale = 0 Then MScale = 1

    ReDim myPoints(NumPoints)
    For I = 0 To NumPoints
      Set myPoints(I) = New clsPoint
      Set myPoints(I).Parent = Me
      myPoints(I).LoadAPTData FileData
    Next I
    
    ' Airport stores taxiway vertices with the endpoint, not the start point
    ' shift taxiway lighting property over by 1
    longTemp = myPoints(0).PtType
    For I = 1 To NumPoints
      myPoints(I - 1).PtType = myPoints(I).PtType
      If myPoints(I).PtType = FSMoveTaxi And ShapeType = OT_Polygon And DotSpacing = 0 Then
        DotSpacing = 20
      End If
    Next I
    myPoints(NumPoints).PtType = longTemp
  End With
  
  mDrawable = True
  UpdateObject
End Sub

Public Sub LoadBinaryData(ByVal FileNum As Integer, ByVal Version As Integer)
  Dim ShapeTypeB As Byte, _
    DummyX As Single, DummyY As Single, _
    Count As Integer, Width As Byte, Altitude As Single
  Dim Color1 As String, Color2 As String, _
    Ext As String, Base As String, NightBool As Boolean, _
    byteDotSpacing As Byte
  Dim myDirs As String, myFile As String
    
  Select Case Version
    Case Is >= 11
      GetBinaryData FileNum, Name, DummyX, DummyY, _
        ShapeTypeB, Texture, Color1, Color2, Complexity, _
        Count, Layer, Width, byteDotSpacing, NightBool, _
        MScale, Extra3, Altitude, Locked
    Case Is >= 8
      GetBinaryData FileNum, Name, DummyX, DummyY, _
        ShapeTypeB, Texture, Color1, Color2, Complexity, _
        Count, Layer, Width, byteDotSpacing, NightBool, _
        MScale, Extra3, Altitude
    Case Is >= 3
      GetBinaryData FileNum, DummyX, DummyY, ShapeTypeB, _
        Texture, Color1, Color2, Complexity, Count, _
        Layer, Width, byteDotSpacing, NightBool, MScale, Extra3

      If Version >= 6 Then
        GetBinaryData FileNum, Altitude
      End If
      Altitude = 0
'    Case Else
'      GetData FileNum, Pos, ShapeTypeB, OldStr, ComplexityB, Count, Layer, Width
'      If Version > 1 Then
'        GetData FileNum, Dot
'      Else
'        Dot = 60
'      End If
'      If Layer = 23 Then Layer = 20
'      If Left$(OldStr, 6) = "File: " Then
'        File = Mid$(OldStr, 7)
'      Else
'        Color1 = ColorHexFormat(Mid$(OldStr, 8)) & " F0"
'      End If
'      If ShapeTypeB = OT_Polygon Then
'        Color2 = "11 F0"
'      ElseIf ShapeTypeB = OT_Line Then
'        Color2 = Color1
'      Else
'        NightBool = True
'      End If
'      Scl = 1
  End Select
  
  If Color1 = "" Then
    Color = 0
  ElseIf Right$(Color1, 2) = "F0" Then
    Color = PalMask Or Val(Format$("&H" & Left$(Color1, 2)))
  Else
    Color = RGBMask Or RGB(Format$("&H" & Left$(Color1, 2)), Format$("&H" & Mid$(Color1, 4, 2)), Format$("&H" & Mid$(Color1, 7, 2))) Or &H1000000 * Format$("&H" & Right$(Color1, 1))
  End If
  
  If Color2 = "" Then
    DotColor = 0
  ElseIf Right$(Color2, 2) = "F0" Then
    DotColor = PalMask Or Format$("&H" & Left$(Color2, 2))
  Else
    DotColor = RGBMask Or RGB(Format$("&H" & Left$(Color2, 2)), Format$("&H" & Mid$(Color2, 4, 2)), Format$("&H" & Mid$(Color2, 7, 2))) Or &H1000000 * Format$("&H" & Right$(Color2, 1))
  End If
  
  DotSpacing = byteDotSpacing
'  ShapeType = ShapeTypeB
  Select Case ShapeType
    Case OT_Polygon
    
    Case OT_Taxiway
      If Extra3 = 1 Then
        LineWidth = DotSpacing
        LineColor = DotColor
        If (LineWidth = 0 Or LineWidth = 60) Then LineWidth = 1
      End If
      DotSpacing = 0
      DotColor = 0
    
      If Version <= 6 Then
        If Layer >= 16 Then Layer = 0
      End If
      Extra1 = Width
      Extra2 = NightBool
      Extra3 = 0
    Case OT_Road
      Extra1 = Width
      Extra2 = NightBool
    Case OT_River
      Extra1 = Width
      Extra2 = NightBool
    Case OT_Line
      Extra1 = Width
      If Extra1 = 0 Then Extra1 = 0.1
      Extra2 = NightBool

      ' Not supported
      DotSpacing = 0
    Case OT_FlatArea
      Extra1 = Altitude
  End Select

  If Left$(Texture, 3) = "[T]" Then Texture = Options.FSPath & Mid$(Texture, 4)
  
  If Texture <> "" Then
    If Not FileExists(Texture) Then
      myFile = GetFileTitle(Texture)
      myDirs = AddDir(Options.FSPath, "Texture") & ";" & Options.TexturePath
      myFile = MultiDir(myFile, myDirs)
      If myFile <> "" Then Texture = GetRealName(myFile)
    End If
    
    Ext = Mid$(Texture, InStrRev(Texture, "."))
    Base = Left$(Texture, Len(Texture) - Len(Ext))
    Texture = " " & Texture
    
    If FileExists(Base & "_HW" & Ext) Then Texture = "H" & Texture
    If FileExists(Base & "_WI" & Ext) Then Texture = "W" & Texture
    If FileExists(Base & "_FA" & Ext) Then Texture = "F" & Texture
    If FileExists(Base & "_SP" & Ext) Then Texture = "S" & Texture
    If FileExists(Base & "_LM" & Ext) Then Texture = "L" & Texture
    Texture = "N" & Texture
  End If

  NumPoints = Count
End Sub

Public Sub LoadData(FileData As clsFileData, ByVal Version As Integer)
  Dim myDirs As String, myFile As String, _
    Y As Long, FileStr As String, Link As Integer
  
  With FileData
    .GetData "Name", Name, ""
    .GetData "Locked", Locked, False
    '.GetData "Type", ShapeType
    If ShapeType <> OT_FlatArea And ShapeType <> OT_TaxiwayLine Then
      .GetData "Complexity", Complexity, 0
      .GetData "Texture", Texture, ""
      .GetData "Color", Color, 0
      .GetData "Layer", Layer, 0
      .GetData "Scale", MScale, 1
      .GetData "DotSpacing", DotSpacing, 60
      .GetData "DotColor", DotColor, 0
      .GetData "LineWidth", LineWidth, 0
      .GetData "LineColor", LineColor, 0
      .GetData "LinkFrom", LinkFrom, 0
      .GetData "LinkTo", LinkTo, 0
      .GetData "Link", Link, 0
      .GetData "V1", Visibility, 0
    End If
    .GetData "Extra1", Extra1, 0
    .GetData "Extra2", Extra2, False
    .GetData "Extra3", Extra3, 0
  End With
  
  If Link > 0 And Version >= 19 Then
    Set LinkObject = Scenery.CreateNewInstance(Link)
    FileData.ReadSection
    LinkObject.LoadData FileData, Version
  End If

  If Texture <> "" Then
    Y = InStr(Texture, " ")
    FileStr = Mid$(Texture, Y + 1)
    If Not FileExists(FileStr) Then
      myFile = GetFileTitle(FileStr)
      myDirs = AddDir(Options.FSPath, "Texture") & ";" & Options.TexturePath
      myFile = MultiDir(myFile, myDirs)
      If myFile <> "" Then Texture = Left$(Texture, Y) & GetRealName(myFile)
    End If
  End If

  mDrawable = True
End Sub

' Inserts a new point at X, Y, position Index
Public Sub MidPointInsert(ByVal Index As Integer)
  Dim Index2 As Integer, X As Single, Y As Single
  
  If myNumPoints < PointUpperLimit() Then
    Index2 = (Index + 1) Mod (myNumPoints + 1)
    X = (myPoints(Index).X + myPoints(Index2).X) / 2
    Y = (myPoints(Index).Y + myPoints(Index2).Y) / 2
    InsertNewPoint Index, X, Y
  End If
End Sub

Public Property Get NumPoints() As Integer
  NumPoints = myNumPoints
End Property

Public Property Let NumPoints(ByVal Count As Integer)
  Dim X As Integer
  X = UBound(myPoints)
  If Count > X Or Count * 4 < X Then ReDim Preserve myPoints(Count * 2)
  myNumPoints = Count
End Property

Public Property Get Point(ByVal Index As Integer) As clsPoint
  Set Point = myPoints(Index)
End Property

Public Property Set Point(ByVal Index As Integer, ByVal NewPoint As clsPoint)
  Set myPoints(Index) = NewPoint
End Property

Private Function PointUpperLimit() As Integer
  If Locked Then
    PointUpperLimit = 0
  ElseIf ShapeType = OT_FlatArea Then
    If Extra3 = 0 Then
      PointUpperLimit = 21
    Else
      PointUpperLimit = 1999
    End If
  Else
    PointUpperLimit = IIf(Options.FSVersion >= Version_FS2K, 1999, 79)
  End If
End Function

' Returns true if Test is between M1 and M2
Private Function PtBetween(ByVal M1 As Single, ByVal M2 As Single, ByVal Test As Single)
  Dim Temp As Single
  If M1 > M2 Then
    Temp = M1
    M1 = M2
    M2 = Temp
  End If
  Temp = picEditor.ScaleX(Point_Threshold)
  M1 = M1 - Temp
  M2 = M2 + Temp
  PtBetween = (M1 < Test) And (Test < M2)
End Function

Public Function PrepareLink()
  Dim ShapeObject As clsShape
  If ShapeType = OT_Polygon Then
    If Not LinkObject Is Nothing Then
      LinkTo = Scenery.LinkCounter
      Scenery.LinkCounter = LinkTo + 1
      If LinkObject.ObjectType = OT_Polygon Then
        Set ShapeObject = LinkObject
        ShapeObject.LinkFrom = LinkTo
        Set ShapeObject = Nothing
      End If
    End If
  Else
    LinkTo = 0
  End If
End Function

Public Sub SaveData(FileData As clsFileData)
  Dim I As Integer
  
  With FileData
    .InitializeForWrite ObjectNames(ShapeType), Caption
    .SaveData "Name", Name, ""
    .SaveData "Locked", Locked, False
    '.SaveData "Type", ShapeType
    If ShapeType <> OT_FlatArea And ShapeType <> OT_TaxiwayLine Then
      .SaveData "Complexity", Complexity, 0
      .SaveData "Texture", Texture, ""
      .SaveData "Color", Color, 0
      .SaveData "Layer", Layer, 0
      .SaveData "Scale", MScale, 1
      .SaveData "DotSpacing", DotSpacing, 60
      .SaveData "DotColor", DotColor, 0
      .SaveData "LineWidth", LineWidth, 0
      .SaveData "LineColor", LineColor, 0
      .SaveData "LinkFrom", LinkFrom, 0
      .SaveData "LinkTo", LinkTo, 0
      If ShapeType = OT_Line And Not LinkObject Is Nothing Then
        If TypeOf LinkObject Is clsBuilding Then
          .SaveData "Link", OT_Building, 0
        ElseIf TypeOf LinkObject Is clsMacro Then
          .SaveData "Link", OT_Macro, 0
        End If
      End If
      .SaveData "V1", Visibility, 0
    Else
      .SaveData "AreaType", Extra3, 0
    End If
    .SaveData "Extra1", Extra1, 0
    .SaveData "Extra2", Extra2, False
    .SaveData "Extra3", Extra3, 0
    .WriteSection
  End With

  If Not LinkObject Is Nothing Then
    LinkObject.SaveData FileData
  End If
  For I = 0 To myNumPoints
    myPoints(I).SavePointData FileData
  Next I
End Sub

Public Sub ScanTextures(ByRef Files() As String, ByRef FileCount As Integer)
  Dim Base As String, Flags As String, Ext As String
  
  If Texture <> "" Then
    AddTexFile Files(), FileCount, Mid$(Texture, InStr(Texture, " ") + 1)
    If Options.FSVersion >= Version_FS2K Then
      ' Check for flags
      Flags = Left$(Texture, InStr(Texture, " ") - 1)
      Base = Mid$(Texture, Len(Flags) + 2)
      Ext = "." & ReadLast(Base, ".")
      
      If -(InStr(Flags, "L") > 0) Then AddTexFile Files(), FileCount, Base & "_LM" & Ext
      If -(InStr(Flags, "S") > 0) Then AddTexFile Files(), FileCount, Base & "_SP" & Ext
      If -(InStr(Flags, "F") > 0) Then AddTexFile Files(), FileCount, Base & "_FA" & Ext
      If -(InStr(Flags, "W") > 0) Then AddTexFile Files(), FileCount, Base & "_WI" & Ext
      If -(InStr(Flags, "H") > 0) Then AddTexFile Files(), FileCount, Base & "_HW" & Ext
    End If
  End If
  
  If ShapeType = OT_Line And Not LinkObject Is Nothing Then
    If TypeOf LinkObject Is clsMacro Then
      Dim MacroObject As clsMacro
      Set MacroObject = LinkObject
      MacroObject.ScanTextures Files, FileCount
      Set MacroObject = Nothing
    End If
  End If

End Sub

' Check if user clicked on a line, and create a point if so
Public Function TestPoint(ByVal X As Single, ByVal Y As Single) As Boolean
  Dim M As Single, B As Single, D As Single, _
      I As Integer, J As Integer, GoUpTo As Integer, _
      BufferLength As Single
  
  Select Case ShapeType
    Case OT_Polygon, OT_FlatArea
      GoUpTo = myNumPoints
    Case Else
      GoUpTo = myNumPoints - 1
  End Select
  BufferLength = picEditor.ScaleX(Point_Threshold)
  
  If myNumPoints < PointUpperLimit() Then
    For I = 0 To GoUpTo
      ' Find m and b in the equ y = mx+b
      ' The distance formula is |mx - y + b|
      '                         sqr(m * m + 1)
      J = (I + 1) Mod (myNumPoints + 1)
      
      If (myPoints(J).X - myPoints(I).X) = 0 Then
        ' Vertical line
        D = Abs(X - myPoints(J).X)
      Else
        M = (myPoints(J).Y - myPoints(I).Y) / (myPoints(J).X - myPoints(I).X)
        B = myPoints(I).Y - M * myPoints(I).X
        D = (M * X - Y + B) / (Sqr(M * M + 1))
      End If
      If Abs(D) < BufferLength Then
        ' We need to test if the X and Y are between the
        ' two points as the distance formula applies for
        ' infinite lines
        If PtBetween(myPoints(I).X, myPoints(J).X, X) And PtBetween(myPoints(I).Y, myPoints(J).Y, Y) Then
          ' We've got a point!
          InsertNewPoint I, X, Y
          TestPoint = True
          Exit For
        End If
      End If
    Next I
  End If
End Function

' Check if there is any point within the specified coordinates, and return the coordinate of the point of the nearest line
Public Function TestPoint2(ByRef X As Single, ByRef Y As Single) As Boolean
  Dim M As Single, B As Single, D As Single, _
      I As Integer, J As Integer, GoUpTo As Integer, _
      BufferLength As Single, Rad As Single

  Select Case ShapeType
    Case OT_Polygon
      GoUpTo = myNumPoints
    Case OT_FlatArea
      Exit Function
    Case Else
      GoUpTo = myNumPoints - 1
  End Select
  BufferLength = picEditor.ScaleX(Point_Threshold)
  
  If myNumPoints < PointUpperLimit() Then
    For I = 0 To GoUpTo
      ' Find m and b in the equ y = mx+b
      ' The distance formula is |mx - y + b|
      '                         sqr(m * m + 1)
      J = (I + 1) Mod (myNumPoints + 1)
      If Scenery.Focus(myPoints(J).SceneryIndex) Then
        ' do not check points that are being dragged
        ' Since the next point is selected, we can skip the next iteration as well
        I = I + 1
      ElseIf Scenery.Focus(myPoints(I).SceneryIndex) Then
        ' do not check this iteration, but not possible to skip the next iteration
      Else
        If (myPoints(J).X - myPoints(I).X) = 0 Then
          ' Vertical line
          D = Abs(X - myPoints(J).X)
        Else
          M = (myPoints(J).Y - myPoints(I).Y) / (myPoints(J).X - myPoints(I).X)
          B = myPoints(I).Y - M * myPoints(I).X
          D = (M * X - Y + B) / (Sqr(M * M + 1))
        End If
        If Abs(D) < BufferLength Then
          ' We need to test if the X and Y are between the
          ' two points as the distance formula applies for
          ' infinite lines
          If PtBetween(myPoints(I).X, myPoints(J).X, X) And PtBetween(myPoints(I).Y, myPoints(J).Y, Y) Then
            ' We've got a point!
            Rad = Atn(M)
            X = X - D * Sin(Rad)
            Y = Y + D * Cos(Rad)
            TestPoint2 = True
            Exit For
          End If
        End If
      End If
    Next I
  End If
End Function

Public Sub UpdateObject()
  Dim NewglID As Integer, File As String
  
  mDrawable = True
  
  NewglID = -1
  
  picEditor.StartDraw
  
  If Texture <> "" Then
    File = Mid$(Texture, InStr(Texture, " ") + 1)
    If glBitmapID > -1 Then
      If picEditor.BitmapName(glBitmapID) = File Then GoTo lblUpdatePoints
    End If
    NewglID = picEditor.LoadBitmap(File)
  End If

  If glBitmapID > -1 Then picEditor.DeleteBitmap glBitmapID
  glBitmapID = NewglID
  
lblUpdatePoints:
  UpdatePoints ShapeType
End Sub

Public Sub UpdatePoints(ByVal myColor As Long)
  Dim I As Integer, J As Integer, _
    Angle As Single, deltaX As Single, deltaY As Single, _
    TempDist As Single, CompX As Single, CompY As Single
  
  If myNumPoints > -1 Then
    ReDim Pts(myNumPoints)
    If (Options.FSVersion <= Version_CFS1 And Texture <> "") Or (ShapeType = OT_Polygon And Options.FillPolygons And myColor <> COLOR_DragColor And Texture <> "") Then
      ' We're texturing, so we need to calculate
      ' the center reference point
      SumX = 0
      SumY = 0
      For I = 0 To myNumPoints
        myPoints(I).GetPoint myColor, Pts(I).X, Pts(I).Y
        SumX = SumX + Pts(I).X
        SumY = SumY + Pts(I).Y
      Next I
      AvgX = SumX / (myNumPoints + 1)
      AvgY = SumY / (myNumPoints + 1)
    Else
      ' Just load the points
      For I = 0 To myNumPoints
        myPoints(I).GetPoint myColor, Pts(I).X, Pts(I).Y
      Next I
    End If
    
    If ShapeType = OT_FlatArea Or _
       (ShapeType = OT_Line And (DotSpacing > 0 Or Extra1 < 0.2)) Or _
       (ShapeType = OT_Polygon And (Not Options.FillPolygons Or myColor = COLOR_DragColor Or (Color = 0 And glBitmapID = -1))) Then
      ' Basic 1-pixel width lines (closed figure)
    
      ' Nothing here
    ElseIf ShapeType = OT_Polygon Then ' (and Options.FillObject)
      ' Filled polygon
      glTesselate Pts, TessPts, (Options.FSVersion <= Version_CFS1)
    ElseIf myColor = COLOR_DragColor Or Not Options.ThickLines Or (Color = 0 And glBitmapID = -1) Then
      ' 1-pixel width lines and dashed lines
      
      ' Nothing here
    Else
      ' Thick lines
      
      ' We need to extrapolate the polygons that we need
      ' to draw for the taxiway/road/river
      
      ReDim TessPts(myNumPoints)
      J = -1
      
      If Options.FSVersion <= Version_CFS1 Then
        ' Simple case, just calculate the polygons
        ' and use a linear scale for texture coordinates
        For I = 0 To myNumPoints - 1
          If myPoints(I).PtType = FSNormal Then
            J = J + 1
            With TessPts(J)
              .Mode = 9 'GL_POLYGON
              ReDim .Pts(5)
              Angle = Atn2((Pts(I).Y - Pts(I + 1).Y), (Pts(I).X - Pts(I + 1).X))
              deltaX = Extra1 / 2 * Sin(Angle)
              deltaY = Extra1 / 2 * Cos(Angle)
              .Pts(0) = MakePoint(Pts(I).X + deltaX, Pts(I).Y - deltaY)
              .Pts(1) = MakePoint(Pts(I).X - deltaX, Pts(I).Y + deltaY)
              .Pts(2) = MakePoint(Pts(I + 1).X - deltaX, Pts(I + 1).Y + deltaY)
              .Pts(5) = MakePoint(Pts(I + 1).X + deltaX, Pts(I + 1).Y - deltaY)
            End With
        
            If J > 0 Then
              If myPoints(I - 1).PtType <> FSMoveTaxi Then
                TessPts(J - 1).Pts(3) = TessPts(J).Pts(1)
                TessPts(J - 1).Pts(4) = TessPts(J).Pts(0)
              Else
                TessPts(J - 1).Pts(3) = TessPts(J - 1).Pts(2)
                TessPts(J - 1).Pts(4) = TessPts(J - 1).Pts(5)
              End If
            End If
          End If
        Next I
        
        TessPts(J).Pts(3) = TessPts(J).Pts(2)
        TessPts(J).Pts(4) = TessPts(J).Pts(5)
        
        ReDim Preserve TessPts(J)
      Else
        ' Complicated case, calculate the polygons as
        ' well, but now we also have to calculate the
        ' correct texture coordinates
        
        ReDim TexCoord(myNumPoints)
        
        For I = 0 To myNumPoints - 1
          If myPoints(I).PtType = FSNormal Then
            J = J + 1
            With TessPts(J)
              .Mode = 9 'GL_POLYGON
              ReDim .Pts(3)
              
              ' Figure out the boundary of the current segment
              Angle = Atn2((Pts(I).Y - Pts(I + 1).Y), (Pts(I).X - Pts(I + 1).X))
              deltaX = Extra1 / 2 * Sin(Angle)
              deltaY = Extra1 / 2 * Cos(Angle)
              
              .Pts(0) = MakePoint(Pts(I).X + deltaX, Pts(I).Y - deltaY)
              .Pts(1) = MakePoint(Pts(I).X - deltaX, Pts(I).Y + deltaY)
              .Pts(2) = MakePoint(Pts(I + 1).X - deltaX, Pts(I + 1).Y + deltaY)
              .Pts(3) = MakePoint(Pts(I + 1).X + deltaX, Pts(I + 1).Y - deltaY)
            End With
        
            If J > 0 Then
              If myPoints(I - 1).PtType = FSNormal Then
                ' Extrapolate the intersection of the
                ' current and prevous line segment
                FindIntersection TessPts(J - 1).Pts(0), TessPts(J - 1).Pts(3), TessPts(J).Pts(0), TessPts(J).Pts(3)
                FindIntersection TessPts(J - 1).Pts(1), TessPts(J - 1).Pts(2), TessPts(J).Pts(1), TessPts(J).Pts(2)
              End If
            End If
    
            With TexCoord(J)
              ReDim .Pts(3)
              .Pts(0).X = 1
              .Pts(1).X = 0
              .Pts(2).X = 0
              .Pts(3).X = 1
              
              TempDist = Int((Distance(TessPts(J).Pts(1).X, TessPts(J).Pts(1).Y, Pts(I).X, Pts(I).Y) / IIf(Extra1 < 1, 1, Extra1) + 1)) * Extra1
              
              CompX = TempDist * Cos(Angle)
              CompY = TempDist * Sin(Angle)
              
              TempDist = Distance2(0, 0, deltaX, deltaY)
              
              .Pts(0).Y = Sqr(Round(Distance2(TessPts(J).Pts(1).X, TessPts(J).Pts(1).Y, Pts(I).X + CompX, Pts(I).Y + CompY) - TempDist, 2)) / Extra1 * 2
              .Pts(1).Y = Sqr(Round(Distance2(TessPts(J).Pts(0).X, TessPts(J).Pts(0).Y, Pts(I).X + CompX, Pts(I).Y + CompY) - TempDist, 2)) / Extra1 * 2
            End With
            
            If J > 0 Then
              TexCoord(J - 1).Pts(2).Y = TexCoord(J - 1).Pts(1).Y - Distance(TessPts(J - 1).Pts(1).X, TessPts(J - 1).Pts(1).Y, TessPts(J - 1).Pts(2).X, TessPts(J - 1).Pts(2).Y) / Extra1 * 2
              TexCoord(J - 1).Pts(3).Y = TexCoord(J - 1).Pts(0).Y - Distance(TessPts(J - 1).Pts(0).X, TessPts(J - 1).Pts(0).Y, TessPts(J - 1).Pts(3).X, TessPts(J - 1).Pts(3).Y) / Extra1 * 2
            End If
          End If
        Next I
      
        TexCoord(J).Pts(2).Y = TexCoord(J).Pts(1).Y - Distance(TessPts(J).Pts(1).X, TessPts(J).Pts(1).Y, TessPts(J).Pts(2).X, TessPts(J).Pts(2).Y) / Extra1 * 2
        TexCoord(J).Pts(3).Y = TexCoord(J).Pts(0).Y - Distance(TessPts(J).Pts(0).X, TessPts(J).Pts(0).Y, TessPts(J).Pts(3).X, TessPts(J).Pts(3).Y) / Extra1 * 2
      
        ReDim Preserve TessPts(J)
        ReDim Preserve TexCoord(J)
      End If
    End If
  End If
End Sub

Private Function FindArcPoints(ByVal ArcRadius As Single, Pt1 As PointType, Pt2 As PointType, Pt3 As PointType, PtRes1 As PointType, PtRes2 As PointType) As Boolean
  Dim Theta As Double
  '              (x1*x2 + y1*y2)
  ' cos(theta) = ---------------
  '                  |x||y|

  Dim X1 As Double, X2 As Double, _
      Y1 As Double, Y2 As Double, _
      D1 As Single, D2 As Single, _
      DX As Single, DY As Single, _
      Dist As Double
  X1 = Pt2.X - Pt1.X
  Y1 = Pt2.Y - Pt1.Y
  X2 = Pt2.X - Pt3.X
  Y2 = Pt2.Y - Pt3.Y
  
  D1 = Distance(0, 0, X1, Y1)
  D2 = Distance(0, 0, X2, Y2)

  If D1 < ArcRadius Or D2 < ArcRadius Then FindArcPoints = False: Exit Function
  
  Theta = Arccos((X1 * X2 + Y1 * Y2) / D1 / D2)
  Dist = ArcRadius / Atn(Theta / 2)
  
  DX = X1 / D1 * Dist
  DY = Y1 / D1 * Dist
  PtRes1.X = Pt2.X - DX
  PtRes1.Y = Pt2.Y - DY
  
  DX = X2 / D2 * Dist
  DY = Y2 / D2 * Dist
  PtRes2.X = Pt2.X - DX
  PtRes2.Y = Pt2.Y - DY

  FindArcPoints = True
End Function

Private Sub FindIntersection(Pt1 As PointType, Pt1A As PointType, Pt2 As PointType, Pt2A As PointType)
  Dim D1 As Single, D2 As Single, _
    M1 As Single, M2 As Single, _
    B1 As Single, B2 As Single, _
    TempVal As PointType

  D1 = Pt1A.X - Pt1.X
  If D1 <> 0 Then
    M1 = (Pt1A.Y - Pt1.Y) / D1
    B1 = Pt1A.Y - M1 * Pt1A.X
  End If

  D2 = Pt2A.X - Pt2.X
  If D2 <> 0 Then
    M2 = (Pt2A.Y - Pt2.Y) / D2
    B2 = Pt2A.Y - M2 * Pt2A.X
  End If
  
  If D1 <> 0 And D2 <> 0 Then
    If M1 - M2 <> 0 Then
      TempVal.X = (B2 - B1) / (M1 - M2)
      TempVal.Y = M1 * TempVal.X + B1
      Pt1A = TempVal
      Pt2 = TempVal
    Else
      ' nothing
    End If
  ElseIf D1 <> 0 Then
    TempVal.X = Pt2A.X
    TempVal.Y = M1 * TempVal.X + B1
    Pt1A = TempVal
    Pt2 = TempVal
  ElseIf D2 <> 0 Then
    TempVal.X = Pt1A.X
    TempVal.Y = M2 * TempVal.X + B2
    Pt1A = TempVal
    Pt2 = TempVal
  Else
    ' nothing
  End If
End Sub

Public Sub VersionChanged(ByVal NewVersion As Integer)
  clsObject_VersionChanged NewVersion
End Sub

Private Function clsObject_Add(ByVal NewX As Single, ByVal NewY As Single) As Boolean

End Function

Private Property Get clsObject_Caption(Optional ByVal Default As Boolean = False) As String
  clsObject_Caption = Caption(Default)
End Property

Private Sub clsObject_Compile(ByVal FileNum As Integer)
  
End Sub

Private Property Get clsObject_Complexity() As Byte

End Property

Private Property Let clsObject_Complexity(ByVal RHS As Byte)

End Property

Private Sub clsObject_CopyTo(Destination As clsObject)
  CopyTo Destination
End Sub

Private Sub clsObject_DrawBottom(Optional ByVal DrawX As Single = -999999#, Optional ByVal DrawY As Single)

End Sub

Private Sub clsObject_DrawTop(Optional ByVal DrawX As Single = -999999#, Optional ByVal DrawY As Single)

End Sub

Private Function clsObject_EditProperties() As Boolean

End Function

Private Sub clsObject_LoadAPTData(FileData As clsAPTData)
  LoadAPTData FileData
End Sub

Private Sub clsObject_LoadBinaryData(ByVal FileNum As Integer, ByVal Version As Integer, ByVal UserData As Long)
  LoadBinaryData FileNum, Version
End Sub

Private Sub clsObject_LoadData(FileData As clsFileData, ByVal Version As Integer)

End Sub

Private Property Get clsObject_Locked() As Boolean

End Property

Private Property Let clsObject_Locked(ByVal RHS As Boolean)

End Property

Private Property Get clsObject_Name() As String

End Property

Private Property Let clsObject_Name(ByVal RHS As String)

End Property

Private Property Get clsObject_ObjectIndex() As Integer
  clsObject_ObjectIndex = ObjectIndex
End Property

Private Property Let clsObject_ObjectIndex(ByVal RHS As Integer)
  ObjectIndex = RHS
End Property

Private Property Get clsObject_ObjectType() As ObjectTypeEnum
  clsObject_ObjectType = ShapeType
End Property

Private Property Get clsObject_PositionX() As Single

End Property

Private Property Let clsObject_PositionX(ByVal RHS As Single)

End Property

Private Property Get clsObject_PositionY() As Single

End Property

Private Property Let clsObject_PositionY(ByVal RHS As Single)

End Property

Private Property Get clsObject_Rotation() As Single

End Property

Private Property Let clsObject_Rotation(ByVal RHS As Single)

End Property

Private Sub clsObject_SaveData(FileData As clsFileData)

End Sub

Private Property Get clsObject_SceneryIndex() As Integer

End Property

Private Property Let clsObject_SceneryIndex(ByVal RHS As Integer)

End Property

Private Property Get clsObject_SortIndex() As Long

End Property

Private Sub clsObject_UpdateObject()

End Sub

Private Sub clsObject_VersionChanged(ByVal NewVersion As Integer)
  ' Type of road
  If ShapeType = OT_Road And Extra3 > 0 Then Extra3 = 0
  ' May need to change tesselation
  UpdatePoints ShapeType
End Sub
