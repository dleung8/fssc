VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRunway"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Runway Class
' Provides functions for using runways in the program

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
Public Length As Integer, _
       Width As Integer, _
       Rotation As Single, _
       ID As String, _
       Surface As Byte, _
       Markers As Byte, _
       Markers2 As Byte, _
       EdgeLights As Byte, _
       CenterLights As Byte, _
       RunwayPos As Boolean, _
       AFDEntry As Boolean, _
       RunwayVisible As Boolean

Public Near As clsRunwayOpt, _
       Far As clsRunwayOpt

Public mDrawable As Boolean

' Drawing cache
Private Pts() As PointType, TexCoords() As PointType
Private glBitmapID As Integer, _
        glBitmapIDN As Integer, _
        glBitmapIDF As Integer
Private MarkerPts() As PointType
Private ExtLineN() As PointType, _
        ExtPtsN() As PointType, _
        ExtCoordsN() As PointType, _
        ExtLineF() As PointType, _
        ExtPtsF() As PointType, _
        ExtCoordsF() As PointType, _
        ILSCoordsN() As PointType, _
        ILSCoordsF() As PointType

Private WirePts() As PointType
Private ExtWirePts() As PointType

Public Property Get Caption(Optional ByVal Default As Boolean = False)
  Caption = clsObject_Caption(Default)
End Property

Private Function GetRunwayTexture(ByVal Base As Byte, ByVal Ext As Byte) As String
  If Ext = 99 Then
    GetRunwayTexture = AddDir(Options.FSPath, "Texture\Runway") & Format$(Base, "00") & IIf(Base >= 10, ".bmp", ".r8")
  Else
    GetRunwayTexture = AddDir(Options.FSPath, "Texture\Runway") & Format$(Ext, "00") & IIf(Ext >= 10, ".bmp", ".r8")
  End If
End Function

Private Sub Class_Initialize()
  glBitmapID = -1
  glBitmapIDN = -1
  glBitmapIDF = -1
  Set Near = New clsRunwayOpt
  Set Far = New clsRunwayOpt
End Sub

Private Sub Class_Terminate()
  Set Near = Nothing
  Set Far = Nothing
  If glBitmapID > -1 Then picEditor.DeleteBitmap glBitmapID
  If glBitmapIDN > -1 Then picEditor.DeleteBitmap glBitmapIDN
  If glBitmapIDF > -1 Then picEditor.DeleteBitmap glBitmapIDF
'  If Closing Then Error 1
End Sub

Private Function clsObject_Add(ByVal NewX As Single, ByVal NewY As Single) As Boolean
  Dim TempInt As Integer, boolVal As Boolean
  
  If Defaults(OT_Runway) Is Nothing Or Not Options.Remember Then
    ' Default values:
    Length = 3000
    Width = 46
    Surface = 1
    RunwayPos = True
    RunwayVisible = True
    AFDEntry = True
    
    TempInt = CInt(EnsureRotation(ReturnPoint(X, Y).MagVar()) / 10)
    If TempInt = 0 Then TempInt = 36
    ID = Format$(TempInt)
    
    With Near
      Set .Parent = Me
      .Automatic = True
      .ExtSurface = 99
      .GlideSlope = 3
      .ILSBeamWidth = 2.5
      .ILSFlags = 193
      .ILSRange = 16
      .ILSFrequency = 110.5
      .ILSName = "ILSA"
      .ILSID = .ILSName
    End With
    
    With Far
      Set .Parent = Me
      .Automatic = True
      .ExtSurface = 99
      .GlideSlope = 3
      .ILSBeamWidth = 2.5
      .ILSFlags = 193
      .ILSRange = 16
      .ILSFrequency = 110.5
      .ILSName = "ILSB"
      .ILSID = .ILSName
    End With
  Else
    Defaults(OT_Runway).CopyTo Me
  End If
  X = NewX
  Y = NewY
  mDrawable = True
  boolVal = clsObject_EditProperties()
  If Not boolVal Then
    Set Near = Nothing
    Set Far = Nothing
  End If
  clsObject_Add = boolVal
End Function

Private Property Get clsObject_Caption(Optional ByVal Default As Boolean = False) As String
  If Name = "" Or Default Then
    clsObject_Caption = Lang.GetString(RES_Obj_Runway) & " " & ObjectIndex
  Else
    clsObject_Caption = Name
  End If
End Property

Private Sub clsObject_Compile(ByVal FileNum As Integer)
  Dim myX As Single, myY As Single, _
    myLatLon As clsLatLon, Pos As String, _
    OppRotation As Single, _
    strBuffer As String, _
    I As Integer

  Const RunwayLayer = 24
  
  If RunwayPos Then
    ' subtract scenery.header.rotation since ReturnPoint(CalcDist) adds it back
    PolarToRect Length / 2 - Near.ThrLength, 90 - Rotation + Scenery.Header.Rotation, myX, myY
  End If
  myX = X + myX
  myY = Y + myY
  Set myLatLon = ReturnPoint(myX, myY)
  Pos = myLatLon.LatLon

  OppRotation = EnsureRotation(Rotation + 180)

  If RunwayVisible Then
    Print #FileNum, "; " & clsObject_Caption
    Print #FileNum, "Area( 5 " & Pos & " 12 )"
    Print #FileNum, "    LayerCall( :lblLayer " & RunwayLayer & " )"
    Print #FileNum, "    Jump( : )"
    Print #FileNum, ":lblLayer"
    
    If Options.FSVersion >= Version_FS2K Then
      Print #FileNum, "    Runway( " & Pos
    Else
      Print #FileNum, "    RunwayData( " & Pos
    End If
      
    Print #FileNum, "            Altitude    " & FloatFormat(Scenery.Header.Altitude, "####0.0#")
    Print #FileNum, "            Heading     " & FloatFormat(Rotation, "##0.0#")
    Print #FileNum, "            Length      " & CInt(Length * MToFt)
    Print #FileNum, "            Width       " & CInt(Width * MToFt)
    Print #FileNum, "            ID          " & ID
    If Surface = 99 Then
      Print #FileNum, "            Surface     -1"
    Else
      Print #FileNum, "            Surface     0x" & Hex$(CInt(Surface))
    End If
    Print #FileNum, "            Markers     " & Hex$(CInt(Markers))
    If Options.FSVersion >= Version_FS2K Then
      Print #FileNum, "            Threshold   " & Hex$(CInt(Markers2))
    End If
    Print #FileNum, "            Lights      " & Hex$(EdgeLights + CenterLights * 4)
    
    Near.Compile FileNum, "N"
    Far.Compile FileNum, "F"
    Print #FileNum, "    )"

    Dim NearDist As Single, FarDist As Single
    NearDist = Length / 2 - Near.ThrLength - 100
    FarDist = Length / 2 - Far.ThrLength - 100
  
    If Near.PAPI > 0 Then
      Print #FileNum, "    RefPoint( abs :EndPAPIN" & SceneryIndex & " 1 " & Pos & " E= " & FloatFormat(Scenery.Header.Altitude, "####0.0#") & " v1= 10000 v2= 0" & " )"
      Print #FileNum, "    RotatedCall( :StartPAPIN" & SceneryIndex & " 0 0 " & FloatFormat(Rotation, "##0.0#") & " )"
      Print #FileNum, "    Jump( :EndPAPIN" & SceneryIndex & " )"
      WritePAPI FileNum, "N" & SceneryIndex, NearDist, Near.GlideSlope
    End If
    If Far.PAPI > 0 Then
      Print #FileNum, "    RefPoint( abs :EndPAPIF" & SceneryIndex & " 1 " & Pos & " E= " & FloatFormat(Scenery.Header.Altitude, "####0.0#") & " v1= 10000 v2= 0" & " )"
      Print #FileNum, "    RotatedCall( :StartPAPIF" & SceneryIndex & " 0 0 " & FloatFormat(OppRotation, "##0.0#") & " )"
      Print #FileNum, "    Jump( :EndPAPIF" & SceneryIndex & " )"
      WritePAPI FileNum, "F" & SceneryIndex, FarDist, Far.GlideSlope
    End If
    
    Print #FileNum, "    Return"
    Print #FileNum, "EndA" & vbCrLf
  End If

  If AFDEntry Then
    Dim RwyID As String, RwyType As String, RwySurface As Byte, LowRotation As Single

    If Rotation >= 180 Then
      LowRotation = Rotation - 180
      RwyID = ReverseRunway(ID)
    Else
      LowRotation = Rotation
      RwyID = ID
    End If

    If InStr(RwyID, "L") Then
      RwyType = "L"
    ElseIf InStr(RwyID, "R") Then
      RwyType = "R"
    ElseIf InStr(RwyID, "C") Then
      RwyType = "C"
    Else
      RwyType = "N"
    End If

    If Between(Surface, 5, 8) Then
      RwySurface = Surface
    ElseIf Surface = 4 Then
      RwySurface = 9
    Else
      RwySurface = Surface + 1
    End If

    If Options.FSVersion >= Version_FS2K2 Then
      Scenery.RunwayNAVBuffer = Scenery.RunwayNAVBuffer & "    RwyLoc( " & Pos & " " & FloatFormat(Scenery.Header.Altitude, "####0.0#") & " " & FloatFormat(LowRotation, "##0.0#") & " " & FloatFormat(ReturnPoint(myX, myY).MagVar, "##0.0#") & " " & CInt(Length * MToFt) & " " & CInt(Width * MToFt) & " " & RwyType & " " & Val(RwyID) & " " & RwySurface & " 0x60 " & FloatFormat(Scenery.Header.Altitude, "####0.0#") & " )" & vbCrLf & vbCrLf
    Else
      Scenery.RunwayNAVBuffer = Scenery.RunwayNAVBuffer & "    RwyLoc( " & Pos & "  " & FloatFormat(Scenery.Header.Altitude, "####0.0#") & " " & FloatFormat(LowRotation, "##0.0#") & " " & FloatFormat(ReturnPoint(myX, myY).MagVar, "##0.0#") & " " & CInt(Length * MToFt) & " " & CInt(Width * MToFt) & " " & RwyType & " " & Val(RwyID) & " " & RwySurface & " )" & vbCrLf & vbCrLf
    End If
    If Scenery.NameListABuffer = "" Then
      Scenery.NameListABuffer = "    NameEntry( AIRPORT " & Scenery.Count + 1 & " " & QuoteString(Scenery.Header.Name) & " )" & vbCrLf
    End If
    Scenery.RunwayLocBuffer = Scenery.RunwayLocBuffer & "    SetupLoc( " & myLatLon.CalcPoint((Length / 2 - 20) * MToNm, LowRotation + 180).LatLon & " " & FloatFormat(Scenery.Header.Altitude, "####0.0#") & " " & FloatFormat(LowRotation, "##0.0#") & " 1 " & RwyType & " " & Val(RwyID) & " )" & vbCrLf

    RwyID = ReverseRunway(RwyID)
    If InStr(RwyID, "L") Then
      RwyType = "L"
    ElseIf InStr(RwyID, "R") Then
      RwyType = "R"
    ElseIf InStr(RwyID, "C") Then
      RwyType = "C"
    Else
      RwyType = "N"
    End If
    Scenery.RunwayLocBuffer = Scenery.RunwayLocBuffer & "    SetupLoc( " & myLatLon.CalcPoint((Length / 2 - 20) * MToNm, LowRotation).LatLon & " " & FloatFormat(Scenery.Header.Altitude, "####0.0#") & " " & FloatFormat(EnsureRotation(LowRotation + 180), "##0.0#") & " 1 " & RwyType & " " & Val(RwyID) & " )" & vbCrLf
  End If
  
  strBuffer = strBuffer & Near.CompileBeacons(myLatLon, Rotation, OppRotation, NearDist, FarDist)
  strBuffer = strBuffer & Far.CompileBeacons(myLatLon, OppRotation, Rotation, FarDist, NearDist)
  
  If strBuffer <> "" Then
    Print #FileNum, "; Runway Beacons"
    Print #FileNum, strBuffer
  End If

  If Options.FSVersion < Version_FS2K Then
    Print #FileNum, "LandMe( " & myLatLon.CalcPoint(NearDist * MToNm, OppRotation).LatLon & " " & FloatFormat(Scenery.Header.Altitude, "####0.0#") & " " & ID & " " & myLatLon.CalcPoint(FarDist * MToNm, Rotation).LatLon & " " & FloatFormat(Scenery.Header.Altitude, "####0.0#") & " " & ReverseRunway(ID) & " " & ReplaceParens(Replace(Scenery.Header.Name, " ", "_")) & " )" & vbCrLf
  End If
End Sub

Private Property Get clsObject_Complexity() As Byte
'
End Property

Private Property Let clsObject_Complexity(ByVal RHS As Byte)
'
End Property

Private Sub clsObject_CopyTo(Destination As clsObject)
  Dim Temp As clsRunway

  If TypeOf Destination Is clsRunway Then
    Set Temp = Destination
    With Temp
      .Name = Name
      .Locked = Locked
      .X = X
      .Y = Y
      
      .Length = Length
      .Width = Width
      .Rotation = Rotation
      .ID = ID
      .Surface = Surface
      .Markers = Markers
      .Markers2 = Markers2
      .EdgeLights = EdgeLights
      .CenterLights = CenterLights
      .RunwayPos = RunwayPos
      .AFDEntry = AFDEntry
      .RunwayVisible = RunwayVisible
      
      Near.CopyTo .Near
      Far.CopyTo .Far
      
      Set .Near.Parent = Temp
      Set .Far.Parent = Temp
    End With
  End If
End Sub

Private Sub clsObject_DrawBottom(Optional ByVal DrawX As Single = -999999#, Optional ByVal DrawY As Single)
  Dim myColor As Long

  If Not mDrawable Or Not Options.ObjectVisible(OT_Runway) Then Exit Sub
  
  If Not RunwayVisible Then Exit Sub

  glPushMatrix
  If DrawX = -999999 Then
    ' No drag
    If Scenery.Focus(SceneryIndex) Then
      myColor = COLOR_FocusColor
    Else
      myColor = OT_Runway
    End If
    glTranslatef X, Y, 0
  Else
    ' Drag
    myColor = COLOR_DragColor
    glTranslatef DrawX, DrawY, 0
  End If
  glRotatef Scenery.Header.Rotation - Rotation, 0, 0, 1
  If RunwayPos Then glTranslatef 0, Length / 2 - Near.ThrLength, 0

  Options.SetGLColor myColor

  If myColor <> COLOR_DragColor Then
    If UBound(TexCoords) > 0 Then
      If glBitmapID > -1 Then
        picEditor.SelectBitmap glBitmapID
        glPaintTexturedRegion2 Pts, TexCoords
      End If
      If glBitmapIDN > -1 Then
        picEditor.SelectBitmap glBitmapIDN
        glPaintTexturedRegion2 ExtPtsN, ExtCoordsN
      End If
      If glBitmapIDF > -1 Then
        picEditor.SelectBitmap glBitmapIDF
        glPaintTexturedRegion2 ExtPtsF, ExtCoordsF
      End If
      
      glForeColor RGB(192, 192, 0)
      If glBitmapIDN > -1 Then glDrawQuads ExtLineN
      If glBitmapIDF > -1 Then glDrawQuads ExtLineF
      
      glForeColor vbWhite
      glDrawQuads MarkerPts
    End If
  End If
  glPopMatrix
End Sub

Private Sub clsObject_DrawTop(Optional ByVal DrawX As Single = -999999#, Optional ByVal DrawY As Single)
  Dim myColor As Long

  If Not mDrawable Or Not Options.ObjectVisible(OT_Runway) Then Exit Sub

  glPushMatrix
  If DrawX = -999999 Then
    ' No drag
    If Scenery.Focus(SceneryIndex) Then
      myColor = COLOR_FocusColor
    Else
      myColor = OT_Runway
    End If
    glTranslatef X, Y, 0
  Else
    ' Drag
    myColor = COLOR_DragColor
    glTranslatef DrawX, DrawY, 0
  End If
  glPushMatrix
  glRotatef Scenery.Header.Rotation - Rotation, 0, 0, 1
  If RunwayPos Then glTranslatef 0, Length / 2 - Near.ThrLength, 0

  Options.SetGLColor myColor

  If (myColor = COLOR_DragColor Or UBound(TexCoords) = 0) And RunwayVisible Then
    glDrawPolygon WirePts
    
    If (Markers And 16) > 0 Then
      glDashLine
      glBeginLines
      glDrawLine (WirePts(1).X + WirePts(2).X) / 2, (WirePts(1).Y + WirePts(2).Y) / 2, (WirePts(3).X + WirePts(4).X) / 2, (WirePts(3).Y + WirePts(4).Y) / 2
      glEnd
      glSmoothLine
    End If
  End If
  
  glForeColor RGB(192, 192, 0)
  If Near.ILSEnabled Then glDrawPolygon ILSCoordsN()
  If Far.ILSEnabled Then glDrawPolygon ILSCoordsF()
  
  glPopMatrix
  If DrawX = -999999 Then
    If Options.FocusCircle Then
      Options.SetGLColor myColor
      glDrawCircle X, Y, FocusWidth, Options.ObjectColor(myColor)
    End If
  End If
  glPopMatrix
End Sub

Private Function clsObject_EditProperties() As Boolean
  If frmRunway.EditData(Me) Then
    clsObject_EditProperties = True
    clsObject_UpdateObject
    If Defaults(OT_Runway) Is Nothing Then Set Defaults(OT_Runway) = New clsRunway
    clsObject_CopyTo Defaults(OT_Runway)
  End If
End Function

Private Sub clsObject_LoadAPTData(FileData As clsAPTData)
  Dim Temp As String, Temp2 As String, _
    byteTemp As Byte, I As Integer
  
  With FileData
    X = ValEx(.GetData("PosX"))
    Y = ValEx(.GetData("PosY"))
    Length = ValEx(Meter(.GetData("Longueur")))
    Width = ValEx(Meter(.GetData("Largeur")))
    Rotation = ValEx(.GetData("Cap")) - Scenery.Header.APTMagVar
    ID = Replace(.GetData("ID"), " ", "")
    Temp = .GetData("Type")
    Temp2 = ReadNext(Temp, " ")
    If Len(Temp2) <= 8 Then
      Markers = BinaryToDec(Temp2)
    Else
      Markers = -(Mid$(Temp2, 6, 1) = "1") * 1 + _
                -(Mid$(Temp2, 7, 1) = "1") * 2 + _
                -(Mid$(Temp2, 8, 1) = "1") * 4 + _
                -(Mid$(Temp2, 9, 1) = "1") * 8 + _
                -(Mid$(Temp2, 10, 1) = "1") * 16 + _
                -(Mid$(Temp2, 11, 1) = "1") * 32 + _
                -(Mid$(Temp2, 12, 1) = "1") * 64 + _
                -(Mid$(Temp2, 5, 1) = "1") * 128
      Markers2 = -(Mid$(Temp2, 1, 1) = "1") * 128 + _
                 -(Mid$(Temp2, 2, 1) = "1") * 64 + _
                 -(Mid$(Temp2, 3, 1) = "1") * 32 + _
                 -(Mid$(Temp2, 4, 1) = "1") * 16
    End If
    
    If Len(Temp2) = 7 Then
      Select Case ReadNext(Temp, " ")
        Case "0": EdgeLights = 0: CenterLights = 0
        Case "1": EdgeLights = 1: CenterLights = 0
        Case "2": EdgeLights = 3: CenterLights = 0
        Case "4": EdgeLights = 0: CenterLights = 1
        Case "5": EdgeLights = 1: CenterLights = 1
        Case "7": EdgeLights = 3: CenterLights = 3
      End Select
    Else
      Select Case ReadNext(Temp, " ")
        Case "0": EdgeLights = 0: CenterLights = 0
        Case "1": EdgeLights = 1: CenterLights = 0
        Case "2": EdgeLights = 2: CenterLights = 0
        Case "3": EdgeLights = 3: CenterLights = 0
        Case "4": EdgeLights = 1: CenterLights = 1
        Case "5": EdgeLights = 2: CenterLights = 2
        Case "6": EdgeLights = 3: CenterLights = 3
      End Select
      If ReadNext(Temp, " ") = "1" Then CenterLights = CenterLights Or 4
    End If
    
    Temp = .GetData("Surface")
    Surface = MatchText(Temp, "Terre", "Goudron", "Herbe", "Beton")
    If Surface = 10 Then Surface = 99
 
    Near.LoadAPTData .GetData("SeuilDebut"), .GetData("SignalDebut"), .GetData("VASIDebut"), .GetData("PAPIDebut"), .GetData("ILSDebut"), .GetData("ILSDebutRange"), .GetData("IMDebut"), .GetData("MMDebut"), .GetData("OMDebut")
    Far.LoadAPTData .GetData("SeuilFin"), .GetData("SignalFin"), .GetData("VASIFin"), .GetData("PAPIFin"), .GetData("ILSFin"), .GetData("ILSFinRange"), .GetData("IMFin"), .GetData("MMFin"), .GetData("OMFin")

    For I = 0 To 5
      Temp = .GetData("Beacons")
      If Temp = "" Then Exit For
      Select Case ValEx(ReadNext(Temp, " "))
        Case 0: Near.IM = Nautical(Temp)
        Case 1: Near.MM = Nautical(Temp)
        Case 2: Near.OM = Nautical(Temp)
        Case 3: Far.IM = Nautical(Temp)
        Case 4: Far.MM = Nautical(Temp)
        Case 5: Far.OM = Nautical(Temp)
      End Select
    Next I

    Name = .GetData("Comment")
  End With

  RunwayVisible = True
  Set Near.Parent = Me
  Set Far.Parent = Me
  
  AFDEntry = True
  Scenery.AFDRefresh = True
  
  mDrawable = True
  clsObject_UpdateObject
End Sub

Private Sub clsObject_LoadBinaryData(ByVal FileNum As Integer, ByVal Version As Integer, ByVal UserData As Long)
  Dim NearB As Byte, FarB As Byte, _
    NearI As Integer, FarI As Integer, _
    NearO As Boolean, FarO As Boolean, _
    NearS As Single, FarS As Single
  Dim sngW As Single, sngL As Single
  
  Dim Lights As Byte, Intensity As Byte
  Dim Mark1 As Byte, Mark2 As Byte
  
  Select Case Version
    Case Is >= 9
      GetBinaryData FileNum, Name, X, Y, Locked, Rotation, _
        sngL, sngW, ID, Surface, Mark1, Lights, _
        Intensity, RunwayPos, AFDEntry, Mark2
      Markers = (Mark1 And 127) + (Mark2 And 1) * 128
      Markers2 = -((Mark1 And 128) > 0) * 48
    Case Is >= 8
      GetBinaryData FileNum, Name, X, Y, Locked, Rotation, _
        sngL, sngW, ID, Surface, Markers, Lights, _
        Intensity, RunwayPos, AFDEntry
      If Intensity = 1 Then Intensity = 2
    Case Else
      GetBinaryData FileNum, X, Y, Rotation, sngL, sngW, ID, Surface, _
        Markers, Lights, Intensity
      If Intensity = 1 Then Intensity = 2
  End Select
  Length = sngL
  Width = sngW
  EdgeLights = -(Lights = 1 Or Lights = 3 Or Lights = 4) * (Intensity + 1)
  CenterLights = -(Lights = 2 Or Lights = 3 Or Lights = 4) * (Intensity + 1) + -(Lights = 4) * 4
  
'  If Version < 3 Then
'    GetData FileNum, NearB, FarB
'    Near.ThrLights = NearB:  Far.ThrLights = FarB
'    GetData FileNum, NearB, FarB
'
'    NearB = IIf(NearB > 3, NearB + 2, NearB)
'    FarB = IIf(FarB > 3, FarB + 2, FarB)
'
'    Near.ApprLights = NearB: Far.ApprLights = FarB
'    GetData FileNum, NearB, FarB
'    Near.VASIType = NearB:   Far.VASIType = FarB
'    GetData FileNum, NearB, FarB
'    Near.PAPIType = NearB:   Far.PAPIType = FarB
'    GetData FileNum, NearI, FarI
'    Near.ThrLength = NearI:  Far.ThrLength = FarI
'    GetData FileNum, NearI, FarI
'    Near.ExtLength = NearI:  Far.ExtLength = FarI
'    GetData FileNum, NearB, FarB
'    Near.NumStrobes = NearB: Far.NumStrobes = FarB
'    GetData FileNum, NearS, FarS
'    Near.GlideSlope = NearS: Far.GlideSlope = FarS
'    GetData FileNum, NearS, FarS
'    Near.IM = NearS:         Far.IM = FarS
'    GetData FileNum, NearS, FarS
'    Near.MM = NearS:         Far.MM = FarS
'    GetData FileNum, NearS, FarS
'    Near.OM = NearS:         Far.OM = FarS
'    GetData FileNum, NearO, FarO
'    Near.ILSEnabled = NearO: Far.ILSEnabled = FarO
'    Near.ILS.LoadData FileNum, 0, Version
'    Far.ILS.LoadData FileNum, 0, Version
' Else
    Near.LoadBinaryData FileNum, Version, 0
    Far.LoadBinaryData FileNum, Version, 0
'  End If

  If Version < 8 Then
    Near.Automatic = True
    Far.Automatic = True
  End If
  
  RunwayVisible = True
  Set Near.Parent = Me
  Set Far.Parent = Me
  
  If AFDEntry Then Scenery.AFDRefresh = True
  
  mDrawable = True
  clsObject_UpdateObject
End Sub

Private Sub clsObject_LoadData(FileData As clsFileData, ByVal Version As Integer)
  Dim lngValue As Long
  With FileData
    .GetData "Name", Name, ""
    .GetData "Locked", Locked, False
    .GetData "X", X, 0
    .GetData "Y", Y, 0
    .GetData "Length", Length, 50
    .GetData "Width", Width, 50
    .GetData "Rotation", Rotation, 0

    .GetData "ID", ID, ""
    .GetData "Surface", Surface, -1
    .GetData "Markers", lngValue, 0
    Markers = lngValue And 255
    Markers2 = lngValue \ 256

    .GetData "Edge Lights", EdgeLights, 0
    .GetData "Center Lights", CenterLights, 0
    .GetData "Runway Position", RunwayPos, True
    .GetData "AFD Entry", AFDEntry, True
    .GetData "Visible", RunwayVisible, True
  
    Near.GetData FileData, "Near "
    Far.GetData FileData, "Far "
    Set Near.Parent = Me
    Set Far.Parent = Me
  
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
  clsObject_ObjectType = OT_Runway
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
    .InitializeForWrite ObjectNames(OT_Runway), Caption
    .SaveData "Name", Name, ""
    .SaveData "Locked", Locked, False
    .SaveData "X", X, 0
    .SaveData "Y", Y, 0
    .SaveData "Length", Length, 50
    .SaveData "Width", Width, 50
    .SaveData "Rotation", Rotation, 0
    .SaveData "ID", ID, ""
    .SaveData "Surface", Surface, -1
    .SaveData "Markers", CLng(Markers) + CLng(Markers2) * 256, 0
    .SaveData "Edge Lights", EdgeLights, 0
    .SaveData "Center Lights", CenterLights, 0
    .SaveData "Runway Position", RunwayPos, True
    .SaveData "AFD Entry", AFDEntry, True
    .SaveData "Visible", RunwayVisible, True
    Near.SaveData FileData, "Near "
    Far.SaveData FileData, "Far "
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
  clsObject_SortIndex = &H8000& * 24 + SceneryIndex
End Property

Private Sub clsObject_UpdateObject()
  Dim HalfL As Single, HalfW As Single, _
    NearThr As Single, FarThr As Single
  Dim NumLines As Long, MaxPts As Long, Count As Long
  Dim I As Long

  mDrawable = True
  
  Dim NewglID As Integer, _
    NewglIDN As Integer, NewglIDF As Integer
  NewglID = -1
  NewglIDN = -1
  NewglIDF = -1

  If Not mDrawable Then Exit Sub
  ReDim TexCoords(0)

  HalfL = Length / 2
  HalfW = Width / 2
  NearThr = -HalfL + Near.ThrLength
  FarThr = HalfL - Far.ThrLength

  picEditor.StartDraw

  NewglID = picEditor.LoadBitmap(GetRunwayTexture(255, Surface))

  If Options.FillPolygons And NewglID > -1 And RunwayVisible Then
    ReDim TexCoords(3)
    If Surface <> 66 Then
      TexCoords(0) = MakePoint(0.5, 0)
      TexCoords(1) = MakePoint(0.75, 0)
      TexCoords(2) = MakePoint(0.75, -Length * MToFt / 256)
      TexCoords(3) = MakePoint(0.5, TexCoords(2).Y)
    Else
      ' Hack for steel mats in CFS2
      TexCoords(0) = MakePoint(0.375, 0)
      TexCoords(1) = MakePoint(0.125, 0)
      TexCoords(2) = MakePoint(0.125, -Length * MToFt / 256)
      TexCoords(3) = MakePoint(0.375, TexCoords(2).Y)
    End If
  
    ReDim Pts(3)
    Pts(0) = MakePoint(-HalfW - IIf((Markers And 128) > 0, 1, 0), -HalfL)
    Pts(1) = MakePoint(-Pts(0).X, Pts(0).Y)
    Pts(2) = MakePoint(-Pts(0).X, HalfL)
    Pts(3) = MakePoint(Pts(0).X, Pts(2).Y)
'    MultiRotate Pts, Rotation
  
    If (Markers And 16) > 0 Then
      NumLines = (FarThr - NearThr) * MToFt / 200
    End If
    If (Markers And 1) > 0 Then
      MaxPts = NumLines + 16
    Else
      MaxPts = NumLines
    End If
    
    If MaxPts > 0 Then
      ReDim MarkerPts(MaxPts * 4 - 1)
    
      For I = 0 To NumLines * 4 - 1 Step 4
        MarkerPts(I) = MakePoint(-1, (I / 4) * FtToM * 200 + NearThr)
        MarkerPts(I + 1) = MakePoint(-MarkerPts(I).X, MarkerPts(I).Y)
        MarkerPts(I + 2) = MakePoint(-MarkerPts(I).X, MarkerPts(I).Y + 120 * FtToM)
        MarkerPts(I + 3) = MakePoint(MarkerPts(I).X, MarkerPts(I + 2).Y)
      Next I
    
      Count = NumLines * 4
      
      If (Markers And 1) > 0 Then
        MarkerPts(Count) = MakePoint(-HalfW, -HalfL)
        MarkerPts(Count + 1) = MakePoint(-HalfW, HalfL)
        MarkerPts(Count + 2) = MakePoint(-HalfW + 1, HalfL)
        MarkerPts(Count + 3) = MakePoint(-HalfW + 1, -HalfL)

        MarkerPts(Count + 4) = MakePoint(HalfW, -HalfL)
        MarkerPts(Count + 5) = MakePoint(HalfW, HalfL)
        MarkerPts(Count + 6) = MakePoint(HalfW - 1, HalfL)
        MarkerPts(Count + 7) = MakePoint(HalfW - 1, -HalfL)

        MarkerPts(Count + 8) = MakePoint(-HalfW, NearThr)
        MarkerPts(Count + 9) = MakePoint(HalfW, NearThr)
        MarkerPts(Count + 10) = MakePoint(HalfW, NearThr + 1)
        MarkerPts(Count + 11) = MakePoint(-HalfW, NearThr + 1)

        MarkerPts(Count + 12) = MakePoint(-HalfW, FarThr)
        MarkerPts(Count + 13) = MakePoint(HalfW, FarThr)
        MarkerPts(Count + 14) = MakePoint(HalfW, FarThr - 1)
        MarkerPts(Count + 15) = MakePoint(-HalfW, FarThr - 1)
      End If
    Else
      ReDim MarkerPts(0)
    End If
  
'    MultiRotate MarkerPts, Rotation
    
    If Near.ExtLength > 0 Then
      NewglIDN = picEditor.LoadBitmap(GetRunwayTexture(Surface, Near.ExtSurface))
      ReDim ExtLineN(3)
      ReDim ExtPtsN(3)
      ReDim ExtCoordsN(3)
      
      ExtLineN(0) = MakePoint(-HalfW, -HalfL)
      ExtLineN(1) = MakePoint(HalfW, -HalfL)
      ExtLineN(2) = MakePoint(HalfW, -HalfL - 1)
      ExtLineN(3) = MakePoint(-HalfW, -HalfL - 1)
      
      ExtPtsN(0) = MakePoint(-HalfW, -HalfL)
      ExtPtsN(1) = MakePoint(HalfW, -HalfL)
      ExtPtsN(2) = MakePoint(HalfW, -HalfL - Near.ExtLength)
      ExtPtsN(3) = MakePoint(-HalfW, ExtPtsN(2).Y)
      
      ExtCoordsN(0) = MakePoint(0.5, 0)
      ExtCoordsN(1) = MakePoint(0.75, 0)
      ExtCoordsN(2) = MakePoint(0.75, -Near.ExtLength * MToFt / 256)
      ExtCoordsN(3) = MakePoint(0.5, ExtCoordsN(2).Y)
    Else
      ReDim ExtPtsN(0)
    End If
  
    If Far.ExtLength > 0 Then
      NewglIDF = picEditor.LoadBitmap(GetRunwayTexture(Surface, Far.ExtSurface))
      ReDim ExtLineF(3)
      ReDim ExtPtsF(3)
      ReDim ExtCoordsF(3)
      ExtLineF(0) = MakePoint(-HalfW, HalfL)
      ExtLineF(1) = MakePoint(HalfW, HalfL)
      ExtLineF(2) = MakePoint(HalfW, HalfL + 1)
      ExtLineF(3) = MakePoint(-HalfW, HalfL + 1)
    
      ExtPtsF(0) = MakePoint(-HalfW, HalfL)
      ExtPtsF(1) = MakePoint(HalfW, HalfL)
      ExtPtsF(2) = MakePoint(HalfW, HalfL + Far.ExtLength)
      ExtPtsF(3) = MakePoint(-HalfW, ExtPtsF(2).Y)
    
      ExtCoordsF(0) = MakePoint(0.5, 0)
      ExtCoordsF(1) = MakePoint(0.75, 0)
      ExtCoordsF(2) = MakePoint(0.75, -Far.ExtLength * MToFt / 256)
      ExtCoordsF(3) = MakePoint(0.5, ExtCoordsF(2).Y)
    Else
      ReDim ExtPtsF(0)
    End If
  Else
    ReDim TexCoords(0)
  End If
  
  ReDim WirePts(12)
  WirePts(0) = MakePoint(-HalfW, -HalfL)
  WirePts(1) = MakePoint(WirePts(0).X, NearThr)
  WirePts(2) = MakePoint(-WirePts(0).X, WirePts(1).Y)
  WirePts(3) = MakePoint(WirePts(2).X, FarThr)
  WirePts(4) = MakePoint(WirePts(0).X, WirePts(3).Y)
  WirePts(5) = MakePoint(WirePts(0).X, -WirePts(0).Y)
  WirePts(6) = MakePoint(WirePts(2).X, WirePts(5).Y)
  WirePts(7) = WirePts(3)
  WirePts(8) = WirePts(4)
  WirePts(9) = WirePts(1)
  WirePts(10) = WirePts(2)
  WirePts(11) = MakePoint(WirePts(2).X, WirePts(0).Y)
  WirePts(12) = WirePts(0)

  ReDim ILSCoordsN(2)
  ILSCoordsN(0) = MakePoint(0, WirePts(1).Y)
  ILSCoordsN(1) = MakePoint(WirePts(1).X * 2, WirePts(1).Y - 2000)
  ILSCoordsN(2) = MakePoint(-ILSCoordsN(1).X, ILSCoordsN(1).Y)

  ReDim ILSCoordsF(2)
  ILSCoordsF(0) = MakePoint(0, WirePts(3).Y)
  ILSCoordsF(1) = MakePoint(WirePts(4).X * 2, WirePts(4).Y + 2000)
  ILSCoordsF(2) = MakePoint(-ILSCoordsF(1).X, ILSCoordsF(1).Y)
  
  If glBitmapID > -1 Then picEditor.DeleteBitmap glBitmapID
  If glBitmapIDN > -1 Then picEditor.DeleteBitmap glBitmapIDN
  If glBitmapIDF > -1 Then picEditor.DeleteBitmap glBitmapIDF
  glBitmapID = NewglID
  glBitmapIDN = NewglIDN
  glBitmapIDF = NewglIDF
End Sub

Private Sub clsObject_VersionChanged(ByVal NewVersion As Integer)
  If NewVersion >= Version_FS2K Then
    Near.PAPI = 0
    Far.PAPI = 0
  Else
    If Near.VASILeft > 2 Then Near.VASILeft = 0
    If Far.VASILeft > 2 Then Far.VASILeft = 0
    Near.VASIRight = 0
    Far.VASIRight = 0
    Near.ThrLights = Near.ThrLights And 5
    Far.ThrLights = Far.ThrLights And 5
  End If
  Select Case NewVersion
    Case Version_FS95, Version_FS98, Version_CFS1
      If Surface > 3 Then Surface = 1
    Case Version_FS2K
      If Surface > 9 Then Surface = 1
      ' Even if the surface is 99, it will still be replaced with 99... one less test
      If Near.ExtSurface > 9 Then Near.ExtSurface = 99
      If Far.ExtSurface > 9 Then Far.ExtSurface = 99
    Case Version_CFS2
      If Not Between(Surface, 64, 68) And Surface > 3 Then Surface = 1
      ' Even if the surface is 99, it will still be replaced with 99... one less test
      If Not Between(Near.ExtSurface, 64, 68) And Near.ExtSurface > 3 Then Near.ExtSurface = 99
      If Not Between(Far.ExtSurface, 64, 68) And Far.ExtSurface > 3 Then Far.ExtSurface = 99
    Case Version_FS2K2
      ' nothing yet!
  End Select
End Sub

Public Sub WritePAPI(ByVal FileNum As Integer, ByVal Suffix As String, ByVal DistILS As Single, ByVal GlideSlope As Single)
  Dim HalfWidth As Single
  HalfWidth = Width / 2
  Print #FileNum, ":StartPAPI" & Suffix
  Print #FileNum, "    VectorJump( :Finish" & Suffix & " m 0 0 -32767 " & FloatFormat(DistILS, "####.0#") & " )"
  Print #FileNum, "    LineColor( 0F F0 )"
  Print #FileNum, "    VectorJump( :Line1" & Suffix & " m 0 " & Int(32767 * Cos((GlideSlope + 1) * DegToRad)) & " " & Int(32767 * Sin((GlideSlope + 1) * DegToRad)) & " " & FloatFormat(-DistILS * Tan((GlideSlope + 1) * DegToRad), "##0.0#######") & " )"
  Print #FileNum, "    LineColor( 15 F0 )"
  Print #FileNum, ":Line1" & Suffix
  Print #FileNum, "    MoveTo( " & Int(-HalfWidth - 4) & " 0 " & CInt(-DistILS) & " )"
  Print #FileNum, "    DrawTo( " & Int(-HalfWidth - 9) & " 0 " & CInt(-DistILS) & " )"
  
  Print #FileNum, "    LineColor( 0F F0 )"
  Print #FileNum, "    VectorJump( :Line2" & Suffix & " m 0 " & Int(32767 * Cos((GlideSlope + 0.5) * DegToRad)) & " " & Int(32767 * Sin((GlideSlope + 0.5) * DegToRad)) & " " & FloatFormat(-DistILS * Tan((GlideSlope + 0.5) * DegToRad), "##0.0#######") & " )"
  Print #FileNum, "    LineColor( 15 F0 )"
  Print #FileNum, ":Line2" & Suffix
  Print #FileNum, "    MoveTo( " & Int(-HalfWidth - 12) & " 0 " & CInt(-DistILS) & " )"
  Print #FileNum, "    DrawTo( " & Int(-HalfWidth - 17) & " 0 " & CInt(-DistILS) & " )"
 
  Print #FileNum, "    LineColor( 0F F0 )"
  Print #FileNum, "    VectorJump( :Line3" & Suffix & " m 0 " & Int(32767 * Cos(GlideSlope * DegToRad)) & " " & Int(32767 * Sin(GlideSlope * DegToRad)) & " " & FloatFormat(-DistILS * Tan(GlideSlope * DegToRad), "##0.0#######") & " )"
  Print #FileNum, "    LineColor( 15 F0 )"
  Print #FileNum, ":Line3" & Suffix
  Print #FileNum, "    MoveTo( " & Int(-HalfWidth - 20) & " 0 " & CInt(-DistILS) & " )"
  Print #FileNum, "    DrawTo( " & Int(-HalfWidth - 25) & " 0 " & CInt(-DistILS) & " )"
  
  Print #FileNum, "    LineColor( 0F F0 )"
  Print #FileNum, "    VectorJump( :Line4" & Suffix & " m 0 " & Int(32767 * Cos((GlideSlope - 0.5) * DegToRad)) & " " & Int(32767 * Sin((GlideSlope - 0.5) * DegToRad)) & " " & FloatFormat(-DistILS * Tan((GlideSlope - 0.5) * DegToRad), "##0.0#######") & " )"
  Print #FileNum, "    LineColor( 15 F0 )"
  Print #FileNum, ":Line4" & Suffix
  Print #FileNum, "    MoveTo( " & Int(-HalfWidth - 28) & " 0 " & CInt(-DistILS) & " )"
  Print #FileNum, "    DrawTo( " & Int(-HalfWidth - 33) & " 0 " & CInt(-DistILS) & " )"
  
  Print #FileNum, ":Finish" & Suffix
  Print #FileNum, "    Return"
  Print #FileNum, ":EndPAPI" & Suffix
End Sub

