VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements clsObject

Public Author As String, _
       Copyright As String, _
       Name As String, _
       LangCode As String, _
       Region As String, _
       Country As String, _
       State As String, _
       City As String, _
       ICAOID As String, _
       Center As clsLatLon, _
       Rotation As Single, _
       Horz As Long, _
       Vert As Long, _
       Altitude As Single, _
       MagVar As Single, _
       Exclusion As Byte, _
       Base As Byte, _
       Size As Byte

Public APTMagVar As Single

Private Compass1(9) As PointType, Compass2(9) As PointType
       
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Property Get Caption() As String
  Caption = clsObject_Caption()
End Property

Public Sub EditProperties()
  clsObject_EditProperties
End Sub

Public Sub LoadFS2000APTData(FileData As clsAPTData)
  Dim intTemp As Integer
  With FileData
    If ValEx(.GetData("AFDBGLFile")) > 0 Then
      Scenery.AFDRefresh = True
      intTemp = .GetData("LanguageCode")
      Select Case intTemp
        Case 1
          LangCode = "Deutsch - 407"
        Case 2
          LangCode = "Français - 40C"
        Case 3
          LangCode = "Japanese - 411"
        Case Else ' 0, -1
          LangCode = "US English - 409"
      End Select
      Region = .GetData("AFDRegion")
      Country = .GetData("AFDCountry")
      State = .GetData("AFDState")
      City = .GetData("AFDCity")
      Scenery(Scenery.TowerIndex).LoadAPTData FileData
    Else
      Scenery.TowerIndex = 0
    End If
  End With
End Sub

Private Sub Class_Initialize()
  Dim Buffer As String
  Buffer = String$(260, 0)
  ' Set default values
  Set Center = New clsLatLon

  Name = UntitledName
  GetUserName Buffer, 260 + 1
  Buffer = StripTerminator(Buffer)
  If Buffer = "" Then
   Author = Lang.GetString(RES_Hdr_AuthorName)
  Else
    Author = Buffer
  End If
  Copyright = Lang.ResolveString(RES_Hdr_Copyright, Year(Date$), Author)
  Region = "AnyPlace"
  Country = "AnyCountry"
  City = "AnyTown"
  ICAOID = "ANY"
  Horz = 10000
  Vert = 10000
  LangCode = "US English - 409"
  MagVar = Center.MagVar()

  Compass1(0) = MakePoint(0, 0)
  Compass1(1) = MakePoint(8, 8)
  Compass1(2) = MakePoint(25, 0)
  Compass1(3) = MakePoint(-25, 0)
  Compass1(4) = MakePoint(-8, -8)
  Compass1(5) = MakePoint(0, 0)
  Compass1(6) = MakePoint(0, 40)
  Compass1(7) = MakePoint(-8, 8)
  Compass1(8) = MakePoint(8, -8)
  Compass1(9) = MakePoint(0, -40)
  
  Compass2(0) = MakePoint(0, 0)
  Compass2(1) = MakePoint(-8, 8)
  Compass2(2) = MakePoint(-25, 0)
  Compass2(3) = MakePoint(25, 0)
  Compass2(4) = MakePoint(8, -8)
  Compass2(5) = MakePoint(0, 0)
  Compass2(6) = MakePoint(0, 40)
  Compass2(7) = MakePoint(8, 8)
  Compass2(8) = MakePoint(-8, -8)
  Compass2(9) = MakePoint(0, -40)
End Sub

Private Sub Class_Terminate()
  Set Center = Nothing
'  If Closing Then Error 1
End Sub

Private Function clsObject_Add(ByVal NewX As Single, ByVal NewY As Single) As Boolean
  ' Nothing
End Function

Private Property Get clsObject_Caption(Optional ByVal Default As Boolean = False) As String
  clsObject_Caption = Name
End Property

Private Sub clsObject_Compile(ByVal FileNum As Integer)
  Dim TopLeft As clsLatLon, BottomRight As clsLatLon, _
    I As Integer, J As Integer, L As Integer, _
    CenX As Long, CenY As Long, SynString As String, _
    Scaling As Byte, LargeSize As Integer
  Dim objExcl As clsExclusion
  Dim Pt As PointType
  
  ' Make adjustment since CalcPoint automatically adjusts for rotation
  Pt = MakePoint(-Horz / 2, Vert / 2)
  Rotate Pt, -Scenery.Header.Rotation
  Set TopLeft = ReturnPoint(Pt.X, Pt.Y)
  
  Pt = MakePoint(Horz / 2, -Vert / 2)
  Rotate Pt, -Scenery.Header.Rotation
  Set BottomRight = ReturnPoint(Pt.X, Pt.Y)
  
  ' SCASM File info header
  Print #FileNum, "Set( buf 100 )"
  Print #FileNum, "Set( areamx 64 )"
  Print #FileNum, "Set( linbuf 64 )"
  Print #FileNum, "Set( ppperr "; IIf(Options.FSVersion >= Version_FS2K, 2000, 80) & " )"
'  Select Case Options.FSVersion
'    Case Version_FS95
'      Print #FileNum, "Set( FSVERS 0x600 )"
'    Case Version_FS98
'      Print #FileNum, "Set( FSVERS 0x610 )"
'    Case Version_CFS1
'      Print #FileNum, "Set( FSVERS 0x650 )"
'    Case Version_FS2K
'      Print #FileNum, "Set( FSVERS 0x700 )"
'    Case Version_CFS2
'      Print #FileNum, "Set( FSVERS 0x732 )"
'    Case Version_FS2K2
'      Print #FileNum, "Set( FSVERS 0x800 )"
'  End Select
  Print #FileNum, "UVar( $FSSCTEMP1 " & Options.FSVersion & " )"; vbCrLf
  
  ' Scenery Header
  Scenery.HeaderStr = "Header( 1 " & TopLeft.Latitude() & " " & BottomRight.Latitude() & " " & BottomRight.Longitude() & " " & TopLeft.Longitude() & " )" & vbCrLf & vbCrLf & "LatRange( " & BottomRight.Latitude() & " " & TopLeft.Latitude() & " )" & vbCrLf
  Scenery.AIHeaderStr = "AI_Border( " & TopLeft.Latitude() & " " & BottomRight.Latitude() & " " & BottomRight.Longitude() & " " & TopLeft.Longitude() & " )" & vbCrLf
  Print #FileNum, Scenery.HeaderStr
  
  For I = 1 To Scenery.Count
    If Scenery(I).ObjectType = OT_MenuEntry Then
      Print #FileNum, "Menu( "; ReplaceParens(Name); " )"; vbCrLf
      Exit For
    End If
  Next I
  
  If Copyright <> "" Then
    Print #FileNum, "Copyright( "; ReplaceParens(Copyright); " )"; vbCrLf
  End If
  
  If Exclusion > 0 Then
    Set objExcl = New clsExclusion
    With objExcl
      .Name = "Header Exclusion"
      .Exclusion = Exclusion
      .Horz = Horz
      .Vert = Vert
      .Compile FileNum
    End With
    Set objExcl = Nothing
  End If

  ' Synthetic Scenery
  If Size > 0 And Options.FSVersion <= Version_CFS1 Then
    Scaling = Center.SynScale
    CenX = Int(Center.SyntheticX(Scaling))
    CenY = Int(Center.SyntheticY(Scaling))

    L = Size - 1
    LargeSize = 256 * 2 ^ (Scaling - 1)

    If Base = 0 Then
      SynString = "01 00 0000"
    Else
      SynString = SynNames(Base).ID & " 00 1111"
    End If

    For I = CenY - L To CenY + L
      Print #FileNum, "Synth " & Scaling & " " & I
      For J = CenX - L To CenX + L
        Print #FileNum, "Block "; Int((J + LargeSize) Mod LargeSize) & " " & SynString & " " & CInt(Altitude)
      Next J
      Print #FileNum,
    Next I
    Print #FileNum,
  End If
End Sub

Private Property Get clsObject_Complexity() As Byte
  ' Nothing
End Property

Private Property Let clsObject_Complexity(ByVal RHS As Byte)
  ' Nothing
End Property

Private Sub clsObject_CopyTo(Destination As clsObject)
  ' Nothing
End Sub

Private Sub clsObject_DrawBottom(Optional ByVal DrawX As Single = -999999#, Optional ByVal DrawY As Single)

End Sub

Private Sub clsObject_DrawTop(Optional ByVal DrawX As Single = -999999#, Optional ByVal DrawY As Single)
  Dim I As Long, GridWidth As Integer
  Dim Left As Single, Top As Single, _
    Right As Single, Bottom As Single
  Dim TempLeft As Long, TempTop As Long, _
    TempRight As Long, TempBottom As Long
  Dim hPen As Long, oldPen As Long
  Dim Temp As Long
  
  Dim TopLeft As clsLatLon, BottomRight As clsLatLon
  Dim X1 As Single, X2 As Single, Y1 As Single, Y2 As Single
  Dim RetDist As Double, RetDeg As Single
  Dim Scaling As Byte, CenX As Integer, CenY As Integer
  
  Dim ScaleFactor As String
  ScaleFactor = picEditor.ScaleX(1)
  
  GridWidth = Options.Grid

  Options.SetGLColor COLOR_Gridline
  
  If GridWidth > 0 Then
    
    ' Create pen
    hPen = CreatePen(vbSolid, 0, Options.ObjectColor(COLOR_Gridline))
    ' Set the pen and brush
    oldPen = SelectObject(picEditor.hdc, hPen)
    
    Right = Horz / 2
    Left = -Right
    Top = Vert / 2
    Bottom = -Top
    
    glFlush
    
    With picEditor
      ' Max zoom level is 64 meters/pixel,
      ' so 100 scale units should be a good margin
      If Left < .myLeft Then Left = .myLeft - 100
      If Right > .myRight Then Right = .myRight + 100
      If Top > .myTop Then Top = .myTop + 100
      If Bottom < .myBottom Then Bottom = .myBottom - 100
    
      .ScaleToPixelLong Left, Top, TempLeft, TempTop
      .ScaleToPixelLong Right, Bottom, TempRight, TempBottom
    
      ' Draw rectangle around scenery area
      MoveToEx .hdc, TempLeft, TempTop, ByVal 0&
      LineTo .hdc, TempRight, TempTop
      LineTo .hdc, TempRight, TempBottom
      LineTo .hdc, TempLeft, TempBottom
      LineTo .hdc, TempLeft, TempTop
          
      ' Grid lines
      For I = Fix(Left / GridWidth) * GridWidth To Right Step GridWidth
        .ScaleToPixelLong I, 0, Temp, 0
        MoveToEx .hdc, Temp, TempTop, ByVal 0&
        LineTo .hdc, Temp, TempBottom
      Next I
      For I = Fix(Bottom / GridWidth) * GridWidth To Top Step GridWidth
        .ScaleToPixelLong 0, I, 0, Temp
        MoveToEx .hdc, TempLeft, Temp, ByVal 0&
        LineTo .hdc, TempRight, Temp
      Next I
    End With
        
    SelectObject picEditor.hdc, oldPen
    DeleteObject hPen

    If Size > 0 Then
      Scaling = Center.SynScale
      CenX = Int(Center.SyntheticX(Scaling))
      CenY = Int(Center.SyntheticY(Scaling))
      
      Set TopLeft = New clsLatLon
      Set BottomRight = New clsLatLon
      
      TopLeft.NumLatitude = Center.SyntheticYToLatitude(CenY - (Size - 1), Scaling)
      TopLeft.NumLongitude = Center.SyntheticXToLongitude(CenX + Size, Scaling)
      BottomRight.NumLatitude = Center.SyntheticYToLatitude(CenY + Size, Scaling)
      BottomRight.NumLongitude = Center.SyntheticXToLongitude(CenX - (Size - 1), Scaling)
      
      Center.CalcDistance TopLeft, RetDist, RetDeg
      PolarToRect RetDist * NmToM, 90 - RetDeg, X1, Y1
      
      Center.CalcDistance BottomRight, RetDist, RetDeg
      PolarToRect RetDist * NmToM, 90 - RetDeg, X2, Y2
      
      glPushMatrix
      glRotatef Scenery.Header.Rotation, 0, 0, 1
      Options.SetGLColor COLOR_SyntheticColor
      glLineWidth 3
      glDrawRect X1, Y1, X2, Y2
      glLineWidth 1
      glPopMatrix
      
      Set TopLeft = Nothing
      Set BottomRight = Nothing
    End If
  End If

  glDrawCircle 0, 0, FocusWidth, Options.ObjectColor(COLOR_Gridline)
  
  If Options.ShowCompass Then
    Dim A1(9) As PointType, B1(9) As PointType
    Dim A2() As TessType, B2() As TessType
    
    For I = 0 To 9
      A1(I).X = Compass1(I).X * ScaleFactor
      A1(I).Y = Compass1(I).Y * ScaleFactor
      B1(I).X = Compass2(I).X * ScaleFactor
      B1(I).Y = Compass2(I).Y * ScaleFactor
    Next I
      
    glTesselate A1, A2, True
    glTesselate B1, B2, True

    glPushMatrix
    glForeColor Options.ObjectColor(COLOR_Background)
    glTranslatef picEditor.myRight - 35 * ScaleFactor, picEditor.myBottom + 70 * ScaleFactor, 0
    glRotatef Scenery.Header.Rotation + MagVar, 0, 0, 1
    glDrawTesselated B2
    glForeColor Options.ObjectColor(COLOR_CompassColor)
    glDrawTesselated A2
    glDrawPolygon B1
    glPopMatrix
  End If
  glPopMatrix
End Sub

Private Function clsObject_EditProperties() As Boolean
  Dim OldRotation As Single
  OldRotation = Rotation
  If frmHeader.EditData(Me) Then
    If OldRotation <> Rotation Then
      Scenery.TransformScenery 0, 0, -(Scenery.Header.Rotation - OldRotation), True, False
    End If
    clsObject_EditProperties = True
  End If
End Function

Private Sub clsObject_LoadAPTData(FileData As clsAPTData)
  Dim I As Integer, BackgroundObject As clsBackground, _
    Temp As String, Temp2 As String
  
  With FileData
    .ReadSection
    Author = Replace(.GetData("Auteur"), "_", " ")
    Horz = Meter(.GetData("Range"))
    Vert = Horz
    Temp = .GetData("BGLDirectory")
    If LCase$(Right$(Temp, 8)) = "\scenery" Then
      Scenery.ExportPath = Left$(Temp, Len(Temp) - 8)
    Else
      Scenery.ExportPath = Temp
    End If
     
    ' Background image
    Temp = .GetData("BitmapFile")
    If Temp <> "" Then
      Set BackgroundObject = Scenery(1)
      With BackgroundObject
        .File = Temp
        Temp = FileData.GetData("BitmapScale")
        Temp2 = ReadNext(Temp, " ")
        .ZoomX = 100 / ValEx(Temp2)
        .ZoomY = .ZoomX
        .X = ValEx(ReadNext(Temp, " "))
        .Y = -ValEx(ReadNext(Temp, " "))
        .LockAspectRatio = True
        .ObjectIndex = 1
        .SceneryIndex = 1
        .mDrawable = True
      End With
      Scenery(1).UpdateObject
    End If
    
    .ReadSection
    Temp = .GetData("Latitude")
    If Mid$(Temp, 2, 1) = " " Then
      Temp2 = ReadNext(Temp, " ")
    Else
      Temp2 = ""
    End If
    Temp2 = Temp2 & CSng(ReadNext(Temp, " D ")) & " " & CSng(ReadNext(Temp, " M ")) & " " & CSng(ReadNext(Temp, " S"))
    Center.Latitude = Temp2
    
    Temp = .GetData("Longitude")
    If Mid$(Temp, 2, 1) = " " Then
      Temp2 = ReadNext(Temp, " ")
    Else
      Temp2 = ""
    End If
    Temp2 = CSng(ReadNext(Temp, " D ")) & " " & CSng(ReadNext(Temp, " M ")) & " " & CSng(ReadNext(Temp, " S"))
    Center.Longitude = Temp2

    .ReadSection
    Name = Replace(.GetData("Nom"), "_", " ")
    Altitude = Meter(.GetData("Altitude"))
    APTMagVar = ValEx(ReadNext(.GetData("CorMag"), " "))
    Temp = .GetData("Paysage")
    Temp2 = ReadNext(Temp, " ")
    If Temp2 = "NO" Then
      Size = 0
    Else
      For I = 0 To 26
        If SynNames(I).ID = Temp2 Then Base = I: Exit For
      Next I
      Temp2 = ReadNext(Temp, " ")
      Select Case Temp2
        Case 5
          Size = 3
        Case 3
          Size = 2
        Case Else ' 1
          Size = 1
      End Select
    End If

    Exclusion = ValEx(Trim$(Right$(.GetData("Exclusion"), 2)))

    Copyright = Lang.ResolveString(RES_Hdr_Copyright, Year(Date$), Author)
  End With
  MagVar = Center.MagVar()
End Sub

Private Sub clsObject_LoadBinaryData(ByVal FileNum As Integer, ByVal Version As Integer, ByVal UserData As Long)
  Dim valAltitude As Integer, valX As Integer, valY As Integer, _
    valDummy As Byte, SynString As String, FlattenNum As Integer
  Dim valHorz As Single, valVert As Single
  Dim valSize As Integer

  Dim ShowPic As Boolean, PicFile As String, _
    XPic As Single, YPic As Single, Scl As Single
  Dim BackgroundObject As clsBackground
  
  Select Case Version
    Case Is >= 10
      GetBinaryData FileNum, Name, Author, Copyright, _
        Region, Country, State, City, ICAOID, Center, _
        valHorz, valVert, Exclusion, Altitude, Base, _
        valSize, MagVar, LangCode
    Case Is >= 8
      GetBinaryData FileNum, Name, Author, Copyright, _
        Region, Country, State, City, ICAOID, Center, _
        valHorz, valVert, Exclusion, Altitude, Base, _
        valSize, MagVar
    Case Is >= 6
      GetBinaryData FileNum, Name, Author, Copyright, _
        Center, valHorz, valVert, Exclusion, Altitude, _
        ShowPic, PicFile, XPic, YPic, Scl, Base, _
        valSize, MagVar
    Case Is >= 3
      GetBinaryData FileNum, Name, Author, Copyright, _
        Center, valHorz, valVert, Exclusion, Altitude, _
        ShowPic, PicFile, XPic, YPic, Scl, Base, _
        valSize, MagVar, FlattenNum
'    Case Is >= 1
'      GetBinaryData FileNum, Name, Author, Copyright, _
'        Center, valHorz, valVert, Exclusion, valAltitude, _
'        ShowPic, PicFile, valX, valY, Scl, valDummy, _
'        SynString, valSize
'      Altitude = valAltitude
'      XPic = CSng(valX)
'      YPic = CSng(valY)
'    Case Else
'      GetBinaryData FileNum, Name, Author, Copyright, _
'        Center, valHorz, valVert, Exclusion, valAltitude, _
'        valDummy, SynString, Size
'      Altitude = valAltitude
  End Select
  Horz = valHorz
  Vert = valVert
  Size = valSize
  MagVar = Center.MagVar()

  If PicFile <> "" Then
    Set BackgroundObject = Scenery(1)
    With BackgroundObject
      .File = PicFile
      .X = XPic
      .Y = YPic
      .ZoomX = Scl
      .ZoomY = Scl
      .LockAspectRatio = True
      .ObjectIndex = 1
      .SceneryIndex = 1
      .mDrawable = True
    End With
    Scenery(1).UpdateObject
  End If
  If Size > 5 Then Size = 5
  If SynString <> "" Then Base = Val(Mid$(SynString, 12))
End Sub

Private Sub clsObject_LoadData(FileData As clsFileData, ByVal Version As Integer)
  With FileData
    .GetData "Author", Author, ""
    .GetData "Copyright", Copyright, ""
    .GetData "Name", Name, ""
    .GetData "LanguageCode", LangCode, ""
    .GetData "Region", Region, ""
    .GetData "Country", Country, ""
    .GetData "State", State, ""
    .GetData "City", City, ""
    .GetData "ICAO", ICAOID, ""
    Center.LoadData FileData
    .GetData "Rotation", Rotation, 0
    .GetData "Horizontal", Horz, 10000
    .GetData "Vertical", Vert, 10000
    .GetData "Altitude", Altitude, 0
    .GetData "MagVar", MagVar, 0
    .GetData "Exclusion", Exclusion, 0
    .GetData "Texture", Base, 0
    .GetData "SyntheticSize", Size, 0
  End With
  MagVar = Center.MagVar()
  clsObject_UpdateObject
End Sub

Private Property Get clsObject_Locked() As Boolean
  ' Nothing
End Property

Private Property Let clsObject_Locked(ByVal RHS As Boolean)
  ' Nothing
End Property

Private Property Get clsObject_Name() As String
  clsObject_Name = Name
End Property

Private Property Let clsObject_Name(ByVal RHS As String)
  Name = RHS
End Property

Private Property Get clsObject_ObjectIndex() As Integer
  ' Nothing
End Property

Private Property Let clsObject_ObjectIndex(ByVal RHS As Integer)
  ' Nothing
End Property

Private Property Get clsObject_ObjectType() As ObjectTypeEnum
  clsObject_ObjectType = OT_Header
End Property

Private Property Get clsObject_PositionX() As Single
  ' Nothing
End Property

Private Property Let clsObject_PositionX(ByVal RHS As Single)
  ' Nothing
End Property

Private Property Get clsObject_PositionY() As Single
  ' Nothing
End Property

Private Property Let clsObject_PositionY(ByVal RHS As Single)
  ' Nothing
End Property

Private Property Get clsObject_Rotation() As Single
  ' Nothing
End Property

Private Property Let clsObject_Rotation(ByVal RHS As Single)
  ' Nothing
End Property

Private Sub clsObject_SaveData(FileData As clsFileData)
  With FileData
    .InitializeForWrite ObjectNames(OT_Header), Caption, 35
    .SaveData "Author", Author, ""
    .SaveData "Copyright", Copyright, ""
    .SaveData "Name", Name, ""
    .SaveData "LanguageCode", LangCode, ""
    .SaveData "Region", Region, ""
    .SaveData "Country", Country, ""
    .SaveData "State", State, ""
    .SaveData "City", City, ""
    .SaveData "ICAO", ICAOID, ""
    Center.SaveData FileData
    .SaveData "Rotation", Rotation, 0
    .SaveData "Horizontal", Horz, 10000
    .SaveData "Vertical", Vert, 10000
    .SaveData "Altitude", Altitude, 0
    .SaveData "MagVar", MagVar, 0
    .SaveData "Exclusion", Exclusion, 0
    .SaveData "Texture", Base, 0
    .SaveData "SyntheticSize", Size, 0
    .WriteSection
  End With
End Sub

Private Property Get clsObject_SceneryIndex() As Integer
  ' Nothing
End Property

Private Property Let clsObject_SceneryIndex(ByVal RHS As Integer)
  ' Nothing
End Property

Private Property Get clsObject_SortIndex() As Long
  clsObject_SortIndex = &H8000& * -3
End Property

Private Sub clsObject_UpdateObject()

End Sub

Private Sub clsObject_VersionChanged(ByVal NewVersion As Integer)
  If NewVersion > Version_FS98 Then
    Base = 0
  Else
    If Altitude < 0 Then Altitude = 0
  End If
End Sub
