Attribute VB_Name = "basConversion"
' [basConversion]
' Unit conversion functions
Option Explicit

' Color conversion
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Const CLR_INVALID = -1

' Conversion constants
Public Const MToFt = 3.28
Public Const MToKm = 0.001
Public Const MToNm = 1 / 1852
Public Const MToMi = MToFt / 5280

Public Const FtToM = 1 / MToFt
Public Const KmToM = 1 / MToKm
Public Const NmToM = 1 / MToNm
Public Const MiToM = 1 / MToMi

Public Const FtToNm = FtToM * MToNm
Public Const KmToNm = KmToM * MToNm
Public Const MiToNm = MiToM * MToNm

Public Const PI = 3.14159265358979

Public Const DegToRad = PI / 180
Public Const RadToDeg = 180 / PI

Public Function Arccos(ByVal X As Double) As Double
  ' Inverse Cosine Arccos(X) = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
  If Abs(X) >= 1 Then
    Arccos = PI
  Else
    Arccos = Atn(-X / Sqr(1 - X * X)) + PI / 2
  End If
End Function

' Returns the arctangent of y/x in degrees
Public Function Atn2(ByVal Y As Single, ByVal X As Single) As Single
  Dim Ans As Single
  Select Case X
    Case Is > 0
      Ans = Atn(Y / X)
    Case Is = 0
      Ans = IIf(Y >= 0, PI / 2, 3 * PI / 2)
    Case Is < 0
      Ans = Atn(Y / X) + PI
  End Select
  If Ans < 0 Then Ans = Ans + 2 * PI
  Atn2 = Ans
End Function

' Add/subtract a value from a rotation and ensure that
' it is 0 <= rotation < 360
Public Function EnsureRotation(ByVal Rotation As Single) As Single
  Do
    If Rotation >= 360 Then
      Rotation = Rotation - 360
    ElseIf Rotation < 0 Then
      Rotation = Rotation + 360
    Else
      Exit Do
    End If
  Loop
  EnsureRotation = Rotation
End Function

' Converts to magnetic if user desires
Public Function GeographicToUser(ByVal Rotation As Single) As String
  If Options.Magnetic Then
    GeographicToUser = Append(EnsureRotation(Rotation + Scenery.Header.MagVar), RES_Unit_AbbrevMag, "##0.0")
  Else
    GeographicToUser = Append(Rotation, RES_Unit_AbbrevGeo, "##0.0")
  End If
End Function

' Makes a PointAPI from X and Y
Public Function MakeAPIPoint(ByVal X As Single, ByVal Y As Single) As POINTAPI
  MakeAPIPoint.X = CLng(X)
  MakeAPIPoint.Y = CLng(Y)
End Function

' Makes a PointType from X and Y
Public Function MakePoint(ByVal X As Single, ByVal Y As Single) As PointType
  MakePoint.X = X
  MakePoint.Y = Y
End Function

' Convert the given string with units into meters
Public Function Meter(ByVal Data As String) As Single
  Select Case UCase$(GetUnit(Data))
    Case "M", "MAGL", "M AGL"
      Meter = ValEx(Data)
    Case "FT", "FTAGL", "FT AGL"
      Meter = ValEx(Data) * FtToM
    Case "MMSL", "M MSL"
      Meter = ValEx(Data) - Scenery.Header.Altitude
    Case "FTMSL", "FT MSL"
      Meter = ValEx(Data) * FtToM - Scenery.Header.Altitude
    Case "KM"
      Meter = ValEx(Data) * KmToM
    Case "NM"
      Meter = ValEx(Data) * NmToM
    Case Else
      Meter = ValEx(Data)
  End Select
End Function

' Converts to feet if user desires
Public Function MeterToUser(ByVal Meter As Single, Optional ByVal FormStr As String) As String
  If Options.Metric Then
    MeterToUser = Append(Meter, RES_Unit_AbbrevM, FormStr)
  Else
    MeterToUser = Append(Meter * MToFt, RES_Unit_AbbrevFt, FormStr)
  End If
End Function

' Rotates Pts by Rotate degrees
Public Sub MultiRotate(Pts() As PointType, ByVal Deg As Single)
  Dim I As Integer
  For I = 0 To UBound(Pts())
    Rotate Pts(I), Deg
  Next I
End Sub

' Convert the string with units into nautical miles
Public Function Nautical(ByVal Data As String) As Single
  Select Case UCase$(GetUnit(Data))
    Case "NM"
      Nautical = ValEx(Data)
    Case "KM"
      Nautical = ValEx(Data) * KmToNm
    Case "M"
      Nautical = ValEx(Data) * MToNm
    Case "FT"
      Nautical = ValEx(Data) * FtToNm
    Case Else
      Nautical = ValEx(Data)
  End Select
End Function

' Appends a nautical unit. Provided for consistency
' with MeterToUser
Public Function NauticalToUser(ByVal Nautical As Single, Optional FormStr As String) As String
  NauticalToUser = Append(Nautical, RES_Unit_AbbrevNm, FormStr)
End Function

' Polar to Rectangular coordinates
Public Sub PolarToRect(ByVal R As Double, ByVal H As Single, ByRef cX As Single, ByRef cY As Single)
  cX = R * Cos(H * DegToRad)
  cY = R * Sin(H * DegToRad)
End Sub

' Rectangular to Polar coordinates
Public Sub RectToPolar(ByVal cX As Single, ByVal cY As Single, ByRef R As Double, ByRef H As Single)
  R = Distance(0, 0, cX, cY)
  H = Atn2(cY, cX) * RadToDeg
End Sub

' Uses the Scenery.Header.Center Point and calculates
' point at [X],[Y] meters
Public Function ReturnPoint(ByVal cX As Single, ByVal cY As Single) As clsLatLon
  Set ReturnPoint = Scenery.Header.Center.CalcPoint(Distance(cX, cY, 0, 0) * MToNm, EnsureRotation(90 - Atn2(cY, cX) * RadToDeg + Scenery.Header.Rotation))
End Function

' Rotates a point by Deg Degrees
Public Sub Rotate(Pt As PointType, ByVal Deg As Single)
  Dim R As Double, H As Single
  RectToPolar Pt.X, Pt.Y, R, H
  PolarToRect R, H - Deg, Pt.X, Pt.Y
End Sub

' Extension of the Val function, takes care of regional
' preferences
Public Function StrEx(ByVal Num As Single) As String
  StrEx = Trim$(Str$(Num))
End Function

' Translate an OLE_COLOR to RGB
Public Function TranslateColor(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0) As Long
  ' Convert Automation color to Windows color
  If OleTranslateColor(oClr, hPal, TranslateColor) Then
    TranslateColor = CLR_INVALID
  End If
End Function

' Converts to geographic if in magnetic
Public Sub UserToGeographic(ByVal User As String, ByRef Value As Single, ByRef ErrValue As Integer)
  Dim Unit As String, I As Integer
  If InStr(DigitsDecimalSigns, Left$(User, 1)) = 0 Or User = "" Then
    ' Non numeric first character
    ErrValue = RES_ERR_Numeric
  Else
    Unit = GetUnit(User)
    For I = 1 To Len(User) - Len(Unit)
      If InStr(DigitsDecimalSigns, Mid$(User, I, 1)) = 0 Then
        ErrValue = RES_ERR_Numeric
        Exit Sub
      End If
    Next I
    
    If Unit = "" Then
      Unit = Lang.GetString(IIf(Options.Magnetic, RES_Unit_AbbrevMag, RES_Unit_AbbrevGeo))
    End If
    
    Unit = Replace(Unit, "°", "")
    
    If StrComp(Replace(Lang.GetString(RES_Unit_AbbrevMag), "°", ""), Unit, vbTextCompare) = 0 Then
      Value = EnsureRotation(ValEx(User) - Scenery.Header.MagVar)
    ElseIf StrComp(Replace(Lang.GetString(RES_Unit_AbbrevGeo), "°", ""), Unit, vbTextCompare) = 0 Then
      Value = ValEx(User)
    Else
      ' Unrecognized label
      ErrValue = RES_ERR_RotUnits
    End If
  End If
End Sub

' Given a data string, return the unit of the string,
' and the conversion factor
Public Sub UserToMeter(ByVal User As String, ByRef UnitLabel As Integer, ByRef ConversionFactor As Single, ByRef ConversionCorrection As Single, ByRef ErrValue As Integer)
  Dim Unit As String, I As Integer
  If InStr(DigitsDecimalSigns, Left$(User, 1)) = 0 Or User = "" Then
    ' Non numeric first character
    ErrValue = RES_ERR_Numeric
  Else
    Unit = GetUnit(User)
    For I = 1 To Len(User) - Len(Unit)
      If InStr(DigitsDecimalSigns, Mid$(User, I, 1)) = 0 Then
        ErrValue = RES_ERR_Numeric
        Exit Sub
      End If
    Next I
    
    If Unit = "" Then _
      Unit = Lang.GetString(IIf(Options.Metric, RES_Unit_AbbrevM, RES_Unit_AbbrevFt))
    
    Select Case UCase$(Unit)
      Case UCase$(Lang.GetString(RES_Unit_AbbrevM))
        ConversionFactor = 1
        UnitLabel = RES_Unit_M
      Case UCase$(Lang.GetString(RES_Unit_AbbrevFt))
        ConversionFactor = MToFt
        UnitLabel = RES_Unit_Ft
      Case UCase$(Lang.GetString(RES_Unit_AbbrevNm))
        ConversionFactor = MToNm
        UnitLabel = RES_Unit_Nm
      Case UCase$(Lang.GetString(RES_Unit_AbbrevKm))
        ConversionFactor = MToKm
        UnitLabel = RES_Unit_Km
      Case UCase$(Lang.GetString(RES_Unit_AbbrevMi))
        ConversionFactor = MToMi
        UnitLabel = RES_Unit_Mi
      Case Else
        ' Unrecognized label
        ErrValue = RES_ERR_Units
    End Select
  End If
End Sub

' Given a data string, return the unit of the string,
' and the conversion factor
Public Sub UserToNautical(ByVal User As String, ByRef UnitLabel As Integer, ByRef ConversionFactor As Single, ByRef ConversionCorrection As Single, ByRef ErrValue As Integer)
  Dim Unit As String, I As Integer
  If InStr(DigitsDecimalSigns, Left$(User, 1)) = 0 Or User = "" Then
    ' Non numeric first character
    ErrValue = RES_ERR_Numeric
  Else
    Unit = GetUnit(User)
    For I = 1 To Len(User) - Len(Unit)
      If InStr(DigitsDecimalSigns, Mid$(User, I, 1)) = 0 Then
        ErrValue = RES_ERR_Numeric
        Exit Sub
      End If
    Next I
    
    If Unit = "" Then _
      Unit = Lang.GetString(RES_Unit_AbbrevNm)
    
    Select Case UCase$(Unit)
      Case UCase$(Lang.GetString(RES_Unit_AbbrevM))
        ConversionFactor = NmToM
        UnitLabel = RES_Unit_M
      Case UCase$(Lang.GetString(RES_Unit_AbbrevFt))
        ConversionFactor = NmToM * MToFt
        UnitLabel = RES_Unit_Ft
      Case UCase$(Lang.GetString(RES_Unit_AbbrevNm))
        ConversionFactor = 1
        UnitLabel = RES_Unit_Nm
      Case UCase$(Lang.GetString(RES_Unit_AbbrevKm))
        ConversionFactor = NmToM * MToKm
        UnitLabel = RES_Unit_Km
      Case UCase$(Lang.GetString(RES_Unit_AbbrevMi))
        ConversionFactor = NmToM * MToMi
        UnitLabel = RES_Unit_Mi
      Case Else
        ' Unrecognized label
        ErrValue = RES_ERR_Units
    End Select
  End If
End Sub

' Extension of the Val function, takes care of regional
' preferences
Public Function ValEx(ByVal Num As String) As Double
  Dim Y As Long, Sep As String
  Sep = LocaleInfo(LOCALE_SDECIMAL)
  Y = InStr(Num, Sep)
  If Y > 0 Then Mid$(Num, Y, 1) = "."
  ValEx = Val(Num)
End Function
