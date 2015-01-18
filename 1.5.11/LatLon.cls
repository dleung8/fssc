VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLatLon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Latitude/Longitude Class
' Stores Latitude and Longitude

Public NumLatitude As Double, NumLongitude As Double

' Uses the current point and calculates the distance
' and heading to the given second point.
Public Sub CalcDistance(Sec As clsLatLon, ByRef Distance As Double, ByRef Angle As Single)
  Dim secLatitude As Double, secLongitude As Double
  
  secLatitude = Sec.NumLatitude
  secLongitude = Sec.NumLongitude
  
  If secLatitude = NumLatitude Then
    ' The angle is 90 or 270, tangent is undefined
    Angle = IIf(secLongitude <= NumLongitude, 270, 90)
    Distance = (secLongitude - NumLongitude) * 60 * Cos(NumLatitude * DegToRad)
  Else
    Angle = Atn2(Cos((NumLatitude + secLatitude) / 2 * DegToRad) * (secLongitude - NumLongitude), (secLatitude - NumLatitude)) * RadToDeg
    Distance = 60 * (secLatitude - NumLatitude) / Cos(Angle * DegToRad)
  End If
  Angle = Angle - Scenery.Header.Rotation
End Sub

' Uses the current point and calculates point at
' [distance] nautical miles at [heading]
Public Function CalcPoint(ByVal Distance As Double, ByVal Heading As Single) As clsLatLon
  Dim DeltaLat As Double, Ans As clsLatLon, myLat As Double
  Set Ans = New clsLatLon
  DeltaLat = Distance * Cos(Heading * DegToRad)
  Ans.NumLatitude = NumLatitude + (DeltaLat / 60)
  Ans.NumLongitude = NumLongitude + Distance * Sin(Heading * DegToRad) / Cos((NumLatitude + DeltaLat / 120) * DegToRad) / 60
  Set CalcPoint = Ans
End Function

' Simulates = operator
Public Sub CopyTo(Destination As clsLatLon)
  With Destination
    .NumLatitude = NumLatitude
    .NumLongitude = NumLongitude
  End With
End Sub

' Splices the Res string setting Data to a
' numeric value
Private Sub Cut(ByRef Res As String, ByRef Data As String)
  Dim I As Integer, Y As Long
  
  For I = 1 To Len(Res)
    If InStr(DigitsDecimal, Mid$(Res, I, 1)) = 0 Then
      Y = I
      Data = Left$(Res, Y - 1)
      Exit For
    End If
  Next I
  If Y = 0 Then
    Data = Res
    Res = ""
    Exit Sub
  End If

  For I = Y To Len(Res)
    If InStr(DigitsDecimal, Mid$(Res, I, 1)) > 0 Then
      Res = Mid$(Res, I)
      Exit Sub
    End If
  Next I
  Res = ""
End Sub

' Gets the magnetic variation of current point
Public Function MagVar() As Single
  Dim MagV As Single, FileNum As Integer, _
    Data As Integer, File As String
  
  On Error Resume Next

  If Not Options.MagVarMissing Then
    File = AddDir(Options.FSPath, "Scenery\MagDec.bgl")
    If FileExists(File) Then
      If FileLen(File) = 0 Then
        Kill File
        File = AddDir(Options.FSPath, "Scenery\Base\Scenery\MagDec.bgl")
      End If
    Else
      File = AddDir(Options.FSPath, "Scenery\Base\Scenery\MagDec.bgl")
    End If
  
    If FileExists(File) Then
      FileNum = FreeFile
      Open File For Binary As #FileNum
      If Err > 0 Then
        Options.MagVarMissing = True
        If frmSplash.isShown Then frmSplash.Hide
        MsgBoxEx frmMain, Lang.GetString(RES_ERR_NoMagVar), vbCritical, RES_ERR_NoMagVar
        If frmSplash.isShown Then frmSplash.Show
      Else
        Get #FileNum, (CLng(NumLatitude) + 90 + 181 * (CLng(NumLongitude + 360) Mod 360)) * 2 + 137, Data
        MagV = (360 * CLng(Data)) / 65536
        If MagV > 180 Then MagV = MagV - 360
        MagVar = Round(-MagV, 2)
        Options.MagVarMissing = False
        Close #FileNum
      End If
    Else
      Options.MagVarMissing = False
    End If
  End If
End Function

' Returns the String representation of Latitude
Public Property Get Latitude() As String
  Dim LatD As Integer, LatM As Integer, _
    LatS As Single, Temp As Single
  LatD = Int(Abs(NumLatitude))
  Temp = (Abs(NumLatitude) - LatD) * 60
  LatM = Int(Temp)
  LatS = (Temp - LatM) * 60
  
  Latitude = IIf(NumLatitude >= 0, "N", "S") & Format$(LatD, "00") & ":" & Format$(LatM, "00") & ":" & FloatFormat(LatS, "00.0000")
End Property

' Reads string Latitude
Public Property Let Latitude(ByVal Res As String)
  Dim X As Single, Data As String, _
    North As Boolean, TempVal As Double
  On Error Resume Next
  Res = Trim$(Res)
  
  Data = Left$(Res, 1)
  If StrComp(Data, "s", vbTextCompare) = 0 Or Data = "-" Then
    Res = Mid$(Res, 2)
  ElseIf StrComp(Data, "n", vbTextCompare) = 0 Or Data = "+" Then
    North = True
    Res = Mid$(Res, 2)
  End If
  
  Cut Res, Data
  If Data = "" Then Err.Raise 5
  NumLatitude = ValEx(Data)
  Cut Res, Data
  X = ValEx(Data)
  If X >= 60 Then Err.Raise 5
  NumLatitude = NumLatitude + Int(X) / 60
  If Int(X) <> X Then
    TempVal = (X - Int(X) + 0.005) / 60
    If TempVal >= 1 Then Err.Raise 5
  Else
    Cut Res, Data
    X = ValEx(Data)
    If X >= 60 Then Err.Raise 5
    TempVal = X / 3600
  End If
  If NumLatitude >= 89.5 Then Err.Raise 5
  NumLatitude = IIf(North, 1, -1) * (NumLatitude + TempVal)
  If Err.Number > 0 Then NumLatitude = -999
End Property

' Returns the String representation of Latitude
' formating it using user options
Public Property Get LatitudeUser() As String
  Dim LatD As Integer, LatM As Single, _
    LatS As Single, Temp As Single
  LatD = Int(Abs(NumLatitude))
  
  If Options.ShowFractionalMinutes Then
    ' Default
    LatM = Round((Abs(NumLatitude) - LatD) * 60 - 0.005, 2)
    LatitudeUser = IIf(NumLatitude >= 0, "N", "S") & Format$(LatD, "00") & "* " & FloatFormat(LatM, "00.00") & "'"
  Else
    ' Decimal minutes
    Temp = (Abs(NumLatitude) - LatD) * 60
    LatM = Int(Temp)
    LatS = (Temp - LatM) * 60
    LatitudeUser = IIf(NumLatitude >= 0, "N", "S") & Format$(LatD, "00") & "* " & Format$(LatM, "00") & "' " & FloatFormat(LatS, "00.0000") & Chr$(34)
  End If
End Property

' Return Latitude + Longitude
Public Function LatLon() As String
  LatLon = Latitude & " " & Longitude
End Function

' Returns the String representation of Latitude and
' Longitude formating for config file
Public Function LatLonConfig() As String
  Dim SLat As String, SLon As String, _
    D As Integer, M As Single
    
  D = Int(Abs(NumLatitude))
  M = Round((Abs(NumLatitude) - D) * 60 - 0.005, 2)
  SLat = IIf(NumLatitude >= 0, "N", "S") & Format$(D, "00") & " " & FloatFormat(M, "00.00")
  
  D = Int(Abs(NumLongitude))
  M = Round((Abs(NumLongitude) - D) * 60 - 0.005, 2)
  SLon = IIf(NumLongitude >= 0, "E", "W") & Format$(D, "00") & " " & FloatFormat(M, "00.00")
  
  LatLonConfig = SLat & ", " & SLon
End Function

' Loads a clsLatLon from an Airport file
Public Sub LoadAPTData(FileData As clsAPTData)
  Dim Temp As String
  With FileData
    Temp = .GetData("Latitude")
    Latitude = ReadNext(Temp, " ") & CSng(ReadNext(Temp, " D ")) & " " & CSng(ReadNext(Temp, " M ")) & " " & CSng(ReadNext(Temp, " S"))
    
    Temp = .GetData("Longitude")
    Longitude = ReadNext(Temp, " ") & CSng(ReadNext(Temp, " D ")) & " " & CSng(ReadNext(Temp, " M ")) & " " & CSng(ReadNext(Temp, " S"))
  End With
End Sub

' Loads a clsLatLon from a binary file
Public Sub LoadBinaryData(ByVal FileNum As Integer)
  Dim Direction As Byte, D As Byte, _
    M As Byte, S As Single
  Get #FileNum, , Direction
  
  Get #FileNum, , D
  Get #FileNum, , M
  Get #FileNum, , S
  NumLatitude = IIf((Direction And 1) > 0, 1, -1) * (D + M / 60 + S / 3600)
  
  Get #FileNum, , D
  Get #FileNum, , M
  Get #FileNum, , S
  NumLongitude = IIf((Direction And 2) > 0, 1, -1) * (D + M / 60 + S / 3600)
End Sub

' Loads a clsLatLon from a file
Public Sub LoadData(FileData As clsFileData)
  Dim Temp1 As String, Temp2 As String
  With FileData
    .GetData "Latitude", Temp1, "N0"
    .GetData "Longitude", Temp2, "E0"
  End With
  Latitude = Temp1
  Longitude = Temp2
End Sub

' Returns the String representation of Longitude
Public Property Get Longitude() As String
  Dim LonD As Integer, LonM As Integer, _
    LonS As Single, Temp As Single
  LonD = Int(Abs(NumLongitude))
  Temp = (Abs(NumLongitude) - LonD) * 60
  LonM = Int(Temp)
  LonS = (Temp - LonM) * 60
  
  Longitude = IIf(NumLongitude >= 0, "E", "W") & Format$(LonD, "000") & ":" & Format$(LonM, "00") & ":" & FloatFormat(LonS, "00.0000")
End Property

' Reads string Longitude
Public Property Let Longitude(ByVal Res As String)
  Dim X As Single, Data As String, _
    East As Boolean, TempVal As Double
  On Error Resume Next
  Res = Trim$(Res)
  
  Data = Left$(Res, 1)
  If StrComp(Data, "w", vbTextCompare) = 0 Or Data = "-" Then
    Res = Mid$(Res, 2)
  ElseIf StrComp(Data, "e", vbTextCompare) = 0 Or Data = "+" Then
    East = True
    Res = Mid$(Res, 2)
  End If

  Cut Res, Data
  If Data = "" Then Err.Raise 5
  NumLongitude = ValEx(Data)
  Cut Res, Data
  X = ValEx(Data)
  If X >= 60 Then Err.Raise 5
  NumLongitude = NumLongitude + Int(X) / 60
  If Int(X) <> X Then
    TempVal = (X - Int(X) + 0.005) / 60
    If TempVal >= 1 Then Err.Raise 5
  Else
    Cut Res, Data
    X = ValEx(Data)
    If X >= 60 Then Err.Raise 5
    TempVal = X / 3600
  End If
  If NumLongitude >= 180 Then Err.Raise 5
  NumLongitude = IIf(East, 1, -1) * (NumLongitude + TempVal)
  If Err.Number > 0 Then NumLongitude = -999
End Property

' Returns the String representation of Longitude
' formating it using user options
Public Property Get LongitudeUser() As String
  Dim LonD As Integer, LonM As Single, _
    LonS As Single, Temp As Single
  LonD = Int(Abs(NumLongitude))
  
  If Options.ShowFractionalMinutes Then
    ' Default
    LonM = Round((Abs(NumLongitude) - LonD) * 60 - 0.005, 2)
    LongitudeUser = IIf(NumLongitude >= 0, "E", "W") & Format$(LonD, "000") & "* " & FloatFormat(LonM, "00.00") & "'"
  Else
    ' Decimal minutes
    Temp = (Abs(NumLongitude) - LonD) * 60
    LonM = Int(Temp)
    LonS = (Temp - LonM) * 60
    LongitudeUser = IIf(NumLongitude >= 0, "E", "W") & Format$(LonD, "000") & "* " & Format$(LonM, "00") & "' " & FloatFormat(LonS, "00.0000") & Chr$(34)
  End If
End Property

' Saves clsLatLon into a file
Public Sub SaveData(FileData As clsFileData)
  With FileData
    .SaveData "Latitude", Latitude, ""
    .SaveData "Longitude", Longitude, ""
  End With
End Sub

' Helper function of SyntheticY
Private Function SynRecurse(ByVal Band As Integer) As Double
  ' Cache
  Static LastBand As Long, LastAns As Double
  Dim R As Double
  Const Base = 1.4086402
  
  If Band = 0 Then
    SynRecurse = 0
  ElseIf Band = 1 Then
    SynRecurse = Base
  ElseIf Band = LastBand Then
    SynRecurse = LastAns
  Else
    R = SynRecurse(Band - 1)
    LastAns = R + Base * Cos(R * DegToRad)
    LastBand = Band
    SynRecurse = LastAns
  End If
End Function

' Determine which size block the scenery should use
Public Function SynScale() As Byte
  Select Case Abs(NumLatitude)
    Case Is <= 60: SynScale = 6
    Case Is <= 75: SynScale = 5
    Case Is <= 82: SynScale = 4
    Case Is <= 85: SynScale = 3
    Case Is <= 87: SynScale = 2
    Case Else: SynScale = 1
  End Select
End Function

' Return the size of the synthetic block in meters
Public Function SynSize() As Long
  SynSize = 156543 / 2 ^ (SynScale - 1) * Cos(NumLatitude * DegToRad)
End Function

' Determine the X coordinate of the synthetic scenery
Public Function SyntheticX(ByVal Scaling As Byte) As Long
  Dim Temp As Double
  Temp = NumLongitude
  If Temp < 0 Then Temp = Temp + 360
  SyntheticX = Int(Temp / 360 * 256 * 2 ^ (Scaling - 1))
End Function

' Determine the longitude of the given synthetic block
Public Function SyntheticXToLongitude(ByVal XBlock As Long, ByVal Scaling As Byte) As Double
  Dim Temp As Double
  Temp = 360 / 256 * XBlock / 2 ^ (Scaling - 1)
  If Temp >= 180 Then Temp = Temp - 360
  SyntheticXToLongitude = Temp
End Function

' Determine the Y coordinate of the synthetic scenery
Public Function SyntheticY(ByVal Scaling As Byte) As Long
  Dim I As Integer, Low As Double, High As Double, _
      LowBase As Integer, Lat As Double
  For I = 0 To 256
    High = SynRecurse(I)
    If Abs(NumLatitude) < High Then
      LowBase = I - 1
      Exit For
    Else
      Low = High
    End If
  Next I
    
  For I = 0 To 32
    If (Low + ((High - Low) / (2 ^ (Scaling - 1)) * I)) > Abs(NumLatitude) Then SyntheticY = Sgn(NumLatitude) * (LowBase * (2 ^ (Scaling - 1)) + I - 1) + (NumLatitude < 0): Exit For
  Next I
End Function

' Determine the latitude of the given synthetic block
Public Function SyntheticYToLatitude(ByVal YBlock As Long, ByVal Scaling As Byte) As Double
  Dim Low As Double, High As Double, _
    Block As Long, LargeBlock As Long, SmallBlock As Long
  
  Block = Abs(YBlock - (YBlock < 0))
  LargeBlock = Int(Block / 2 ^ (Scaling - 1))
  SmallBlock = Block - LargeBlock * 2 ^ (Scaling - 1) - (YBlock < 0)
  
  Low = SynRecurse(LargeBlock)
  High = SynRecurse(LargeBlock + 1)
  
  SyntheticYToLatitude = Sgn(YBlock) * (Low + ((High - Low) / (2 ^ (Scaling - 1)) * SmallBlock))
End Function

' Returns the String representation of Latitude and
' Longitude formating it using user options
Public Function ToString() As String
  ToString = LatitudeUser & "  " & LongitudeUser
End Function

' Returns an error message if any
Public Function Validate(ByRef Msg As String) As Boolean
  If NumLatitude < -990 Or NumLongitude < -990 Then
    ' Lat, Lon = -999 is a parse error flag
    Msg = Lang.GetString(RES_ERR_LatParse)
  ElseIf Not Between(Int(Abs(NumLatitude)), 0, 89) Then
    Msg = Lang.FormatErrorMessage(RES_LBL_Latitude, 0, 89, RES_Unit_Deg)
  ElseIf NumLatitude >= 89.5 Then
    Msg = Lang.GetString(RES_ERR_Poles)
  ElseIf Not Between(Int(Abs(NumLongitude)), 0, 179) Then
    Msg = Lang.FormatErrorMessage(RES_LBL_Longitude, 0, 179, RES_Unit_Deg)
  End If
  Validate = (Msg = "")
End Function
