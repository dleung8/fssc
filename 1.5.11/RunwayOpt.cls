VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRunwayOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' RunwayOpt Class
' Provides functions to supplant the runway in the program

Public ThrLength As Integer, _
       ExtLength As Integer, _
       ExtSurface As Byte, _
       ThrLights As Byte, _
       ApprLights As Byte, _
       NumStrobes As Byte, _
       VASILeft As Byte, _
       VASIRight As Byte, _
       PAPI As Byte, _
       GlideSlope As Single, _
       Automatic As Boolean, _
       HDistance As Single, _
       VDistance As Single, _
       BarSpacing As Single, _
       SignHSpacing As Integer
       
Public IM As Single, _
       MM As Single, _
       OM As Single, _
       ILSEnabled As Boolean, _
       ILSName As String, _
       ILSID As String, _
       ILSFrequency As Single, _
       ILSRange As Byte, _
       ILSFlags As Byte, _
       ILSBeamWidth As Single

Public Parent As clsRunway

Public Function Compare(Other As clsRunwayOpt) As Boolean
  With Other
    Compare = .IM = IM Or _
              .MM = MM Or _
              .OM = OM Or _
              .ILSEnabled = ILSEnabled Or _
              .ILSName = ILSName Or _
              .ILSID = ILSID Or _
              .ILSFrequency = ILSFrequency Or _
              .ILSFlags = ILSFlags
  End With
End Function

Public Sub Compile(ByVal FileNum As Integer, ByVal Suffix As String)
  If ThrLights > 0 Then Print #FileNum, "            ThrLights" & Suffix & "  " & ThrLights
  If NumStrobes > 0 Then Print #FileNum, "            Strobes" & Suffix & "    " & NumStrobes
  If ThrLength > 0 Then Print #FileNum, "            ThrOff" & Suffix & "     " & CInt(ThrLength * MToFt)
  If ExtSurface <> 99 Then
    If ExtLength > 0 Then Print #FileNum, "            OverRun" & Suffix & "    " & ExtSurface & " " & CInt(ExtLength * MToFt) & " " & CInt(Parent.Width * MToFt)
  Else
    If ExtLength > 0 Then Print #FileNum, "            Ext" & Suffix & "        " & CInt(ExtLength * MToFt)
  End If
  If ApprLights > 0 Then Print #FileNum, "            AprL" & Suffix & "       " & ApprLights
  If ThrLights > 0 Then
    If Automatic Then
      Dim Num1 As Integer, Num2 As Integer
      Select Case VASILeft
        Case 0: Num1 = 0
        Case 7, 8, 9, 10: Num1 = 1
        Case 1, 3, 5: Num1 = 2
        Case 2, 4, 6: Num1 = 3
      End Select
      
      If Options.FSVersion >= Version_FS2K Then
        Select Case VASIRight
          Case 0: Num2 = 0
          Case 7, 8, 9, 10: Num2 = 1
          Case 1, 3, 5: Num2 = 2
          Case 2, 4, 6: Num2 = 3
        End Select
      Else
        Select Case VASIRight
          Case 0: Num2 = 0
          Case 1: Num2 = 1
        End Select
      End If
      
      If Num2 > Num1 Then
        Num1 = Num2
      End If

      HDistance = Parent.Width / 2 + 15
      Select Case Num1
        Case 0
        Case 1
          VDistance = Parent.Length / 2 - ThrLength - 250
          BarSpacing = 0
        Case 2
          VDistance = Parent.Length / 2 - ThrLength - 250
          BarSpacing = 215
        Case 3
          VDistance = Parent.Length / 2 - ThrLength - 300
          BarSpacing = 130
      End Select
    End If
    If VASILeft > 0 Or VASIRight > 0 Then
      Print #FileNum, "            VasiSlope" & Suffix & "  " & FloatFormat(GlideSlope, "0.0")
      Print #FileNum, "            VasiSide" & Suffix & "   " & CInt(HDistance * MToFt)
      Print #FileNum, "            VasiDist" & Suffix & "   " & CInt(VDistance * MToFt)
      Print #FileNum, "            VasiSpace" & Suffix & "  " & CInt(BarSpacing * MToFt)
      Print #FileNum, "            VasiBars" & Suffix & "   " & Hex$(VASILeft + VASIRight * 16)
    End If
  End If
  If SignHSpacing > 0 Then Print #FileNum, "            Distance" & Suffix & "   " & "0 " & CInt(SignHSpacing * MToFt)
End Sub

Public Function CompileBeacons(ByVal myPos As clsLatLon, ByVal ForwardRotation As Single, ByVal BackwardRotation As Single, ByVal ForwardDist As Single, ByVal BackwardDist As Single) As String
  Dim TempStr As String, strBuffer As String, _
    AFDBuffer As String
  Dim ILSLoc As String, ILSGP As String
  Dim Dist As Single
    
  ILSLoc = myPos.CalcPoint(BackwardDist * MToNm, ForwardRotation).LatLon
  ILSGP = myPos.CalcPoint(ForwardDist * MToNm, BackwardRotation).LatLon
  
  Dist = Parent.Length / 2 * MToNm

  ' SCASM, FreeSC incompatibility
  If Scenery.QuoteStrings Then
    TempStr = QuoteString(ILSName)
  Else
    TempStr = ReplaceParens(Replace(ILSName, " ", "_"))
  End If

  If ILSEnabled Then
    If Options.FSVersion >= Version_FS2K Then
      strBuffer = strBuffer & "ILS2( " & FloatFormat(ILSFrequency, "000.00") & " " & CInt(ILSRange * NmToM * MToKm) & " " & Hex$(ILSFlags) & " " & FloatFormat(Scenery.Header.Center.MagVar(), "##0.0#") & " " & _
                              ILSID & " " & TempStr & " " & _
                              ILSLoc & " " & CInt(Scenery.Header.Altitude) & " " & FloatFormat(ForwardRotation, "##0.0#") & " " & _
                              ILSGP & " " & CInt(Scenery.Header.Altitude) & " " & FloatFormat(GlideSlope, "0.0") & " " & _
                              FloatFormat(ILSBeamWidth, "0.0") & " )" & vbCrLf
    Else
      strBuffer = strBuffer & "ILS( " & FloatFormat(ILSFrequency, "000.00") & " " & CInt(ILSRange * NmToM * MToKm) & " " & Hex$(ILSFlags) & " " & FloatFormat(Scenery.Header.Center.MagVar(), "##0.0#") & " " & _
                              ILSID & " " & TempStr & " " & _
                              ILSLoc & " " & CInt(Scenery.Header.Altitude) & " " & FloatFormat(ForwardRotation, "##0.0#") & " " & _
                              ILSGP & " " & CInt(Scenery.Header.Altitude) & " " & FloatFormat(GlideSlope, "0.0") & " )" & vbCrLf
    End If
    
    If (ILSFlags And 1) > 0 Then
      ' DME set
      strBuffer = strBuffer & "DMEPos( " & FloatFormat(ILSFrequency, "000.00") & " 1 "
      If (ILSFlags And 16) Then
        ' DME at glideslope
        strBuffer = strBuffer & ILSGP
      Else
        ' DME at localizer
        strBuffer = strBuffer & ILSLoc
      End If
      strBuffer = strBuffer & " " & CInt(Scenery.Header.Altitude) & " )" & vbCrLf
    End If
    strBuffer = strBuffer & vbCrLf
    
    If (ILSFlags And 128) > 0 Then AFDBuffer = AFDBuffer & "      Navaid( " & ILSLoc & " " & FloatFormat(Scenery.Header.Altitude, "####0.0#") & " " & FloatFormat(ForwardRotation, "##0.0#") & " " & FloatFormat(ILSFrequency, "000.00") & " LOC TERMINAL 0 " & QuoteString(ILSID) & " )" & vbCrLf
    If (ILSFlags And 64) > 0 Then AFDBuffer = AFDBuffer & "      Navaid( " & ILSGP & " " & FloatFormat(Scenery.Header.Altitude, "####0.0#") & " " & FloatFormat(GlideSlope, "0.0") & " " & FloatFormat(ILSFrequency, "000.00") & " SLOPE TERMINAL 0 " & QuoteString(ILSID) & " )" & vbCrLf
    If (ILSFlags And 1) > 0 Then AFDBuffer = AFDBuffer & "      Navaid( " & IIf((ILSFlags And 16) > 0, ILSGP, ILSLoc) & " " & FloatFormat(Scenery.Header.Altitude, "####0.0#") & " " & FloatFormat(ForwardRotation, "##0.0#") & " " & FloatFormat(ILSFrequency, "000.00") & " DME TERMINAL 0 " & QuoteString(ILSID) & " )" & vbCrLf
    Scenery.RunwayNAVBuffer = Scenery.RunwayNAVBuffer & AFDBuffer & vbCrLf
  End If
  
  If IM > 0 Then strBuffer = strBuffer & "IM( " & myPos.CalcPoint(Dist + IM, BackwardRotation).LatLon & " " & CInt(Scenery.Header.Altitude) & " )" & vbCrLf
  If MM > 0 Then strBuffer = strBuffer & "MM( " & myPos.CalcPoint(Dist + MM, BackwardRotation).LatLon & " " & CInt(Scenery.Header.Altitude) & " )" & vbCrLf
  If OM > 0 Then strBuffer = strBuffer & "OM( " & myPos.CalcPoint(Dist + OM, BackwardRotation).LatLon & " " & CInt(Scenery.Header.Altitude) & " )" & vbCrLf

  If strBuffer <> "" Then strBuffer = strBuffer & vbCrLf

  CompileBeacons = strBuffer
End Function

Public Sub CopyTo(Destination As clsRunwayOpt)
  With Destination
    .ThrLength = ThrLength
    .ExtLength = ExtLength
    .ExtSurface = ExtSurface
    .ThrLights = ThrLights
    .ApprLights = ApprLights
    .NumStrobes = NumStrobes
    .VASILeft = VASILeft
    .VASIRight = VASIRight
    .PAPI = PAPI
    .GlideSlope = GlideSlope
    .Automatic = Automatic
    .HDistance = HDistance
    .VDistance = VDistance
    .BarSpacing = BarSpacing
    .SignHSpacing = SignHSpacing
    
    .IM = IM
    .MM = MM
    .OM = OM
    .ILSEnabled = ILSEnabled
    .ILSName = ILSName
    .ILSID = ILSID
    .ILSFrequency = ILSFrequency
    .ILSRange = ILSRange
    .ILSFlags = ILSFlags
    .ILSBeamWidth = ILSBeamWidth
  End With
End Sub

Public Sub GetData(FileData As clsFileData, Prefix As String)
  With FileData
    .GetData Prefix & "Threshold Length", ThrLength, 0
    .GetData Prefix & "Overrun Length", ExtLength, 0
'    .GetData Prefix & "Overrun Surface", ExtSurface, 99
    '!!!!!!!!!!!!
    ExtSurface = 99
    .GetData Prefix & "Threshold Lights", ThrLights, 0
    .GetData Prefix & "Approach Length", ApprLights, 0
    .GetData Prefix & "Strobes", NumStrobes, 0
    .GetData Prefix & "VASI Left", VASILeft, 0
    .GetData Prefix & "VASI Right", VASIRight, 0
    .GetData Prefix & "PAPI", PAPI, 0
    .GetData Prefix & "Glide Slope", GlideSlope, 3
    .GetData Prefix & "Automatic Calculation", Automatic, True
    .GetData Prefix & "Horz Distance", HDistance, 0
    .GetData Prefix & "Vert Distance", VDistance, 0
    .GetData Prefix & "Bar Spacing", BarSpacing, 0
    .GetData Prefix & "Sign Spacing", SignHSpacing, 0
    .GetData Prefix & "Inner Marker", IM, 0
    .GetData Prefix & "Middle Marker", MM, 0
    .GetData Prefix & "Outer Marker", OM, 0
    .GetData Prefix & "ILS Enabled", ILSEnabled, False
    If ILSEnabled Then
      .GetData Prefix & "ILS Name", ILSName, ""
      .GetData Prefix & "ILS ID", ILSID, ""
      .GetData Prefix & "ILS Frequency", ILSFrequency, 0
      .GetData Prefix & "ILS Range", ILSRange, 0
      .GetData Prefix & "ILS Flags", ILSFlags, 0
      .GetData Prefix & "ILS Beam Width", ILSBeamWidth, 0
    Else
      ILSFrequency = 110.5
      ILSRange = 16
      ILSFlags = 193
      ILSBeamWidth = 5
    End If
  End With
End Sub

Public Sub LoadAPTData(ByVal strExtend As String, ByVal strSignal As String, ByVal strVASI As String, ByVal strPAPI As String, ByVal strILS As String, ByVal strILSRange As String, ByVal strIM As String, ByVal strMM As String, ByVal strOM As String)
  Dim strTemp As String, byteTemp As Byte

  If ReadNext(strExtend, " ") <> "0" Then
    ThrLength = Meter(ReadNext(strExtend, " ") & " " & ReadNext(strExtend, " "))
    ExtLength = Meter(ReadNext(strExtend, " ") & " " & ReadNext(strExtend, " "))
  End If

  strTemp = ReadNext(strSignal, " ")
  Dim strTemp2 As String
  strTemp2 = Format$("&H" & strTemp)
  If IsNumeric(strTemp2) Then
    ThrLights = ValEx(strTemp2)
  End If
  If ThrLights > 0 Then ThrLights = ThrLights Or 1
  ApprLights = MatchText(ReadNext(strSignal, " "), "Non", "Flash", "MALSR", "MALSF", "SSALF", "SSALR", "ALSF-1", "ALSF-2", "Strobes", "ICAOIII-1", "ICAOIII-2")
  NumStrobes = ValEx(strSignal)
  If ApprLights > 0 Or ThrLights > 0 Then ThrLights = ThrLights Or 1
  
  byteTemp = ValEx(ReadNext(strVASI, "b"))
  ReadNext strVASI, "d"
  Select Case ReadNext(strVASI, " ")
    Case 1
      VASIRight = byteTemp
    Case 2
      VASILeft = byteTemp
      VASIRight = byteTemp
    Case Else ' 0
      VASILeft = byteTemp
  End Select
  
  If ValEx(ReadNext(strPAPI, " ")) > 0 Then PAPI = 1

  ILSFrequency = ValEx(ReadNext(strILS, "mhz"))
  If ILSFrequency = 0 Then
    GlideSlope = 3
    ILSFrequency = 110.5
    ILSRange = 16
    ILSFlags = 193
  Else
    ILSEnabled = True
    GlideSlope = ValEx(ReadNext(strILS, "d"))
    ILSID = Replace(ReadNext(strILS, " "), "_", " ")
    ILSName = Replace(ReadNext(strILS, " "), "_", " ")

    byteTemp = 128 Or 64 Or 16 Or 1
    Do Until strILS = ""
      Select Case ReadNext(strILS, " ")
        Case "NODME"
          byteTemp = byteTemp And (Not 17)
        Case "NOLOC"
          byteTemp = byteTemp And (Not 128)
        Case "NOGLIDE"
          byteTemp = byteTemp And (Not 64)
        Case "DMELOC"
          byteTemp = (byteTemp Or 1) And (Not 16)
      End Select
    Loop
    ILSFlags = byteTemp
  End If
  ILSBeamWidth = 2.5
  Automatic = True
  ExtSurface = 99
  
  ILSRange = Nautical(strILSRange)

  IM = Nautical(strIM)
  MM = Nautical(strMM)
  OM = Nautical(strOM)
End Sub

' Standard LoadData function to load the data
Public Sub LoadBinaryData(ByVal FileNum As Integer, ByVal Version As Integer, ByVal UserData As Long)
  Dim ByteData As Byte, BoolData As Boolean, _
    DummyByte2 As Byte, DummyStr As String, _
    DummyStr1 As String, DummyStr2 As String, _
    DummyStr3 As String, DummyStr4 As String, _
    sngX As Single, sngY As Single, _
    DummyBool As Boolean

  Select Case Version
    Case Is >= 8
      GetBinaryData FileNum, ThrLights, ApprLights, VASILeft, _
        PAPI, ThrLength, ExtLength, NumStrobes, _
        GlideSlope, IM, MM, OM, ILSEnabled, Automatic, _
        HDistance, VDistance, BarSpacing, BoolData
      If BoolData Then VASIRight = VASILeft
      If Version <= 10 Then
        If ThrLights = 3 Then ThrLights = 4
      End If
    Case Else
      GetBinaryData FileNum, ThrLights, ApprLights, VASILeft, _
        PAPI, ThrLength, ExtLength, NumStrobes, _
        GlideSlope, IM, MM, OM, ILSEnabled
  End Select
  Select Case ThrLights
    Case 2
      ThrLights = 5
    Case 3
      ThrLights = 65
    Case 4
      ThrLights = 69
  End Select
  If Version <= 8 Then ApprLights = IIf(ApprLights > 3, ApprLights + 2, ApprLights)
  ExtSurface = 99
  ILSBeamWidth = 2.5
  
  Select Case Version
    Case Is >= 8
      GetBinaryData FileNum, DummyStr, sngX, sngY, DummyBool, _
        ILSFrequency, ILSRange, ILSName, ILSID, DummyStr1, _
        DummyStr2, DummyStr3, DummyStr4, ByteData, _
        DummyByte2, BoolData
    Case Else
      GetBinaryData FileNum, sngX, sngY, _
        ILSFrequency, ILSRange, ILSName, ILSID, DummyStr1, _
        DummyStr2, DummyStr3, DummyStr4, ByteData, _
        DummyByte2
  End Select
  ILSFlags = -((ByteData And 1) > 0) * 128 + _
             -((ByteData And 2) > 0) * 64 + _
             -((ByteData And 4) > 0) * (16 + 1) + _
             -((ByteData And 8) > 0) * 1
End Sub

Public Sub SaveData(FileData As clsFileData, Prefix As String)
  With FileData
    .SaveData Prefix & "Threshold Length", ThrLength, 0
    .SaveData Prefix & "Overrun Length", ExtLength, 0
    .SaveData Prefix & "Overrun Surface", ExtSurface, 99
    .SaveData Prefix & "Threshold Lights", ThrLights, 0
    If ThrLights > 0 Then
      .SaveData Prefix & "Approach Length", ApprLights, 0
      .SaveData Prefix & "Strobes", NumStrobes, 0
      .SaveData Prefix & "VASI Left", VASILeft, 0
      .SaveData Prefix & "VASI Right", VASIRight, 0
      .SaveData Prefix & "PAPI", PAPI, 0
      .SaveData Prefix & "Glide Slope", GlideSlope, 3
      .SaveData Prefix & "Automatic Calculation", Automatic, True
      If Not Automatic Then
        .SaveData Prefix & "Horz Distance", HDistance, 0
        .SaveData Prefix & "Vert Distance", VDistance, 0
        .SaveData Prefix & "Bar Spacing", BarSpacing, 0
      End If
    End If
    .SaveData Prefix & "Sign Spacing", SignHSpacing, 0
    .SaveData Prefix & "Inner Marker", IM, 0
    .SaveData Prefix & "Middle Marker", MM, 0
    .SaveData Prefix & "Outer Marker", OM, 0
    .SaveData Prefix & "ILS Enabled", ILSEnabled, False
    If ILSEnabled Then
      .SaveData Prefix & "ILS Name", ILSName, ""
      .SaveData Prefix & "ILS ID", ILSID, ""
      .SaveData Prefix & "ILS Frequency", ILSFrequency, 0
      .SaveData Prefix & "ILS Range", ILSRange, 0
      .SaveData Prefix & "ILS Flags", ILSFlags, 0
      .SaveData Prefix & "ILS Beam Width", ILSBeamWidth, 0
    End If
  End With
End Sub

Private Sub Class_Terminate()
 Set Parent = Nothing
End Sub
