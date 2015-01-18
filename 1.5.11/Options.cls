VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Options Class
' Loads, stores, and saves program options

Option Explicit

Public Enum ExtraColors
  COLOR_Background = MAX_OBJ + 1
  COLOR_Gridline = MAX_OBJ + 2
  COLOR_FocusColor = MAX_OBJ + 3
  COLOR_DragColor = MAX_OBJ + 4
  COLOR_SyntheticColor = MAX_OBJ + 5
  COLOR_CompassColor = MAX_OBJ + 6
End Enum

Private Type ToolType
  Name As String
  Exe As String
End Type

Private Type RGBColor
  R As Byte
  G As Byte
  B As Byte
  Reserved As Byte
End Type

Private Type glRGBColor
  R As Single
  G As Single
  B As Single
End Type

Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

' General
Public _
  RememberWindowState As Boolean, _
  NeatRecentFiles As Boolean, _
  ShowHeaderProperties As Boolean, _
  Remember As Boolean, _
  ShowFractionalMinutes As Boolean, _
  Metric As Boolean, _
  Magnetic As Boolean, _
  OldStyleMenus As Boolean, _
  SaveCompressed As Boolean, _
  ShowExportWizard As Boolean, _
  UseMacroDefaults As Boolean, _
  CurrentTip As Integer, _
  ShowTips As Boolean

' FS Version
Public _
    FSVersion As Integer

' Appearance
Public _
  ToolbarVisible As Boolean, _
  StatusbarVisible As Boolean, _
  ScrollbarsVisible As Boolean, _
  CrossHair As Boolean, _
  FocusCircle As Boolean, _
  PointCircle As Boolean, _
  SnapPoints As Boolean, _
  FillPolygons As Boolean, _
  FillObjects As Boolean, _
  ThickLines As Boolean, _
  ShowCompass As Boolean

' Compile
Public _
  EditConfig As Boolean, _
  AutoCompress As Boolean, _
  KeepSourceFile As Boolean, _
  SaveBeforeCompile As Boolean

' Location
Public _
  FSPath As String, _
  TexturePath As String, _
  Compiler As String, _
  CompressName As String, _
  TextEditor As String, _
  MacroPath As String, _
  MacroPicPath As String

' Textbox
Public _
  AutoSave As Integer, _
  TextureFilter As String, _
  Grid As Integer

Public _
  ColorScheme As Integer
  
Public _
  WinMainState As Integer, _
  WinMainLeft As Long, _
  WinMainTop As Long, _
  WinMainWidth As Long, _
  WinMainHeight As Long, _
  ObjectsListWidth As Long

Public AirportFolder As String

Public MagVarMissing As Boolean

Private Colors(MAX_OBJ + 6) As Long
Private glColors(MAX_OBJ + 6) As glRGBColor
Private Visibles(MAX_OBJ) As Boolean
' Index 0-3 for files, 4 used as a dummy
' variable in AddRecentFile
Private myRecentFiles(4) As String
Private myTools() As ToolType

' Private variables
Private FSConfigFile As String

Public Event Changed()
Public Event RecentFileChanged()

Public Sub AddRecentFile(ByVal File As String)
  Dim I As Integer, J As Integer
  ' Add File to the recent file list
  myRecentFiles(4) = File
  For I = 0 To 4
    If StrComp(myRecentFiles(I), File, vbTextCompare) = 0 Then
      For J = I To 1 Step -1
        myRecentFiles(J) = myRecentFiles(J - 1)
      Next J
      myRecentFiles(0) = File
      myRecentFiles(4) = ""
      Exit For
    End If
  Next I
  RaiseEvent RecentFileChanged
End Sub

' Show the Options dialog box
Public Function EditData() As Boolean
  If FSConfigFile <> "" Then
    ShowFractionalMinutes = -GetPrivateProfileInt("MAIN", "Display_Fractional_Minutes", 1, FSConfigFile)
  Else
    ShowFractionalMinutes = -RegGetKey("Display Fractional Minutes", 1)
  End If
  
  If frmOptions.EditData(Me) Then
    If FSConfigFile <> "" Then
      WritePrivateProfileString "MAIN", "Display_Fractional_Minutes", Format$(-ShowFractionalMinutes), FSConfigFile
    Else
      RegSetKey "Display Fractional Minutes", -ShowFractionalMinutes
    End If
    RaiseEvent Changed
    Options.MagVarMissing = False
  End If
End Function

' Changes the colors to the set colors
Public Sub GetColors(DefColors() As Long)
  DefColors = Colors
End Sub

' Changes the colors to the Airport scheme
Public Sub GetDefaultAirportColors(DefColors() As Long)
  Dim I As Integer
  For I = 0 To MAX_OBJ + 6
    DefColors(I) = vbBlack
  Next I
  DefColors(OT_Runway) = vbWhite
  DefColors(OT_Polygon) = RGB(185, 185, 255)
  DefColors(OT_Taxiway) = vbBlue
  DefColors(OT_Road) = vbMagenta
  DefColors(OT_River) = vbBlue
  DefColors(OT_Line) = vbGreen
  DefColors(OT_TaxiwayLine) = RGB(192, 192, 0)
  DefColors(OT_Building) = RGB(128, 128, 128)
  DefColors(OT_ATIS) = vbWhite
  DefColors(OT_VOR) = vbGreen
  DefColors(OT_NDB) = vbBlue
  DefColors(OT_MenuEntry) = vbYellow
  DefColors(OT_Macro) = vbBlue
  DefColors(OT_Point) = vbWhite
  DefColors(OT_SurfaceArea) = vbMagenta
  DefColors(OT_Exclusion) = vbYellow

  DefColors(MAX_OBJ + 1) = vbBlack ' Background
  DefColors(MAX_OBJ + 2) = vbWhite ' Grid line
  DefColors(MAX_OBJ + 3) = vbRed
  DefColors(MAX_OBJ + 4) = vbRed
  DefColors(MAX_OBJ + 5) = vbYellow
  DefColors(MAX_OBJ + 6) = vbWhite
End Sub

' Changes the colors to the FS Scenery Creator scheme
Public Sub GetDefaultColors(DefColors() As Long)
  Dim I As Integer
  For I = 0 To MAX_OBJ + 6
    DefColors(I) = vbBlack
  Next I
  DefColors(OT_Road) = vbRed
  DefColors(OT_River) = RGB(0, 192, 192)
  DefColors(OT_Line) = RGB(255, 128, 0)
  DefColors(OT_TaxiwayLine) = RGB(192, 192, 0)
  DefColors(OT_ATIS) = vbRed
  DefColors(OT_VOR) = RGB(0, 0, 128)
  DefColors(OT_NDB) = RGB(0, 192, 0)
  DefColors(OT_TowerLocation) = RGB(255, 128, 0)
  DefColors(OT_MenuEntry) = RGB(0, 192, 0)
  DefColors(OT_Macro) = vbRed
  DefColors(OT_SurfaceArea) = vbMagenta
  DefColors(OT_Exclusion) = vbRed

  DefColors(MAX_OBJ + 1) = vbWhite   ' Background
  DefColors(MAX_OBJ + 2) = &HC0C0C0  ' Grid line
  DefColors(MAX_OBJ + 3) = vbBlue
  DefColors(MAX_OBJ + 4) = &H808080 ' Gray
  DefColors(MAX_OBJ + 5) = &H404040 ' Gray
End Sub

Public Function GetFSPath(ByVal Version As Long)
  Select Case Version
    Case Version_FS95
      GetFSPath = RegGetKey("Path", "", "SOFTWARE\Microsoft\Microsoft Games\Flight Simulator\6.0", HKEY_LOCAL_MACHINE)
    Case Version_FS98
      GetFSPath = RegGetKey("Path", "", "SOFTWARE\Microsoft\Microsoft Games\Flight Simulator\6.1", HKEY_LOCAL_MACHINE)
    Case Version_CFS1
      GetFSPath = RegGetKey("Path", "", "SOFTWARE\Microsoft\Microsoft Games\Combat Flight Simulator\1.0", HKEY_LOCAL_MACHINE)
    Case Version_FS2K
      GetFSPath = RegGetKey("Path", "", "SOFTWARE\Microsoft\Microsoft Games\Flight Simulator\7.0", HKEY_LOCAL_MACHINE)
    Case Version_CFS2
      GetFSPath = RegGetKey("EXE Path", "", "SOFTWARE\Microsoft\Microsoft Games\Combat Flight Simulator\2.0", HKEY_LOCAL_MACHINE)
    Case Version_FS2K2
      GetFSPath = RegGetKey("EXE Path", "", "SOFTWARE\Microsoft\Microsoft Games\Flight Simulator\8.0", HKEY_LOCAL_MACHINE)
  End Select
End Function

' Convert RGB long values to GL color values
Public Function GetGLColors()
  Dim Temp(MAX_OBJ + 6) As RGBColor, I As Integer
  ' A little memory hack :-)
  CopyMemory Temp(0), Colors(0), 4 * (MAX_OBJ + 6 + 1)

  For I = 1 To UBound(Temp)
    With glColors(I)
      .R = Temp(I).R / 255
      .G = Temp(I).G / 255
      .B = Temp(I).B / 255
    End With
  Next I
End Function

' Load the data from the registry
Public Function LoadData() As Boolean
  Dim Path As String, Temp As String, TempStr As String, TempStr2 As String, _
    VersionNum As Long, I As Integer, TempInt As Integer

  VersionNum = RegGetKey("VersionNum", 0)

  TempStr = RegGetKey("Language", "None")
  If TempStr = "None" Or VersionNum < 104000 Then TempStr = Lang.GetClosestLanguage()
  Lang.Name = TempStr
  If Lang.Name = "" Then
    Unload frmSplash
    LoadData = False
    Exit Function
  End If
  
  Path = GetRealName(App.Path)
  
  FSPath = RegGetKey("FS Path", "")
  FSVersion = RegGetKey("FS Version", 3)
  
  For I = Version_FS2K2 To Version_FS95 Step -1
    If FSPath <> "" Then
      Exit For
    Else
      FSPath = GetFSPath(I)
      FSVersion = I
    End If
  Next I
  
  If FSPath = "" Then
    FSPath = "C:\"
  End If
  
  FSConfigFile = AddDir(FSPath, "FS2002.CFG")
  If Not FileExists(FSConfigFile) Then
    FSConfigFile = AddDir(FSPath, "CFS2.CFG")
    If Not FileExists(FSConfigFile) Then
      FSConfigFile = AddDir(FSPath, "FS2000.CFG")
      If Not FileExists(FSConfigFile) Then
        FSConfigFile = AddDir(FSPath, "CFS.CFG")
        If Not FileExists(FSConfigFile) Then
          FSConfigFile = AddDir(FSPath, "FLTSIM98.CFG")
          If Not FileExists(FSConfigFile) Then
            FSConfigFile = AddDir(FSPath, "FLTSIM95.CFG")
            If Not FileExists(FSConfigFile) Then
              FSConfigFile = ""
            End If
          End If
        End If
      End If
    End If
  End If
  If FSConfigFile <> "" Then
    ShowFractionalMinutes = -GetPrivateProfileInt("MAIN", "Display_Fractional_Minutes", 1, FSConfigFile)
  Else
    ShowFractionalMinutes = -RegGetKey("Display Fractional Minutes", 1)
  End If
  
  RememberWindowState = -RegGetKey("Remember Window State", 0)
  NeatRecentFiles = -RegGetKey("Neat Recent Files", 1, GeneralKey)
  ShowHeaderProperties = -RegGetKey("Show Header Properties", 0)
  Remember = -RegGetKey("Remember", 1)
  
  Metric = RegGetKey("Metric", 1)
  Magnetic = RegGetKey("Magnetic", 0)
  OldStyleMenus = -RegGetKey("Old Style Menus", 0)
  SaveCompressed = -RegGetKey("Save Compressed", 1)
  ShowExportWizard = -RegGetKey("Show Export Wizard", 0)
  UseMacroDefaults = -RegGetKey("Use Macro Defaults", 1)

  CurrentTip = RegGetKey("Current Tip", 0)
  ShowTips = -RegGetKey("Show Tips", 1)

  ToolbarVisible = -RegGetKey("Toolbar", 1)
  StatusbarVisible = -RegGetKey("Statusbar", 1)
  ScrollbarsVisible = -RegGetKey("Scrollbars", 1)
  CrossHair = -RegGetKey("MousePointer", 0)
  FocusCircle = -RegGetKey("Focus Circle", 1)
  PointCircle = -RegGetKey("Point Circle", 1)
  SnapPoints = -RegGetKey("Snap Points", 1)
  FillPolygons = -RegGetKey("Fill Polygons", 1)
  ThickLines = -RegGetKey("Draw Thick Lines", 1)
  FillObjects = -RegGetKey("Fill Objects", 1)
  ShowCompass = -RegGetKey("Show Compass", 0)
  
  EditConfig = -RegGetKey("Edit Scenery Configuration", 1)
  KeepSourceFile = -RegGetKey("Keep Source File", 0)
  SaveBeforeCompile = RegGetKey("Save Before Compile", 0)

  TexturePath = RegGetKey("Texture Path", AddDir(FSPath, "Texture"))
  Compiler = GetRealName(RegGetKey("Compiler", AddDir(Path, "SCASM\SCASM.EXE")))
  CompressName = RegGetKey("Compress Path", AddDir(Path, "Bglzip.exe"))
  AutoCompress = RegGetKey("AutoCompress", FileExists(CompressName))
  TextEditor = RegGetKey("Notepad", GetRealName(AddDir(GetWindowsDir(), "NOTEPAD.EXE")))
  Temp = RegGetKey("PGMDirectory", "None", "Software\Airport 2.xx")
  If Temp <> "None" Then
    TempStr = GetRealName(AddDir(Temp, "API"))
    TempStr2 = GetRealName(RegGetKey("UserMACRODirectory", "", "Software\Airport 2.xx"))
    MacroPath = RegGetKey("Macros", AddDir(Path, "Macros") & ";" & TempStr & ";" & TempStr2)
          
    TempStr = GetRealName(AddDir(Temp, "Resource"))
    MacroPicPath = RegGetKey("MacroPics", AddDir(Path, "Macros\Bitmaps") & ";" & TempStr)
  Else
    MacroPath = RegGetKey("Macros", AddDir(Path, "Macros"))
    MacroPicPath = RegGetKey("MacroPics", AddDir(Path, "Macros\Bitmaps"))
  End If

  AutoSave = RegGetKey("AutoSave", 15)
  TextureFilter = RegGetKey("TextureFilter", "*.bmp;*.r8;*.txr;*.oav;*.?af;*.pat")
  Grid = RegGetKey("Grid", 1000)
  
  ColorScheme = RegGetKey("Color Scheme", 0, "\Colors")
  If ColorScheme = 2 Then ColorScheme = -1
  GetDefaultColors Colors

  For I = 1 To MAX_OBJ + 6
    Colors(I) = RegGetKey(ObjectNames(I), Colors(I), "\Colors")
  Next I
  
  For I = 1 To MAX_OBJ
    Visibles(I) = -RegGetKey(ObjectNames(I), 1, "\Visibles")
  Next I
  GetGLColors
  
  TempInt = RegGetKey("Count", -1, "\Tools")
  If TempInt = -1 Then
    ReDim myTools(3)
    I = 0
    TempStr = RegGetKey("Program", "", "Software\MW\DXTBmp")
    If TempStr <> "" Then
      I = I + 1
      myTools(I).Name = "DXTBmp"
      myTools(I).Exe = TempStr
    End If
    
    TempStr = RegGetKey("AppPath", "", "Software\VB and VBA Program Settings\Brueckner\EOD")
    If TempStr <> "" Then
      I = I + 1
      myTools(I).Name = "Easy Object Designer"
      myTools(I).Exe = AddDir(TempStr, "eod.exe")
    End If
    ReDim Preserve myTools(I)
  Else
    ReDim myTools(TempInt)
    For I = 1 To TempInt
      With myTools(I)
        .Name = RegGetKey("ToolName" & CStr(I), "Error", "\Tools")
        .Exe = RegGetKey("ToolExe" & CStr(I), "", "\Tools")
      End With
    Next I
  End If
  WinMainState = RegGetKey("WindowState", vbMaximized)
  WinMainLeft = RegGetKey("WindowLeft", 0)
  WinMainTop = RegGetKey("WindowTop", 0)
  WinMainWidth = RegGetKey("WindowWidth", 0)
  WinMainHeight = RegGetKey("WindowHeight", 0)
  ObjectsListWidth = RegGetKey("ListWidth", 0)
  
  For I = 1 To 4
    myRecentFiles(I - 1) = RegGetKey("RecentFile" & CStr(I), "")
  Next I
  
  LoadData = True
  If VersionNum < CurExpandedVersion Then
    frmSplash.Hide
    TempInt = Screen.MousePointer
    Screen.MousePointer = vbDefault
    If MsgBox(LoadResString(RES_License), vbInformation Or vbYesNo Or vbMsgBoxHelpButton, , Lang.HelpFile, 300) = vbNo Then
      Unload frmSplash
      LoadData = False
      Exit Function
    End If
    If VersionNum < 105070 Then frmLanguage.EditData True
    frmSplash.Show
    Screen.MousePointer = TempInt
  End If
  
  RaiseEvent Changed
  RaiseEvent RecentFileChanged
End Function

Public Property Get ObjectColor(ByVal Index As Integer) As Long
  ObjectColor = Colors(Index)
End Property

Public Property Let ObjectColor(ByVal Index As Integer, ByVal vData As Long)
  Colors(Index) = vData
End Property

Public Property Get ObjectVisible(ByVal Index As Integer) As Boolean
  ObjectVisible = Visibles(Index)
End Property

Public Property Let ObjectVisible(ByVal Index As Integer, ByVal vData As Boolean)
  Visibles(Index) = vData
End Property

Public Property Get RecentFiles(ByVal Index As Integer)
  RecentFiles = myRecentFiles(Index)
End Property

' Save the data to the registry
Public Sub SaveData()
  Dim I As Integer
  
  If RegGetKey("VersionNum", 0) < CurExpandedVersion Then
    RegSetKey "VersionNum", CurExpandedVersion
  End If

  RegSetKey "Program Folder", GetRealName(App.Path)
  RegSetKey "Language", Lang.Name
  
  RegSetKey "Remember Window State", -RememberWindowState
  RegSetKey "Neat Recent Files", -NeatRecentFiles, GeneralKey
  RegSetKey "Show Header Properties", -ShowHeaderProperties
  RegSetKey "Remember", -Remember
  RegSetKey "Metric", -Metric
  RegSetKey "Magnetic", -Magnetic
  RegSetKey "Old Style Menus", -OldStyleMenus
  RegSetKey "Save Compressed", -SaveCompressed
  RegSetKey "Show Export Wizard", -ShowExportWizard
  RegSetKey "Use Macro Defaults", -UseMacroDefaults
  
  RegSetKey "Current Tip", CurrentTip
  RegSetKey "Show Tips", -ShowTips
  
  RegSetKey "Toolbar", -ToolbarVisible
  RegSetKey "Statusbar", -StatusbarVisible
  RegSetKey "Scrollbars", -ScrollbarsVisible
  RegSetKey "MousePointer", -CrossHair
  RegSetKey "Focus Circle", -FocusCircle
  RegSetKey "Point Circle", -PointCircle
  RegSetKey "Snap Points", -SnapPoints
  RegSetKey "Fill Polygons", -FillPolygons
  RegSetKey "Fill Objects", -FillObjects
  RegSetKey "Draw Thick Lines", -ThickLines
  RegSetKey "Show Compass", -ShowCompass
  
  RegSetKey "Edit Scenery Configuration", -EditConfig
  RegSetKey "AutoCompress", -AutoCompress
  RegSetKey "Keep Source File", -KeepSourceFile
  RegSetKey "Save Before Compile", -SaveBeforeCompile
    
  RegSetKey "FS Version", FSVersion
  RegSetKey "FS Path", FSPath
  RegSetKey "Texture Path", TexturePath
  RegSetKey "Compiler", Compiler
  RegSetKey "Compress Path", CompressName
  RegSetKey "Notepad", TextEditor
  RegSetKey "Macros", MacroPath
  RegSetKey "MacroPics", MacroPicPath

  RegSetKey "AutoSave", AutoSave
  RegSetKey "TextureFilter", TextureFilter
  RegSetKey "Grid", Grid

  RegSetKey "Color Scheme", ColorScheme, "\Colors"

  For I = 1 To MAX_OBJ + 6
    RegSetKey ObjectNames(I), Colors(I), "\Colors"
  Next I

  For I = 1 To MAX_OBJ
    RegSetKey ObjectNames(I), -Visibles(I), "\Visibles"
  Next I

  RegSetKey "Count", UBound(myTools), "\Tools"
  For I = 1 To UBound(myTools)
    With myTools(I)
      RegSetKey "ToolName" & CStr(I), .Name, "\Tools"
      RegSetKey "ToolExe" & CStr(I), .Exe, "\Tools"
    End With
  Next I

  RegSetKey "WindowState", WinMainState
  RegSetKey "WindowLeft", WinMainLeft
  RegSetKey "WindowTop", WinMainTop
  RegSetKey "WindowWidth", WinMainWidth
  RegSetKey "WindowHeight", WinMainHeight
  RegSetKey "ListWidth", ObjectsListWidth
  
  For I = 1 To 4
    RegSetKey "RecentFile" & CStr(I), myRecentFiles(I - 1)
  Next I
End Sub

' Changes the colors to the given colors
Public Sub SetColors(DefColors() As Long)
  CopyMemory Colors(0), DefColors(0), 4 * (UBound(DefColors) - LBound(DefColors) + 1)
End Sub

Public Sub SetGLBackColor()
  With glColors(COLOR_Background)
    glClearColor .R, .G, .B, 0
  End With
End Sub

Public Sub SetGLColor(ByVal Index As Long)
  With glColors(Index)
    glColor3f .R, .G, .B
  End With
End Sub

' Get the number of tools
Public Property Get ToolCount() As Integer
  ToolCount = UBound(myTools)
End Property

Public Property Let ToolCount(ByVal Max As Integer)
  ReDim myTools(Max)
End Property

' Get the exe text of a tool
Public Property Get ToolExe(ByVal Index As Integer) As String
  ToolExe = myTools(Index).Exe
End Property

' Set the exe text of a tool
Public Property Let ToolExe(ByVal Index As Integer, ByVal RHS As String)
  myTools(Index).Exe = RHS
End Property

' Get the show name for a tool
Public Property Get ToolName(ByVal Index As Integer) As String
  ToolName = myTools(Index).Name
End Property

' Set the show name for a tool
Public Property Let ToolName(ByVal Index As Integer, ByVal RHS As String)
  myTools(Index).Name = RHS
End Property

