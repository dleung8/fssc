Attribute VB_Name = "basFssc"
' FS Scenery Creator
' Creates scenery in SCASM format for Flight Simulator
Option Explicit

' Structures for API calls
Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Type POINTAPI
  X As Long
  Y As Long
End Type

Private Type PROCESS_INFORMATION
  hProcess As Long
  hThread As Long
  dwProcessId As Long
  dwThreadId As Long
End Type

Private Type STARTUPINFO
  cb As Long
  lpReserved As String
  lpDesktop As String
  lpTitle As String
  dwX As Long
  dwY As Long
  dwXSize As Long
  dwYSize As Long
  dwXCountChars As Long
  dwYCountChars As Long
  dwFillAttribute As Long
  dwFlags As Long
  wShowWindow As Integer
  cbReserved2 As Integer
  lpReserved2 As Byte
  hStdInput As Long
  hStdOutput As Long
  hStdError As Long
End Type
 
Private Enum enPriority_Class
  NORMAL_PRIORITY_CLASS = &H20
  IDLE_PRIORITY_CLASS = &H40
  HIGH_PRIORITY_CLASS = &H80
End Enum

'Private Type SHFILEOPSTRUCT
'  hwnd As Long
'  wFunc As Long
'  pFrom As String
'  pTo As String
'  fFlags As Long
'  fAnyOperationsAborted As Long
'  hNameMappings As Long
'  lpszProgressTitle As String
'End Type

' FSSC Structures

Public Type SyntheticType
  ID As String
  File As String
End Type

Public Type MacroType
  File As String
  Name As String
  Bitmap As String
  V1 As Integer
  V2 As Integer
  MScale As Single
End Type

Public Type MacroArrayType
  DirName As String
  Data() As MacroType
End Type

Public Type MacroMetaType
  ' VOD
  VODData As String
  MacroDesc As String
  
  ' FSSC
  Defaults As String
  MScale As String
  Points As String
  QuickPoints As String
  RGBEnabled As String
  
  ' Airport
  DefaultScale As String
  DefaultParams As String
  ParamDescr As String
  DefaultRange As String
  DefaultDensity As String
  DesignShape As String
  Textures As String
  
  ' Sort of meta data
  AirportScale As String
  APTMacroScale As String
End Type

Public Type RegionsType
  ID As String
  Regions(8) As String
End Type

Public Type BldgType
  File As String
  Color As Long
  Windows As Byte
  Roof As Byte
End Type

'Public Type TexturesNotExistType
'  OldFile As String
'  NewFile As String
'End Type

Public Type PicType
  Width As Long
  Height As Long
  Depth As Long
  NumCols As Long
  fin As Long
  Comment As String * 80
  CMap As String * 768
  jlib As Long
  tlib As Long
  plib As Long
  buff As Long
  fout As Long
  ptr As Long
  progress As Long
  ilib As Long
  stype As Long
  spare As String * 120
  spare2 As String * 40
End Type

Public Enum HelpIDEnum
  IDH_ERR_OPENFILE = 2090
  IDH_ERR_Program = 2091
End Enum

Private Type MSGBOXPARAMS
  cbSize As Long
  hwndOwner As Long
  hInstance As Long
  lpszText As String
  lpszCaption As String
  dwStyle As Long
  lpszIcon As String
  dwContextHelpId As Long
  lpfnMsgBoxCallback As Long
  dwLanguageId As Long
End Type

Private Type HELPINFO
  cbSize As Long
  iContextType As Long
  iCtrlId As Long
  hItemHandle As Long
  dwContextId As Long
  MousePos As POINTAPI
End Type

' Region Management Functions
Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function RectInRegion Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long
Public Declare Function DeleteRegion Lib "gdi32" Alias "DeleteObject" (ByVal hObject As Long) As Long
Public Const ALTERNATE = 1
Public Const WINDING = 2

' Drawing functions
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Const SRCCOPY = &HCC0020

' DrawEdge functions
Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENOUTER = &H2
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_OUTER = &H3
Public Const BDR_INNER = &HC
Public Const BDR_RAISED = &H5
Public Const BDR_SUNKEN = &HA
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8
Public Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Public Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Public Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Public Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

' TreeView bug
' Clicking Tooltips does not select the item below.
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, bEnable As Long) As Long
Public Const TV_FIRST = &H1100
Public Const TVM_GETTOOLTIPS = (TV_FIRST + 25)

' SendMessage variants
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

' MessageBox: better than Msgbox
Private Declare Function MessageBoxIndirect Lib "user32" Alias "MessageBoxIndirectA" (lpMsgBoxParams As MSGBOXPARAMS) As Long
Private Const LOCALE_USER_DEFAULT = &H400

' Run Process Functions
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Const INFINITE = &HFFFF
Private Const STARTF_USESHOWWINDOW = &H1

' Window management function
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

' Icon drawing functions
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Private Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As Long) As Long
Private Const IDI_QUESTION = 32514&
Private Const IDI_EXCLAMATION = 32515&

' Compacted Path function
Private Declare Function PathCompactPath Lib "shlwapi" Alias "PathCompactPathA" (ByVal hdc As Long, ByVal lpszPath As String, ByVal DX As Long) As Long

' Windows Directory
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

' HTML Help
Public Declare Function HtmlHelpString Lib "HHCtrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As String) As Long
Public Declare Function HtmlHelpLong Lib "HHCtrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Long) As Long
Public Const HH_DISPLAY_TOPIC = &H0
Public Const HH_HELP_CONTEXT = &HF

' MW library bitmap functions
Public Declare Function AnytoBmps Lib "mwvb.dll" Alias "vb_anytobmps" (ByVal Src As String, ByVal Dest As String, Pic As PicType, ByVal pdcsize As Long, ByVal jpegscale As Long) As Long
Public Declare Function BmptoAnys Lib "mwvb.dll" Alias "vb_bmptoanys" (ByVal Src As String, ByVal Dest As String, Pic As PicType, ByVal Format As Long, ByVal jpegquality As Long) As Long
Public Declare Sub BmpResize Lib "mwvb.dll" Alias "vb_bmpresize" (ByVal Src As String, ByVal Dest As String, Pic As PicType, ByVal Width As Long, ByVal Length As Long)

' Constant Variables
Public Const AbsFocusWidth = 10
Public Const MyFileVer = 20
Public Const MinorRevision = 0
Public Const MAX_OBJ = OT_Point
Public Const Email = "fsschelp@yahoo.com"
Public Const Webpage = "http://www.fssc.avsim.net/"
Public Const Digits = "0123456789"
Public Const Version_FS95 = 0
Public Const Version_FS98 = 1
Public Const Version_CFS1 = 2
Public Const Version_FS2K = 3
Public Const Version_CFS2 = 4
Public Const Version_FS2K2 = 5
Public Const Point_Threshold = 5

Public Const MUNKNOWN = 0
Public Const MFILEEXIST = 1
Public Const MFILENOTEXIST = 2

' Color masks
Public Const PalColormask = &HFF
Public Const RGBColorMask = &HFFFFFF
Public Const TransparentMask = &HF000000
Public Const PalMask = &H10000000
Public Const RGBMask = &H20000000
Public Const NightMask = &H40000000

' Fairly Constant Variables
Public DigitsDecimal As String
Public DigitsSigns As String
Public DigitsDecimalSigns As String
Public CurExpandedVersion As Long
Public FocusWidth As Single
Public TempPathName As String

' Objects
Public Scenery As clsScenery
Public Options As clsOptions
Public Lang As clsLanguage
Public picEditor As clsOpenGL
Public cDialog As clsCDialog

' Cached data
Public Regions(6) As RegionsType
Public SynNames(26) As SyntheticType
Public FSColors(49) As Long
Public Building(7) As BldgType
Public Building1_3(8 To 85) As String
Public Building2(4 To 84) As String
Public BuildingR(4 To 33) As String
Public MacroLst() As MacroArrayType
Public ZoomLevels(12) As Single
Public TextureHeader() As Byte
Public ObjectNames(MAX_OBJ + 6) As String
'Public AirportFolder As String
'Public MacrosNotExist() As String
'Public TexturesNotExist() As TexturesNotExistType

' Miscellaneous
Public MainMacroDir As String
Public UntitledName As String
'Public MsgBoxHelpID As Integer
Public MultiSelection As Boolean
Public TabValue As Byte

' Autosave variables
Public NextAutoSave As Long
Public AutoSaveFiles(3) As String
Public AutoSavePointer As Integer

' Default values
Public Defaults(MAX_OBJ) As clsObject

Private ScreenMouseCount As Long

' For debugging
Public Closing As Boolean

' Initializes variables and shows frmMain
Public Sub Main()
  Dim I As Integer
  SetScreenMousePointer vbHourglass
  AppKey = "Software\Leung\FS Scenery Creator\1.5"
  CopyrightYear = "2000-2003"
  CurExpandedVersion = App.Major * 100000 + App.Minor * 1000 + App.Revision * 10 + MinorRevision
  ChangeDir GetRealName(App.Path)
  
  With frmSplash
    .Show
    .LoadPercent 5
    Set Lang = New clsLanguage
    .LoadPercent 20
    Set Options = New clsOptions
    .LoadPercent 30
  End With
  
  ' Cached data
  For I = 0 To UBound(ObjectNames)
    ObjectNames(I) = LoadResString(RES_UnlocalizedObjectNames + I)
  Next I
  
  If Options.LoadData Then
    Load frmMain
    frmMain.Show
  End If
  SetScreenMousePointer vbDefault
End Sub

' Add the texture file NewFile to the list of files if
' it doesn't already exist in the list
Public Sub AddTexFile(ByRef Files() As String, ByRef FileCount As Integer, ByVal NewFile As String, Optional ByVal AdditionalDir As String)
  Dim I As Integer, myFile As String, NewFileTitle As String, Temp As String
  Dim PathIndex As Integer
  
  If NewFile = "" Then Exit Sub
  NewFileTitle = GetFileTitle(NewFile)
  
  ' Find the file
  If FileExists(AddDir(Scenery.ExportPath, "Texture\" & NewFileTitle)) Then
    myFile = AddDir(Scenery.ExportPath, "Texture\" & NewFileTitle)
  ElseIf FileExists(AddDir(Options.FSPath, "Texture\" & NewFileTitle)) Then
    myFile = AddDir(Options.FSPath, "Texture\" & NewFileTitle)
  ElseIf FileExists(AddDir(Options.TexturePath, NewFileTitle)) Then
    myFile = AddDir(Options.TexturePath, NewFileTitle)
  ElseIf FileExists(NewFileTitle) And InStr(NewFileTitle, "\") = 0 Then
    myFile = AddDir(CurDir$, NewFileTitle)
  ElseIf FileExists(AddDir(App.Path, "Texture\" & NewFileTitle)) Then
    myFile = AddDir(App.Path, "Textures\" & NewFileTitle)
  ElseIf FileExists(NewFileTitle) Then
    myFile = NewFileTitle
  ElseIf FileExists(NewFile) Then
    myFile = NewFile
  Else
    If AdditionalDir <> "" Then
      If FileExists(AddDir(AdditionalDir, NewFileTitle)) Then
        myFile = AddDir(AdditionalDir, NewFileTitle)
      Else
        Temp = GetDir(AdditionalDir)
        If Temp <> "" Then
          If FileExists(AddDir(Temp, NewFileTitle)) Then
            myFile = AddDir(Temp, NewFileTitle)
          ElseIf FileExists(AddDir(Temp, "Texture\" & NewFileTitle)) Then
            myFile = AddDir(Temp, "Texture\" & NewFileTitle)
          ElseIf FileExists(AddDir(Temp, "Textures\" & NewFileTitle)) Then
            myFile = AddDir(Temp, "Textures\" & NewFileTitle)
          End If
        End If
      End If
    End If
  
    If myFile = "" Then
      myFile = Lang.ResolveString(RES_ERR_LocateTexture, NewFile)
    End If
  End If
  myFile = GetRealName(myFile)
  
  For I = 0 To FileCount - 1
    If Mid$(Files(I), 2) = myFile Then Exit Sub
  Next I
  If FileCount > UBound(Files) Then ReDim Preserve Files(FileCount * 2)
    
  If FileExists(myFile) And StrComp(GetDir(myFile), AddDir(Options.FSPath, "Texture"), vbTextCompare) = 0 Then
    PathIndex = 4
  Else '''If FileExists(AddDir(App.Path, "Texture\") & GetFileTitle(myFile)) Then
    PathIndex = 2
  End If
  
  Files(FileCount) = PathIndex & myFile
  FileCount = FileCount + 1
End Sub

' Append a unit to a numeric value
Public Function Append(ByVal Value As Single, ByVal Unit As Integer, Optional ByVal FormStr As String) As String
  Append = Format$(Value, FormStr) & Lang.GetString(Unit)
End Function

' Function ChangeDir
' Changes the Directory
' Dir = Dir to start with
Public Sub ChangeDir(ByVal DirName As String)
  On Error Resume Next
  ChDrive DirName
  On Error GoTo DirErr:
  ChDir DirName
  Exit Sub
DirErr:
  If MsgBox(Lang.ResolveString(RES_ERR_DirExists, DirName), vbRetryCancel Or vbExclamation) = vbRetry Then
    Resume
  Else
    Resume Next
  End If
End Sub

' Get the FSSC equivalent of a SCASM color value
Public Function ColorFromSCASM(ByVal ColorStr As String, ByVal Transparency As String, ByVal Palette As String) As Long
  Dim Res As Long
  On Error Resume Next
  If Palette = "F0" Then
    If Len(ColorStr) < 2 Then
      Res = 0
    Else
      Res = PalMask + CLng("&H" & Left$(ColorStr, 2))
    End If
  ElseIf Len(ColorStr) = 8 Then
    Res = RGB(CLng("&H" & Left$(ColorStr, 2)), CLng("&H" & Mid$(ColorStr, 4, 2)), CLng("&H" & Mid$(ColorStr, 7, 2)))
    If Transparency <> "" Then
      Res = Res + CLng("&H" & Right$(Transparency, 1)) * &H1000000
    End If
  Else
    Res = 0
  End If
  ColorFromSCASM = Res
End Function

' Get the SCASM equivalent of a FSSC color value
Public Sub ColorToSCASM(ByVal ColorData As Long, ByRef ColorStr As String, ByRef Palette As String)
  Dim Temp(3) As Byte
  If ColorData And PalMask Then
    ' Palette
    ColorStr = ColorData And PalColormask
    Palette = "F0"
  ElseIf ColorData = 0 Then
    ' None
    ColorStr = "0"
    Palette = "F0"
  Else
    ' RGB
    CopyMemory Temp(0), ColorData, 4
    ColorStr = Temp(0) & " " & Temp(1) & " " & Temp(2)
    Palette = "E" & Hex$((ColorData And TransparentMask) / &H1000000)
  End If
End Sub

Private Function Combine(Arr() As String, ByVal Num As Long, ByVal Delimiter As String) As String
  Dim I As Long, Result As String
  For I = 0 To Num - 1
    Result = Result & Arr(I) & Delimiter
  Next I
  If Len(Result) > 0 Then
    Combine = Left$(Result, Len(Result) - 1)
  End If
End Function

Public Function CompactedPath(ByVal sPath As String, ByVal lMaxPixels As Long, ByVal hdc As Long) As String
  PathCompactPath hdc, sPath, lMaxPixels
  CompactedPath = StripTerminator(sPath)
End Function

' Use the distance formula to find the distance
Public Function Distance(ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double) As Double
  Dim X As Double, Y As Double
  X = X2 - X1
  Y = Y2 - Y1
  Distance = Sqr(CDbl(X) * X + Y * Y)
End Function

' Use the distance formula to find the square of
' the distance, avoids using square root function
Public Function Distance2(ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double) As Double
  Dim X As Double, Y As Double
  X = X2 - X1
  Y = Y2 - Y1
  Distance2 = X * X + Y * Y
End Function

' Draw the exclamation icon
Public Sub DrawExclaimIcon(hdc As Long, ByVal X As Long, ByVal Y As Long)
  DrawIcon hdc, X, Y, LoadIcon(0&, IDI_EXCLAMATION)
End Sub

' Draw the exclamation icon
Public Sub DrawQuestionIcon(hdc As Long, ByVal X As Long, ByVal Y As Long)
  DrawIcon hdc, X, Y, LoadIcon(0&, IDI_QUESTION)
End Sub

' Get the RGB color of a FSSC color value
Public Function ExtractColor(ByVal ColorData As Long) As Long
  If ColorData And PalMask Then
    ' Palette
    ExtractColor = FSColors(ColorData And PalColormask)
  ElseIf ColorData = 0 Then
    ' None
    ExtractColor = vbButtonFace
  Else
    ' RGB
    ExtractColor = ColorData And RGBColorMask
  End If
End Function

' Forces a format using the US locale setting
Public Function FloatFormat(ByVal Num As Double, ByVal FormStr As String) As String
  Dim Res As String, Sep As String, Y As Long
  Res = Format$(Num, FormStr)
  Sep = LocaleInfo(LOCALE_SDECIMAL)
  If Sep <> "." Then
    Y = InStrRev(Res, Sep)
    If Y > 0 Then Mid$(Res, Y, 1) = "."
  End If
  FloatFormat = Res
End Function

' Fills the ParamArray with data from Filenum
Public Sub GetBinaryData(ByVal FileNum As Integer, ParamArray Outputs() As Variant)
  Dim I As Integer, intX As Integer, longX As Long, _
      sngX As Single, byteX As Byte, strX As String, _
      strX2() As Byte
  
  For I = 0 To UBound(Outputs)
    Select Case VarType(Outputs(I))
      Case vbInteger
        Get #FileNum, , intX
        Outputs(I) = intX
      Case vbLong
        Get #FileNum, , longX
        Outputs(I) = longX
      Case vbBoolean
        strX = " "
        Get #FileNum, , strX
        Outputs(I) = (strX = Chr$(1))
      Case vbSingle
        Get #FileNum, , sngX
        Outputs(I) = sngX
      Case vbByte
        Get #FileNum, , byteX
        Outputs(I) = byteX
      Case vbObject
        Outputs(I).LoadBinaryData FileNum
      Case Else ' vbString
        Get #FileNum, , intX
        If intX > 0 Then
          ReDim strX2(intX - 1)
          Get #FileNum, , strX2
          Outputs(I) = StrConv(strX2, vbUnicode)
        Else
          Outputs(I) = ""
        End If
    End Select
  Next I
End Sub

' Get the unit of measurement from a data string
Public Function GetUnit(ByVal Data As String) As String
  Dim I As Integer
  For I = Len(Data) To 1 Step -1
    If InStrRev(DigitsDecimalSigns, Mid$(Data, I, 1)) > 0 Then
      GetUnit = Trim$(Mid$(Data, I + 1))
      Exit For
    End If
  Next I
End Function

' Get the windows directory
Public Function GetWindowsDir() As String
  Dim Buffer As String

  Buffer = String$(260 + 1, 0)
  GetWindowsDirectory Buffer, 260
  GetWindowsDir = StripTerminator(Buffer)
End Function

' Given a number, format it to a hex color value
Public Function HexFormat(ByVal Number As Byte) As String
  Dim X As String
  X = Hex$(Number)
  If Len(X) = 0 Then
    X = "00"
  ElseIf Len(X) = 1 Then
    X = "0" + X
  End If
  HexFormat = X
End Function

' Line Input - can read text only delimited by chr$(10)
Public Sub LineInputEx(ByVal FileNum As Integer, ByRef Res As String)
  Dim Y As Integer, Pos As Long
  Pos = Seek(FileNum)
  Line Input #FileNum, Res
  Y = InStr(Res, Chr$(10))
  If Y > 0 Then
    Res = Left$(Res, Y - 1)
    Seek #FileNum, Pos + Y
  End If
End Sub


' Load the list of macros from macro hierarchy files
' and cache the macro directories
Public Sub LoadMacros()
  Dim FileNum As Integer, MacrosFile As String, _
    Hierarchy() As String, HierarchyCache As String, _
    TempList() As MacroType, NumHierarchies As Integer

  Dim FileVer As Long
    
  Dim myDirs As String, myStr As String, _
    myTabs As Integer, myDir As String, _
    myFile As String, myNewDir As String, _
    myDirNeatList As String
    
  Dim FileList() As String, ListCounter As Integer
  
  Dim I As Integer, J As Integer, _
    FileCounter As Integer, hCounter As Integer
  
  
  ' LoadMacros algorithm
  ' 1. load the file lists into one master list
  '    with fully qualified directory structures.
  ' 2. Sort this list
  ' 3. Reform the master list into an array of
  '    hierarchy
  '
  ' These steps allow items to be appended to an
  '    existing hierarchy from a different file list
  '
  ' 4. Load the macro files from the specified macro
  '    folders into three new hierarchies and sort.
  
  SetScreenMousePointer vbHourglass
  
  myDirs = AddDir(App.Path, "Macros") & ";" & Options.MacroPath
  ReDim FileList(10)
  MacrosFile = MultiDir("*.lst", myDirs)
  ListCounter = 1
  Do Until MacrosFile = ""
    If ListCounter > UBound(FileList) Then ReDim Preserve FileList(ListCounter * 2)
    FileList(ListCounter) = MacrosFile
    ListCounter = ListCounter + 1
    MacrosFile = MultiDir()
  Loop
  ReDim Preserve FileList(ListCounter - 1)
  
  ReDim TempList(100)

  FileNum = FreeFile
  FileCounter = 0
  
  On Error GoTo MacroListError:
  
  For I = 1 To UBound(FileList)
    Open FileList(I) For Input As #FileNum
    On Error Resume Next
    LineInputEx FileNum, myStr
    On Error GoTo MacroListError:
    If myStr = "FS Scenery Creator Macro Hierarchy" Then
      Do Until EOF(FileNum)
        LineInputEx FileNum, myStr
        If StrComp(Left$(myStr, 9), "Versions=", vbTextCompare) = 0 Then
          FileVer = Val(Mid$(myStr, 10))
          If (FileVer And 2 ^ Options.FSVersion) = 0 Then Exit Do
        ElseIf StrComp(Left$(myStr, 9), "Main=True", vbTextCompare) = 0 Then
          MainMacroDir = GetDir(FileList(I))
        ElseIf myStr <> "" Then
          ' Count tabs
          myTabs = 0
          Do While myTabs < Len(myStr)
            If Mid$(myStr, myTabs + 1, 1) <> vbTab Then Exit Do
            myTabs = myTabs + 1
          Loop
          
          myStr = Mid$(myStr, myTabs + 1)
          
          If Right$(myStr, 1) = "\" Then
            ReDim Preserve Hierarchy(myTabs)
            Hierarchy(myTabs) = myStr
            NumHierarchies = NumHierarchies + 1
          Else
            FileCounter = FileCounter + 1
            If FileCounter > UBound(TempList) Then ReDim Preserve TempList(FileCounter * 2)

            With TempList(FileCounter)
              .Name = Combine(Hierarchy, myTabs, ">") & ReadNext(myStr, ",")
              If StrComp(Left$(myStr, 8), "LibObj: ", vbTextCompare) = 0 Then
                .File = ReadNext(myStr, ",")
                .Bitmap = ReadNext(myStr, ",")
                .V2 = Val(ReadNext(myStr, ","))
                .MScale = Val(ReadNext(myStr, ","))
                If .MScale = 0 Then .MScale = 1
              Else
                myFile = ReadNext(myStr, ",")
                If myFile = "" Then
                  FileCounter = FileCounter - 1
                Else
                  If InStr(myFile, ":") > 0 Then
                    .File = GetRealName(myFile)
                  Else
                    .File = GetRealName(MultiDir(myFile, myDirs))
                  End If
                  .Bitmap = ReadNext(myStr, ",")
                End If
              End If
            End With
          End If
        End If
      Loop
    End If
    Close #FileNum
  Next I

  
  ReDim Preserve TempList(FileCounter)
  MacroQuickSort TempList, 1, FileCounter

  ' Rebuild a true hierarchy list
  ReDim MacroLst(NumHierarchies + 10) ' reserve 10 extra
  NumHierarchies = 0
  HierarchyCache = ""
  For I = 1 To FileCounter
    myDir = TempList(I).Name
    myFile = ReadLast(myDir, "\")
    If myDir <> HierarchyCache Then
      ' new hierarchy
      If hCounter > 0 Then ReDim Preserve MacroLst(hCounter).Data(J)
      hCounter = hCounter + 1
      ' Strings that begin with "?" are informational,
      ' and "?" ensures that they come to the top of the
      ' sort order
      MacroLst(hCounter).DirName = Replace(myDir, "\>?", "\>")
      HierarchyCache = myDir
      ReDim MacroLst(hCounter).Data(100)
      J = 0
    End If
    J = J + 1
    If J > UBound(MacroLst(hCounter).Data) Then ReDim Preserve MacroLst(hCounter).Data(J * 2)
    TempList(I).Name = myFile
    MacroLst(hCounter).Data(J) = TempList(I)
  Next I
  ReDim Preserve MacroLst(hCounter).Data(J)

  On Error Resume Next
  
  ' Create the file cache
  myDirs = Options.MacroPath
  myDirNeatList = myDirs

  Do
    myDir = ReadNext(myDirs, ";")
    If myDirs = "" Then Exit Do
    If myDir <> "" Then
      hCounter = hCounter + 1
      If hCounter > UBound(MacroLst) Then ReDim Preserve MacroLst(hCounter + 5)
      ReDim MacroLst(hCounter).Data(100)
      MacroLst(hCounter).DirName = ReadNext(myDirNeatList, ";")
      I = 0
      For J = 1 To 2
        myStr = Dir$(AddDir(myDir, Choose(J, "*.api", "*.scm")))
        Do Until myStr = ""
          myFile = GetRealName(AddDir(myDir, myStr))
          I = I + 1
          If I > UBound(MacroLst(hCounter).Data) Then ReDim Preserve MacroLst(hCounter).Data(I * 2)
          With MacroLst(hCounter).Data(I)
            .File = myFile
            .Name = GetFileTitle(myFile)
          End With
          myStr = Dir$()
        Loop
      Next J
      
      myStr = Dir$(AddDir(myDir, "*"), vbDirectory)
      Do Until myStr = ""
        If myStr <> "." And myStr <> ".." Then
          myFile = GetRealName(AddDir(myDir, myStr))
          On Error Resume Next
          If (GetAttr(myFile) And vbDirectory) = 0 Then
            ' Error goes here if necessary
          Else
            myDirs = myFile & ";" & myDirs
            myDirNeatList = MacroLst(hCounter).DirName & "\>" & myStr & ";" & myDirNeatList
          End If
        End If
        myStr = Dir$()
      Loop
  
      ReDim Preserve MacroLst(hCounter).Data(I)
      If I > 0 Then
        MacroQuickSort MacroLst(hCounter).Data, 1, I
      Else
        hCounter = hCounter - 1
      End If
    End If
  Loop
  SetScreenMousePointer vbDefault

  ReDim Preserve MacroLst(hCounter)
  
  Exit Sub
MacroListError:
  MsgBoxEx frmMain, "Macro list parse error: " & Error$ & vbCrLf & "in" & vbCrLf & FileList(I), vbExclamation, 0
  Resume Next
End Sub

' Given a file, if the file is in R8 texture format,
' convert to BMP format in the specified file name,
' otherwise, just copy the file
Public Sub LoadTexture(ByVal Filename As String, ByRef Result As String, Optional ByRef Width As Long, Optional ByRef Height As Long)
  Dim I As Long, FileNum As Integer, _
    TempLong As Long, _
    Mask1 As Long, Mask2 As Long, Mask3 As Long, _
    TempStr(1) As Byte, Header() As Byte, _
    RawBits() As Byte, NewBits() As Byte

  Dim Pic As PicType, Res As Long
  
  If Not FileExists(Filename) Then Exit Sub

  SetScreenMousePointer vbHourglass

  If Result = "" Then Result = AddDir(TempPathName, "fssc.bmp")
  If FileExists(Result) Then Kill Result
  
  Res = AnytoBmps(Filename, Result, Pic, 0, 0)
  
  If Res > 0 Then
    Width = Pic.Width
    Height = Pic.Height
    GoTo LoadTextureError:
  End If
  On Error GoTo LoadTextureError:
  
  ' Otherwise, use the old load routine

  FileNum = FreeFile

  Open Filename For Binary As #FileNum

  If LOF(FileNum) = 65536 Then
    ' Raw texture format
    ReDim RawBits(65535)
    Get #FileNum, 1, RawBits
    Close #FileNum
    Open Result For Binary As #FileNum
    Put #FileNum, , TextureHeader
    Put #FileNum, 1079, RawBits
    Close #FileNum
    Width = 256
    Height = 256
  Else
    Get #FileNum, 1, TempStr
    If TempStr(0) = 66 And TempStr(1) = 77 Then
      ' Bitmap
      Get #FileNum, 31, TempLong ' Compression
      Get #FileNum, 55, Mask1    ' Red
      Get #FileNum, 59, Mask2    ' Green
      Get #FileNum, 63, Mask3    ' Blue

      If TempLong = 3 And Mask1 = 3840 And Mask2 = 240 And Mask3 = 15 Then
        ' 4444 format
        Get #FileNum, 19, Width
        Get #FileNum, 23, Height

        ReDim RawBits(Width * Height * 2 - 1)
        ReDim NewBits(Width * Height * 3 - 1)

        Get #FileNum, 87, RawBits
        Close #FileNum
        Header = LoadResData("DATA2", 10)
        Open Result For Binary As #FileNum
        Put #FileNum, 1, Header
        Put #FileNum, 19, Width
        Put #FileNum, 23, Height

        For I = 0 To Width * Height - 1
          NewBits(I * 3) = (RawBits(I * 2) And 15) * 16
          NewBits(I * 3 + 1) = (RawBits(I * 2) \ 16) * 16
          NewBits(I * 3 + 2) = (RawBits(I * 2 + 1) And 15) * 16
        Next I
        Put #FileNum, 75, NewBits
        Close #FileNum
      ElseIf TempLong = 3 And Mask1 = 63488 And Mask2 = 2016 And Mask3 = 31 Then
        ' 565 format
        Get #FileNum, 19, Width
        Get #FileNum, 23, Height

        ReDim RawBits(Width * Height * 2 - 1)
        ReDim NewBits(Width * Height * 3 - 1)

        Get #FileNum, 87, RawBits
        Close #FileNum
        Header = LoadResData("DATA2", 10)
        Open Result For Binary As #FileNum
        Put #FileNum, 1, Header
        Put #FileNum, 19, Width
        Put #FileNum, 23, Height

        For I = 0 To Width * Height - 1
          NewBits(I * 3) = (RawBits(I * 2) And 31) * 8
          NewBits(I * 3 + 1) = ((RawBits(I * 2) \ 32) + ((RawBits(I * 2 + 1) And 7) * 8)) * 4
          NewBits(I * 3 + 2) = ((RawBits(I * 2 + 1) And 248) \ 8) * 8
        Next I
        Put #FileNum, 75, NewBits
        Close #FileNum
'      ElseIf TempLong = 827611204 And Mask1 = 808932166 And Mask2 = 20 And Mask3 = 256 Then
      Else
        Close #FileNum
        ' Assume a VB supported bitmap format
        'FileCopy Filename, Result
        GoTo DefaultBitmap:
      End If
    Else
DefaultBitmap:
      Close #FileNum
      On Error Resume Next
      frmMain.picTemp = LoadPicture(Filename)
      If Err = 0 Then
        SavePicture frmMain.picTemp.Image, Result
      End If
      Width = frmMain.picTemp.Width
      Height = frmMain.picTemp.Height
      frmMain.picTemp = LoadPicture()
    End If
  End If
LoadTextureError:
  SetScreenMousePointer vbDefault
End Sub

' Sort a list
Private Sub MacroQuickSort(SortList() As MacroType, ByVal First As Integer, ByVal Last As Integer)
  Dim Low As Integer, High As Integer, _
    Temp As MacroType, TestElement As MacroType
  If First > Last Then Exit Sub
  Low = First
  High = Last
  TestElement = SortList((First + Last) / 2)  'Select an element from the middle.
  Do
    Do While SortList(Low).Name < TestElement.Name     'Find lowest element that is >= TestElement.
      Low = Low + 1
    Loop
    Do While SortList(High).Name > TestElement.Name    'Find highest element that is <= TestElement.
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
  If (First < High) Then MacroQuickSort SortList, First, High
  If (Low < Last) Then MacroQuickSort SortList, Low, Last
End Sub

' Given a string to match and a set of strings, return the
' index of the array that matches the given string
Public Function MatchText(ByVal StringToMatch As String, ParamArray TextArray() As Variant) As Integer
  Dim I As Integer
  For I = 0 To UBound(TextArray)
    If StrComp(StringToMatch, TextArray(I), vbTextCompare) = 0 Then MatchText = I: Exit Function
  Next I
  MatchText = ValEx(StringToMatch)
End Function

' Displays a message box
' Note that the VB message boxes do not have an owner
' and WM_PAINT events are not sent.
Public Function MsgBoxEx(X As Form, ByVal Mess As String, ByVal Flags As VbMsgBoxStyle, ByVal HelpID As Integer) As VbMsgBoxResult
  'MsgBoxHelpID = HelpID
  If frmSplash.isShown Then frmSplash.Hide
  Dim sMsg As MSGBOXPARAMS
  sMsg.cbSize = Len(sMsg)
  sMsg.dwContextHelpId = HelpID
  sMsg.dwLanguageId = LOCALE_USER_DEFAULT
  'sMsg.dwStyle = Flags Or IIf(HelpID > 0, vbMsgBoxHelpButton, 0)
  sMsg.dwStyle = Flags
  sMsg.hwndOwner = X.hwnd
  'sMsg.lpfnMsgBoxCallback = DoAddressOf(AddressOf DoHelp)
  sMsg.lpszCaption = App.Title
  sMsg.lpszText = Mess
  MsgBoxEx = MessageBoxIndirect(sMsg)
  'MsgBoxEx = MessageBox(X.hwnd, Mess, App.Title, Flags Or IIf(HelpID > 0, vbMsgBoxHelpButton, 0))
  If frmSplash.isShown Then frmSplash.Show: frmSplash.Refresh
  'MsgBoxHelpID = 0
End Function

Private Function DoAddressOf(ByVal X As Long)
  DoAddressOf = X
End Function

' Help button in Messagebox
Private Sub DoHelp(hi As HELPINFO)
  HtmlHelpLong frmMain.hwnd, Lang.HelpFile & ">Error", HH_HELP_CONTEXT, hi.dwContextId
End Sub

' Delete all the autosaved files previously created
Public Sub PurgeAutoSaves()
  Dim I As Integer
  For I = 0 To 3
    If AutoSaveFiles(I) <> "" Then Kill AutoSaveFiles(I)
  Next I
  AutoSavePointer = 0
  Erase AutoSaveFiles
End Sub

' Replaces Parenthesis () with Braces {}
Public Function ReplaceParens(ByVal X As String)
  ReplaceParens = Replace(Replace(X, "(", "{"), ")", "}")
End Function

' Reverse Runway ID: 17R -> 35L
Public Function ReverseRunway(ByVal RW As String)
  Dim Temp As String
  Temp = (Val(RW) + 17) Mod 36 + 1
  Select Case Right$(RW, 1)
    Case "L": Temp = Temp & "R"
    Case "R": Temp = Temp & "L"
    Case "C": Temp = Temp & "C"
  End Select
  ReverseRunway = Temp
End Function

' Run a Dos EXE File
Public Sub RunDosFile(ByVal App As String, ByVal CmdLine As String, ByVal WorkDir As String)
  Dim sinfo As STARTUPINFO, pinfo As PROCESS_INFORMATION

  sinfo.cb = Len(sinfo)
  sinfo.dwFlags = STARTF_USESHOWWINDOW
  sinfo.wShowWindow = vbNormalFocus
  
  If CreateProcess(vbNullString, QuoteString(App) & " " & CmdLine, ByVal 0&, ByVal 0&, False, NORMAL_PRIORITY_CLASS, ByVal 0&, WorkDir, sinfo, pinfo) Then
    WaitForSingleObject pinfo.hProcess, INFINITE
  End If
End Sub

' Runs SCASM/FreeSC with SCAFile as the source and
' BGLFile as the destination
Public Function RunSCASM(ByVal SCAFile As String, ByVal BGLFile As String) As Boolean
  If FileExists("scaerror.log") Then Kill "scaerror.log"
  RunDosFile Options.Compiler, QuoteString(SCAFile) & " " & QuoteString(BGLFile) & " -l", GetDir(SCAFile)
  
  If FileExists("scaerror.log") Then
    If FileLen("scaerror.log") <> 30 Then
      SetScreenMousePointer vbDefault
      If MsgBoxEx(frmMain, Lang.GetString(RES_ERR_Compile), vbCritical Or vbYesNo, RES_ERR_Compile) = vbYes Then _
        Shell Options.TextEditor & " scaerror.log", vbNormalFocus
      RunSCASM = False
    Else
      Kill "scaerror.log"
      If Not Options.KeepSourceFile Then Kill SCAFile
      If Options.FSVersion >= Version_FS2K And Options.AutoCompress And FileExists(Options.CompressName) Then
        RunDosFile Options.CompressName, "-f " & QuoteString(BGLFile), GetDir(BGLFile)
      End If
      RunSCASM = True
    End If
  End If
End Function

' Search file for a meta variable and return the
' data string
Public Function SearchMeta(ByVal Filename As String) As MacroMetaType
  Dim FileNum As Integer, Cnt As Integer, _
    Temp As String, Y1 As Long, Y2 As Long, _
    Key As String, Result As MacroMetaType
  
  Cnt = 0
  FileNum = FreeFile
  
  On Error Resume Next
  Open Filename For Input As #FileNum
  If Err Then Exit Function
  Do While Not EOF(FileNum) And Cnt < 100
    Cnt = Cnt + 1
    LineInputEx FileNum, Temp
    
    If Left$(Temp, 1) = ";" Then
      If Left$(Temp, 23) = "; standard size / scale" Then
        Result.AirportScale = ReadNext(LTrim$(Mid$(Temp, 24)), " ")
      ElseIf Left$(Temp, 15) = "; standard size" Then
        Result.APTMacroScale = ReadNext(LTrim$(Mid$(Temp, 16)), " ")
      Else
        Y1 = InStr(Temp, " ")
        Y2 = InStr(Temp, ",")
        Temp = Mid$(Temp, 2)
        Key = ""
        If Y1 = 0 Then Y1 = 999
        If Y2 = 0 Then Y2 = 999
        If Y1 < Y2 Then
          If Y1 > 1 Then Key = ReadNext(Temp, " ")
        Else
          If Y2 > 1 Then Key = ReadNext(Temp, ",")
        End If
        
        Select Case LCase$(Key)
          Case "voddata":   Result.VODData = Temp
          Case "macrodesc": Result.MacroDesc = Temp
          Case "defaults": Result.Defaults = Temp
          Case "scale": Result.MScale = Temp
          Case "points": Result.Points = Temp
          Case "quickpoints": Result.QuickPoints = Temp
          Case "defaultscale": Result.DefaultScale = Temp
          Case "defaultparams": Result.DefaultParams = Temp
          Case "paramdescr": Result.ParamDescr = Temp
          Case "defaultrange": Result.DefaultRange = Temp
          Case "defaultdensity": Result.DefaultDensity = Temp
          Case "designshape": Result.DesignShape = Temp
          Case "textures": Result.Textures = Temp
          Case "rgbenabled": Result.RGBEnabled = Temp
        End Select
      End If
    End If
  Loop
  Close #FileNum
  SearchMeta = Result
End Function

' Set the listindex of a combobox, or change the
' text
Public Sub SetComboText(Ctrl As ComboBox, ByVal Match As String)
  Const CB_ERR = (-1)
  Const CB_FINDSTRINGEXACT = &H158

  Dim X As Integer
  With Ctrl
    X = SendMessageStr(.hwnd, CB_FINDSTRINGEXACT, -1, Match)
    If X <> CB_ERR Then
      .ListIndex = X
    Else
      .Text = Match
    End If
  End With
End Sub

' Set the textbox's enabled state with GUI effects
Public Sub SetEnabled(Ctrl As TextBox, ByVal Enabled As Boolean)
  With Ctrl
    .BackColor = IIf(Enabled, vbWindowBackground, vbButtonFace)
    .Locked = Not Enabled
  End With
End Sub

' Set the screen pointer (prevents setting pointer
' to default when another procedure still requires
' hourglass)
Public Sub SetScreenMousePointer(ByVal Value As MousePointerConstants)
  If Value = vbDefault Then
    If ScreenMouseCount > 0 Then ScreenMouseCount = ScreenMouseCount - 1
    If ScreenMouseCount = 0 Then Screen.MousePointer = vbDefault
  ElseIf Value = vbHourglass Then
    ScreenMouseCount = ScreenMouseCount + 1
    If Screen.MousePointer <> vbHourglass Then Screen.MousePointer = vbHourglass
  ElseIf Value = -1 Then
    ScreenMouseCount = 0
  Else
    ScreenMouseCount = 0
    Screen.MousePointer = Value
  End If
End Sub

' Sub SmartSelectText
' Smart Selects the text of a text box
Public Sub SmartSelectText(Ctrl As TextBox)
  Dim Temp As String, I As Integer, Y As Integer
  With Ctrl
    Temp = .Text
    For I = Len(Temp) To 1 Step -1
      If InStr(DigitsDecimal, Mid$(Temp, I, 1)) > 0 Then Y = I: Exit For
    Next I
    If Y = 0 Then Y = Len(Temp)
    .SelStart = 0
    .SelLength = Y
  End With
End Sub

' Validates data values based on tag properties
Public Function Validate(X As TextBox, ByRef Msg As String, ByRef Value As Single, Optional ByVal LowBound As Single = -999999, Optional ByVal UpBound As Single = -999999) As Boolean
  Dim Lower As Single, Upper As Single, _
    Units As Integer, ID As Integer, _
    ErrValue As Integer, Descriptor As String, _
    strValue As String, numValue As Single, _
    ConvFactor As Single, ConvCorrection As Single
    
  strValue = X.Text
  ID = Val(X.Tag)
  If ID = 0 Then Validate = True: Exit Function
  
  If strValue = "" And MultiSelection Then Validate = True: Exit Function

  Select Case ID
    Case RES_LBL_X, RES_LBL_Y
      Lower = -100000: Upper = 100000
    Case RES_LBL_Rotation
      Lower = 0: Upper = 359.99
      Units = RES_Unit_Deg
      ConvFactor = 1
      UserToGeographic strValue, numValue, ErrValue
    Case RES_Shp_Scale
      Lower = 0.1: Upper = 20
    Case RES_Shp_Spacing
      Lower = 1: Upper = 32767
    Case RES_Shp_LineWidth, RES_Shp_LineObjWidth
      Lower = 0.1: Upper = 255
    Case RES_Shp_Width
      Lower = 0.2: Upper = 255
    Case RES_Shp_V1
      Lower = 0: Upper = 100000
    Case RES_Shp_Z
      Lower = 0: Upper = 32767
    Case RES_Shp_ArcRadius
      Lower = 0: Upper = 100
    Case RES_Macro_Range
      Lower = 1: Upper = 100
      Units = RES_Unit_Nm
    Case RES_Macro_Scale
      Lower = 0.0001: Upper = 20
    Case RES_Macro_Altitude
      Lower = 0: Upper = 32767
    Case RES_Macro_V1
      Lower = 0: Upper = 100000
    Case RES_Macro_V2
      Lower = 0: Upper = 100000
    Case RES_Bldg_Altitude, RES_Tow_Height
      Lower = 0: Upper = 32767
    Case RES_Bldg_Length, RES_Bldg_Width
      Lower = 1: Upper = 32767
    Case RES_Bldg_Height, RES_Bldg_RLength, RES_Bldg_RWidth
      Lower = 0: Upper = 32767
    Case RES_Bldg_Repeat
      Lower = 0.01: Upper = 100
      Units = -2
    Case RES_Rdo_Range
      Lower = 1: Upper = 255
      Units = RES_Unit_Nm
    Case RES_Rdo_FrequencyVOR
      Lower = 108: Upper = 117.95
      Units = RES_Unit_Mhz
    Case RES_Rdo_FrequencyNDB
      Lower = 200: Upper = 999
      Units = RES_Unit_Khz
    Case RES_Rdo_BeamWidth
      Lower = 2: Upper = 9.9
      Units = RES_Unit_Deg
    Case RES_Rwy_Length, RES_Rwy_Width, RES_Rwy_VDistance
      Lower = 10: Upper = 7500
    Case RES_Rwy_HDistance, RES_Rwy_SignsOffset
      Lower = 10: Upper = 250
    Case RES_Rwy_Threshold_Length, RES_Rwy_Overrun_Length
      Lower = 0: Upper = 1000
    Case RES_Rwy_Strobes
      Lower = 0: Upper = 15
      Units = -2
    Case RES_Rwy_RowSeparation
      Lower = 0: Upper = 1000
    Case RES_Rwy_GlideSlope
      Lower = 2: Upper = 9
      Units = RES_Unit_Deg
    Case RES_Rwy_InnerMarker, RES_Rwy_MiddleMarker, RES_Rwy_OuterMarker
      Lower = 0.1: Upper = 20
      Units = RES_Unit_Nm
    Case RES_Rdo_FrequencyATIS, RES_Tow_Frequency1, RES_Tow_Frequency1 + 1, RES_Tow_Frequency1 + 2, _
        RES_Tow_Frequency1 + 3, RES_Tow_Frequency1 + 4, RES_Tow_Frequency1 + 5, RES_Tow_Frequency1 + 6, _
        RES_Tow_Frequency1 + 7, RES_Tow_Frequency1 + 8, RES_Tow_Frequency1 + 9, _
        RES_Tow_Frequency1 + 10, RES_Tow_Frequency1 + 11
      Lower = 118: Upper = 136.975
      Units = RES_Unit_Mhz
    Case RES_Hdr_Horizontal, RES_Hdr_Vertical
      Lower = 1: Upper = 200000
    Case RES_Hdr_MagVar
      Lower = -359.99: Upper = 359.99
      Units = RES_Unit_Deg
    Case RES_Hdr_Altitude, RES_Shp_Altitude, RES_Shp_FlatAltitude
      Lower = IIf(Options.FSVersion >= Version_FS2K, -2000, 0)
      Upper = 32767
    Case RES_Back_ZoomX, RES_Back_ZoomY
      Lower = 0.001: Upper = 1000
    Case RES_Exc_Horz, RES_Exc_Vert
      Lower = 1: Upper = 10000
    Case RES_Cde_Horz, RES_Cde_Vert
      Lower = 0.1: Upper = 10000
    Case RES_Suf_Height
      Lower = 0: Upper = 10000
    Case RES_Trans_Rotate, RES_Hdr_Rotation
      Lower = -359.99: Upper = 359.99
      numValue = ValEx(strValue)
      Units = RES_Unit_Deg
      ConvFactor = 1
    Case RES_OPT_lblAutoSave
      Lower = 0: Upper = 30
      Units = RES_Unit_Min
    Case RES_OPT_lblGrid
      Lower = 0: Upper = 10000
    Case RES_Zoom_Value
      Lower = 0.01: Upper = 64
    Case Else
      Units = -1
  End Select
  
  If LowBound > -999990 Then Lower = LowBound
  If UpBound > -999990 Then Upper = UpBound

  If Units = 0 Then
    ' Need to do conversion
    UserToMeter strValue, Units, ConvFactor, ConvCorrection, ErrValue
    numValue = ValEx(strValue)
    If ErrValue = 0 Then
      Lower = Lower * ConvFactor + ConvCorrection
      Upper = Upper * ConvFactor + ConvCorrection
    End If
  ElseIf Units = RES_Unit_Nm Then
    ' Need to do conversion
    UserToNautical strValue, Units, ConvFactor, ConvCorrection, ErrValue
    numValue = ValEx(strValue)
    If ErrValue = 0 Then
      Lower = Lower * ConvFactor + ConvCorrection
      Upper = Upper * ConvFactor + ConvCorrection
    End If
  ElseIf Units <> -1 Then
    ' Numeric non-conversions
    
    ' Fill numvalue if numvalue wasn't filled
    If ConvFactor = 0 Then
      numValue = ValEx(strValue)
      ConvFactor = 1
    End If
  
    If InStr(DigitsDecimalSigns, Left$(strValue, 1)) = 0 Or strValue = "" Then
      ' Non numeric first character
      ErrValue = RES_ERR_Numeric
    End If
  Else ' Non-numerics
    ErrValue = -1
        
    Select Case ID
      Case RES_Rdo_Name
        If strValue = "" Then ErrValue = RES_ERR_SpecifyName
      Case RES_Rdo_ID
        If strValue = "" Then ErrValue = RES_ERR_SpecifyID
      Case RES_Rdo_Text
        If strValue = "" Then ErrValue = RES_ERR_SpecifyText
      Case RES_Rdo_Runway, RES_Rwy_ID
        ' This condition is already tested on input
        
        ' Or InStr("LRC0123456789", Right$(strValue, 1)) = 0
        If Not Between(ValEx(strValue), 1, 36) Then ErrValue = RES_ERR_RunwayID
      Case RES_Back_Image
        If strValue = "" Then ErrValue = RES_ERR_SpecifyFile
      Case RES_ERR_SpecifyText
        If strValue = "" Then ErrValue = RES_ERR_SpecifyText
    End Select
  End If

  Descriptor = Replace(Replace(Lang.GetString(ID), "&", ""), ":", "")
  Select Case ErrValue
    Case -1
      Validate = True
    Case 0
      If Not Between(numValue, Lower, Upper) Then
        Msg = Lang.FormatErrorMessage(Descriptor, Lower, Upper, Units)
      Else
        Validate = True
        Value = numValue * (1 / ConvFactor) + ConvCorrection
      End If
    Case RES_ERR_Numeric
      Msg = Lang.ResolveString(RES_ERR_Numeric, Descriptor)
    Case RES_ERR_Units, RES_ERR_RotUnits
      Msg = Lang.ResolveString(ErrValue, GetUnit(strValue), Descriptor)
    Case Else
      Msg = Lang.GetString(ErrValue)
  End Select
End Function
