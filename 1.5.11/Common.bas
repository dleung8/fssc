Attribute VB_Name = "basCommon"
' [basCommon]
' Common Functions and Procedures

Option Explicit

Public Enum RootKeyEnum
  HKEY_CLASSES_ROOT = &H80000000
  HKEY_CURRENT_USER = &H80000001
  HKEY_LOCAL_MACHINE = &H80000002
End Enum

Public Enum LocaleInfoEnum
  LOCALE_SABBREVLANGNAME = &H3  '  abbreviated language name
  LOCALE_SLIST = &HC            '  list item separator
  LOCALE_IMEASURE = &HD         '  0 = metric, 1 = US
  LOCALE_SDECIMAL = &HE         '  decimal separator
  LOCALE_SPOSITIVESIGN = &H50   '  positive sign
  LOCALE_SNEGATIVESIGN = &H51   '  negative sign
End Enum

Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Private Const MAX_PATH = 260

Private Type WIN32_FIND_DATA
  dwFileAttributes As Long
  ftCreationTime As FILETIME
  ftLastAccessTime As FILETIME
  ftLastWriteTime As FILETIME
  nFileSizeHigh As Long
  nFileSizeLow As Long
  dwReserved0 As Long
  dwReserved1 As Long
  cFileName As String * MAX_PATH
  cAlternate As String * 14
End Type

Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

' Center form function
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_GETWORKAREA = 48

' Menu Functions
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Const MF_BYPOSITION = &H400&

' File Functions
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Const INVALID_HANDLE_VALUE = -1

' Registry Functions
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Const ERROR_SUCCESS = 0&
Private Const REG_SZ = 1
Private Const REG_DWORD = 4

' Center text function
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

' Regional setting function
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Const LOCALE_USER_DEFAULT = &H400

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal lSize As Long, ByVal lpFileName As String) As Long

Public Const BaseKey = "Software\Leung"
Public Const GeneralKey = "Software\Leung\General\4.x"

Public AppKey As String
Public CopyrightYear As String

' Function AddDir
' Adds a directory and file together adding "\" if necessary
' Directory = Dir to start with
' File = File to add
' Returns the resulting string
Public Function AddDir(ByVal Directory As String, ByVal File As String) As String
  AddDir = Directory & IIf(Right$(Directory, 1) <> "\", "\", "") & File
End Function

' Function Between
' Returns true if num is between min and max
Public Function Between(ByVal Num As Single, ByVal Min As Single, ByVal Max As Single) As Boolean
  Between = (Min <= Num) And (Num <= Max)
End Function

' Convert a binary string to an integer
Public Function BinaryToDec(ByVal Dat As String) As Long
  Dim I As Long, Temp As Long
  For I = 1 To Len(Dat)
    Temp = Temp * 2 - (Mid$(Dat, I, 1) = "1")
  Next I
  BinaryToDec = Temp
End Function

' Sub CenterForm
' Moves the form slightly above the center of the screen
' X = Form to center
Public Sub CenterForm(X As Form)
  Dim ScrRect As RECT, Left As Long, Top As Long
  
  If SystemParametersInfo(SPI_GETWORKAREA, 0, ScrRect, 0) = 0 Then
     ' Call failed - just use standard screen:
     Left = (Screen.Width - X.Width) / 2
     Top = (Screen.Height - X.Height) / 2
  Else
    With ScrRect
      Left = (.Right + .Left) / 2 * Screen.TwipsPerPixelX - X.Width / 2
      Top = ((.Bottom + .Top) / 2 * Screen.TwipsPerPixelX - X.Height / 2) * 0.85
    End With
  End If
  If Left < 0 Then Left = 0
  If Top < 0 Then Top = 0
  X.Move Left, Top
End Sub

' Center text inside a picturebox
Public Sub CenterText(PicBox As PictureBox, ByVal Msg As String, Optional ByVal Left As Long = 0, Optional ByVal Y As Long = -999)
  With PicBox
    If Y = -999 Then Y = .CurrentY
    TextOut .hdc, (.ScaleWidth - Left - .TextWidth(Msg)) / 2 + Left, Y, Msg, Len(Msg)
  End With
End Sub

' Sub DialogMenus
' Removes the separator in the dialog boxes
' X = Form to remove menu from
Public Sub DialogMenus(X As Form, Optional ByVal NoClose As Boolean = False)
  If NoClose Then RemoveMenu GetSystemMenu(X.hwnd, 0), 6, MF_BYPOSITION
  RemoveMenu GetSystemMenu(X.hwnd, 0), 5, MF_BYPOSITION
End Sub

' Sub DrawSymbolBox
' Draws Derek's Programs 2000 standard symbol button
Public Sub DrawSymbolBox(Ctrl As PictureBox, ByVal Depressed As Boolean)
  Dim Color1 As Long, Color2 As Long
  
  If Not Depressed Then
    Color1 = vb3DHighlight
    Color2 = vb3DLight
  Else
    Color1 = vbBlack
    Color2 = vbButtonShadow
  End If
  
  Ctrl.Cls
  Ctrl.Line (0, 0)-(18, 18), vbButtonFace, BF
  Ctrl.Line (1, 17)-(18, 17), vbButtonShadow
  Ctrl.Line (17, 1)-(17, 18), vbButtonShadow
  Ctrl.Line (0, 0)-(18, 0), Color1
  Ctrl.Line (0, 0)-(0, 18), Color1
  Ctrl.Line (0, 18)-(19, 18), vbBlack
  Ctrl.Line (18, 0)-(18, 19), vbBlack
  Ctrl.Line (1, 1)-(17, 1), Color2
  Ctrl.Line (1, 1)-(1, 17), Color2
  TextOut Ctrl.hdc, (9.5 - Ctrl.TextWidth("<") / 2), (9.5 - Ctrl.TextHeight("<") / 2 - 1), "<", 1
End Sub

' Function FileExists
' Check if the file exists
' File = File to check for
Public Function FileExists(ByVal File As String) As Boolean
  Dim hFind As Long
  Dim Data As WIN32_FIND_DATA
  
  If File = "" Then Exit Function
  hFind = FindFirstFile(File, Data)
  If hFind = INVALID_HANDLE_VALUE Then
    FileExists = False
  Else
    FindClose hFind
    If Len(StripTerminator(Data.cFileName)) <= 4 Or Len(File) = 2 Then
      ' Assume DOS device (CON, LPT1, COM1, AUX, NUL, PRN)
      FileExists = False
    Else
      FileExists = True
    End If
  End If
End Function

' Function GetDir
' Retrieves Directory Part of a file name
' File = File Name
' Returns Dir
Public Function GetDir(ByVal File As String)
  Dim Y As Long
  Y = InStrRev(File, "\")
  If Y > 0 Then
    GetDir = Left$(File, Y - 1)
  Else
    GetDir = CurDir$
  End If
End Function

' Function GetFileTitle
' Retrieves the file name without the directory
' File = File Name
' Returns File Title
Public Function GetFileTitle(ByVal File As String)
  Dim Y As Long
  Y = InStrRev(File, "\")
  If Y > 0 Then
    GetFileTitle = Mid$(File, Y + 1)
  Else
    GetFileTitle = File
  End If
End Function

' Get a string from an INI file
Public Function GetINIString(ByVal Section As String, ByVal Key As String, ByVal Default As String, ByVal File As String) As String
  Dim strBuffer As String
  
  strBuffer = String$(255, 0)
    
  If GetPrivateProfileString(Section, Key, Default, strBuffer, 255, File) > 0 Then
    GetINIString = StripTerminator(strBuffer)
  Else
    GetINIString = Default
  End If
End Function

' Function GetRealName
' Retrives the real name (with correct cases) of a file or directory
' Path = Path to correct
' Returns corrected path
Public Function GetRealName(ByVal Path As String)
  Dim Running As String, _
    CurrentDir As String, _
    CurrentSegment As String, _
    hFind As Long, _
    Data As WIN32_FIND_DATA
    
  On Error Resume Next
  
  If Path = "" Then
    GetRealName = ""
    Exit Function
  End If
  
  If Len(Path) <= 3 Then
    GetRealName = UCase$(Left$(Path, 1)) & ":\"
    Exit Function
  End If
  
  CurrentDir = UCase$(Left$(Path, 2))
  Running = UCase$(Mid$(Path, 4))
  Do Until Running = ""
    CurrentSegment = ReadNext(Running, "\")
    
    hFind = FindFirstFile(CurrentDir & "\" & CurrentSegment, Data)
    If hFind = INVALID_HANDLE_VALUE Then
      CurrentDir = CurrentDir & "\" & CurrentSegment
      Err.Number = 1
      Exit Do
    Else
      CurrentDir = CurrentDir & "\" & StripTerminator(Trim$(Data.cFileName))
      FindClose hFind
    End If
  Loop
  If Err.Number = 0 Then
    GetRealName = CurrentDir
  Else
    GetRealName = Path
  End If
End Function

' Function GetShortName
' Retrieve the short pathname version of a path possibly
'   containing long subdirectory and/or file names
Public Function GetShortName(ByVal LongPath As String) As String
  Dim ShortPath As String

  ShortPath = String$(MAX_PATH + 1, 0)
  GetShortPathName LongPath, ShortPath, MAX_PATH
  
  GetShortName = StripTerminator(ShortPath)
End Function

' Function GetTempPathName
' Get the Windows Temp path
Public Function GetTempPathName() As String
  Dim Temp As String
  Temp = String$(MAX_PATH + 1, 0)
  GetTempPath MAX_PATH, Temp
  GetTempPathName = GetRealName(StripTerminator(Temp))
End Function

' Function LocaleInfo
' Get Local setting
Public Function LocaleInfo(ByVal LocaleType As LocaleInfoEnum) As String
  Dim Buffer As String
  Buffer = String$(MAX_PATH + 1, 0)
  GetLocaleInfo LOCALE_USER_DEFAULT, LocaleType, Buffer, MAX_PATH
  LocaleInfo = StripTerminator(Buffer)
End Function

' Function MakeFileNameNeat
' Strips the directory and extension off
' File = String to make neat
' Returns neat file name
' Uses: Application Title and Recent File List
Public Function MakeFileNameNeat(ByVal File As String) As String
  Dim Temp As String, Y As Long
  
  If Len(File) = 0 Then Exit Function
  Temp = GetFileTitle(File)

  If Len(Temp) <= 12 And StrComp(UCase$(Temp), Temp, vbBinaryCompare) = 0 And InStr(Temp, " ") = 0 Then
    Temp = UCase$(Left$(Temp, 1)) & LCase$(Mid$(Temp, 2))
  End If
  Y = InStr(Temp, ".")
  If Y > 0 Then
    MakeFileNameNeat = Left$(Temp, Y - 1)
  Else
    MakeFileNameNeat = Temp
  End If
End Function

'-----------------------------------------------------------
' FUNCTION: MakePathAux
'
' Creates the specified directory path.
'
' No user interaction occurs if an error is encountered.
' If user interaction is desired, use the related
'   MakePathAux() function.
'
' IN: [strDirName] - name of the dir path to make
'
' Returns: True if successful, False if error.
'-----------------------------------------------------------
'
Public Function MakePath(ByVal strDirName As String) As Boolean
  Dim strPath As String, intOffset As Integer, _
    intAnchor As Integer, strOldPath As String

  On Error Resume Next

  '
  'Add trailing backslash
  '
  strDirName = AddDir(strDirName, "")

  strOldPath = CurDir$
  intAnchor = 0

  '
  'Loop and make each subdir of the path separately.
  '
  intOffset = InStr(intAnchor + 1, strDirName, "\")
  intAnchor = intOffset 'Start with at least one backslash, i.e. "C:\FirstDir"
  Do
    intOffset = InStr(intAnchor + 1, strDirName, "\")
    intAnchor = intOffset

    If intAnchor > 0 Then
      strPath = Left$(strDirName, intOffset - 1)
      ' Determine if this directory already exists
      Err = 0
      ChDir strPath
      If Err Then
        ' We must create this directory
        Err = 0
        MkDir strPath
        If Err Then
          MsgBoxEx Screen.ActiveForm, Lang.ResolveString(RES_ERR_MakeFolderFail, Error$), vbCritical, RES_ERR_MakeFolderFail
          GoTo Done
        End If
      End If
    End If
  Loop Until intAnchor = 0

  MakePath = True
Done:
  ChangeDir strOldPath
End Function

Public Function Max(ByVal Val1 As Integer, ByVal Val2 As Integer)
  If Val1 > Val2 Then Max = Val1 Else Max = Val2
End Function

Public Function Min(ByVal Val1 As Integer, ByVal Val2 As Integer)
  If Val1 < Val2 Then Min = Val1 Else Min = Val2
End Function

' Function MultiDir
' A variant of Dir$ scanning multiple directories
Public Function MultiDir(Optional ByVal Filter As String, Optional ByVal Dirs As String) As String
  Static myFilter As String, myDirs As String, Cur As String
  Dim Result As String, One As String
  
  ' Nothing to do!
  If myFilter = "" And Filter = "" Then Exit Function
  
  If Filter <> "" Then
    ' New parameters
    myFilter = Filter
    
    ' Discard repeats
    Do Until Dirs = ""
      One = ReadNext(Dirs, ";") & ";"
      If InStr(UCase$(myDirs), UCase$(One)) = 0 Then myDirs = myDirs & One
    Loop
    
    Cur = ReadNext(myDirs, ";")
    Result = Dir$(AddDir(Cur, myFilter))
  Else
    Result = Dir$()
  End If
  
  ' Loop while there is no file and more dirs to do
  Do While Result = "" And myDirs <> ""
    Cur = ReadNext(myDirs, ";")
    Result = Dir$(AddDir(Cur, myFilter))
  Loop
  If Result = "" Then
    ' No more files to read. Reset parameters
    myFilter = ""
  Else
    Dim Y As Long
    Y = InStrRev(myFilter, "\")
    If Y > 0 Then
      Result = AddDir(Cur, Left$(myFilter, Y) & Result)
    Else
      Result = AddDir(Cur, Result)
    End If
  End If
  MultiDir = Result
End Function

' Function QuoteString
' Quote a string if quotes don't already exist
Public Function QuoteString(ByVal myStr As String)
  If Left$(myStr, 1) <> Chr$(34) Or Right$(myStr, 1) <> Chr$(34) Then
    QuoteString = Chr$(34) & myStr & Chr$(34)
  Else
    QuoteString = myStr
  End If
End Function

' Function ReadNext
' Reads the next item (separated by a delimiter) from
' a string and delete it from the string
Public Function ReadNext(ByRef Str As String, ByVal Delimiter As String) As String
  Dim Y As Long
  Y = InStr(1, Str, Delimiter, vbTextCompare)
  If Y = 0 Then
    ReadNext = Str
    Str = ""
  Else
    ReadNext = Left$(Str, Y - 1)
    Str = Trim$(Mid$(Str, Y + Len(Delimiter)))
  End If
End Function

' Function ReadLast
' Reads the last item (separated by a delimiter) from
' a string and delete it from the string
Public Function ReadLast(ByRef Str As String, ByVal Delimiter As String) As String
  Dim Y As Long
  Y = InStrRev(Str, Delimiter)
  If Y = 0 Then
    ReadLast = Str
    Str = ""
  Else
    ReadLast = Trim$(Mid$(Str, Y + Len(Delimiter)))
    Str = Left$(Str, Y - 1)
  End If
End Function

' Function RegGetKey
' Gets a value from the Registry
' Key = Key of value to get
' Default = Default value to return if error
' Optional Group = Group of the key (If first char = "\", then creates subfolder of default group
' Optional RootKey = Group of the Group Variable
' Returns value
Public Function RegGetKey(ByVal Key As String, ByVal Default As Variant, Optional ByVal Group As String, Optional ByVal RootKey As RootKeyEnum = HKEY_CURRENT_USER) As Variant
  Dim hKey As Long, LongData As Long, _
    IniStr As String, KeyFormat As Long, _
    KeyValSize As Long

  If Group = "" Then Group = AppKey
  If Left$(Group, 1) = "\" Then Group = AppKey & Group

  RegGetKey = Default
  If RegOpenKey(RootKey, Group, hKey) = ERROR_SUCCESS Then
    KeyValSize = 1024
    IniStr = String$(KeyValSize, 0)
    If RegQueryValueEx(hKey, Key, 0&, KeyFormat, ByVal IniStr, KeyValSize) = ERROR_SUCCESS Then
      If KeyFormat = REG_DWORD Then
        If RegQueryValueEx(hKey, Key, 0&, KeyFormat, LongData, 4) = ERROR_SUCCESS Then
          RegGetKey = LongData
        End If
      Else
        RegGetKey = StripTerminator(IniStr)
      End If
    End If
  End If
  RegCloseKey hKey
End Function

' Sub RegSetKey
' Sets a key in the Registry
' Key = Key to set value to
' Value = Value to set to Key
' Optional Group = Group of the key (If first char = "\", then creates subfolder of default group
' Optional RootKey = Group of the Group Variable
Public Sub RegSetKey(ByVal Key As String, ByVal Value As Variant, Optional ByVal Group As String, Optional ByVal RootKey As RootKeyEnum = HKEY_CURRENT_USER)
  Dim hKey As Long
  
  If Group = "" Then Group = AppKey
  If Left$(Group, 1) = "\" Then Group = AppKey & Group
  
  If RegCreateKey(RootKey, Group, hKey) = ERROR_SUCCESS Then
    Select Case VarType(Value)
      Case vbInteger, vbLong, vbBoolean
        RegSetValueEx hKey, Key, 0&, REG_DWORD, CLng(Value), 4
      Case Else
        RegSetValueEx hKey, Key, 0&, REG_SZ, ByVal CStr(Value) & vbNullChar, Len(CStr(Value)) + 1
    End Select
  End If
  RegCloseKey (hKey)
End Sub

' Function ReturnSymbol
' Returns a symbol when a user types a key combination
' KeyCode = The keycode from a KeyDown event/Keypress event
' Shift = The state of the Ctrl/Alt/Shift keys
' Status = 1 if KeyPress event
' Status = 2 to reset symbol buffer
' Returns a symbol or nothing
Public Function ReturnSymbol(ByVal KeyCode As Integer, ByVal Shift As Integer, Optional ByVal Status As Integer = 0) As Integer
  ' More KeyCodes
  Const vbKeyOpenSngQuote = 192  ' `
  Const vbKeySemiColon = 186     ' ;
  Const vbKeySngQuote = 222      ' '
  Const vbKeyComma = 188         ' ,
  Const vbKeyPeriod = 190        ' .
  Const vbKeySlash = 191         ' /
  
  Static Symbols As Integer
  Dim Result As Integer, DontUCase As Boolean
  
  ' Suppress "Ding" when keycode = 30: Ctrl+Shift+6
  If Status = 1 Then ReturnSymbol = IIf(KeyCode <> 30, KeyCode, 0)
  If Status = 2 Then Symbols = 0: Exit Function
  
  If Shift = (vbCtrlMask Or vbAltMask) Then
    ' Single key combination
    Select Case KeyCode
      Case vbKeyC: ReturnSymbol = 169      ' ©
      Case vbKeyR: ReturnSymbol = 174      ' ®
      Case vbKeyT: ReturnSymbol = 153      ' ™
      Case vbKeyPeriod: ReturnSymbol = 133 ' …
    End Select
  ElseIf Shift = (vbCtrlMask Or vbAltMask Or vbShiftMask) Then
    ' Single key combination
    Select Case KeyCode
      Case vbKey1:  ReturnSymbol = 161    ' ¡
      Case vbKeySlash: ReturnSymbol = 191 ' ¿
    End Select
  ElseIf Shift = vbCtrlMask Then
    ' Multiple key combination (First Key)
    Select Case KeyCode
      Case vbKeyOpenSngQuote: Symbols = 1
      Case vbKeySngQuote:     Symbols = 2
      Case vbKeyComma:        Symbols = 8
      Case vbKeySlash:        Symbols = 9
    End Select
  ElseIf Shift = (vbCtrlMask Or vbShiftMask) Then
    ' Multiple key combination (First Key)
    Select Case KeyCode
      Case vbKey6:            Symbols = 3
      Case vbKeyOpenSngQuote: Symbols = 4
      Case vbKeySemiColon:    Symbols = 5
      Case vbKey2:            Symbols = 6
      Case vbKey7:            Symbols = 7
    End Select
  ElseIf Status = 1 And Symbols > 0 And KeyCode <> 30 Then
    ' Keycode = 30 when user presses Ctrl+Shift+6 (?Why)
    ' Multiple key combination (Second Key)
    Select Case Symbols
      Case 1
        Select Case LCase$(Chr$(KeyCode))
          Case "a": Result = 192 ' À
          Case "e": Result = 200 ' È
          Case "i": Result = 204 ' Ì
          Case "o": Result = 210 ' Ò
          Case "u": Result = 217 ' Ù
          Case ">": Result = 187 ' »
          Case "<": Result = 171 ' «
        End Select
      Case 2
        Select Case LCase$(Chr$(KeyCode))
          Case " ": Result = 180 ' ´
          Case "a": Result = 193 ' Á
          Case "d": Result = 208 ' Ð
          Case "e": Result = 201 ' É
          Case "i": Result = 205 ' Í
          Case "o": Result = 211 ' Ó
          Case "u": Result = 218 ' Ú
          Case "y": Result = 221 ' Ý
        End Select
      Case 3
        Select Case LCase$(Chr$(KeyCode))
          Case " ": Result = 136 ' ˆ
          Case "a": Result = 194 ' Â
          Case "e": Result = 202 ' Ê
          Case "i": Result = 206 ' Î
          Case "o": Result = 212 ' Ô
          Case "s": Result = IIf(StrComp(Chr$(KeyCode), "S", vbBinaryCompare), 154, 138): _
                    DontUCase = True ' Š, š
          Case "u": Result = 219 ' Û
        End Select
      Case 4
        Select Case LCase$(Chr$(KeyCode))
          Case "a": Result = 195 ' Ã
          Case "n": Result = 209 ' Ñ
          Case "o": Result = 213 ' Õ
        End Select
      Case 5
        Select Case LCase$(Chr$(KeyCode))
          Case " ": Result = 168 ' ¨
          Case "a": Result = 196 ' Ä
          Case "e": Result = 203 ' Ë
          Case "i": Result = 207 ' Ï
          Case "o": Result = 214 ' Ö
          Case "u": Result = 220 ' Ü
          Case "y": Result = IIf(StrComp(Chr$(KeyCode), "Y", vbBinaryCompare), 255, 159): _
                    DontUCase = True ' Ÿ, ÿ
        End Select
      Case 6
        Select Case LCase$(Chr$(KeyCode))
          Case " ": Result = 176 ' °
          Case "a": Result = 197 ' Å
        End Select
      Case 7
        Select Case LCase$(Chr$(KeyCode))
          Case "a": Result = 198 ' Æ
          Case "o": Result = IIf(StrComp(Chr$(KeyCode), "O", vbBinaryCompare), 156, 140): _
                    DontUCase = True ' Œ, œ
          Case "s": Result = 223: DontUCase = True ' ß
        End Select
      Case 8
        Select Case LCase$(Chr$(KeyCode))
          Case " ": Result = 184 ' ¸
          Case "c": Result = 199 ' Ç
        End Select
      Case 9
        Select Case LCase$(Chr$(KeyCode))
          Case "c": Result = 162: DontUCase = True ' ¢
          Case "o": Result = 216 ' Ø
        End Select
    End Select
    If (Result <> 0) And (Not DontUCase) Then
      If StrComp(LCase$(Chr$(KeyCode)), Chr$(KeyCode), vbBinaryCompare) = 0 And StrComp(UCase$(Chr$(KeyCode)), Chr$(KeyCode), vbBinaryCompare) <> 0 Then _
        Result = Result + 32
    End If
    ReturnSymbol = Result
    Symbols = 0
  End If
End Function

' Sub SelectText
' Selects the text of a text box
Public Sub SelectText(X As TextBox)
  With X
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

' Function StripTerminator
' Strip chr$(0)s from a string
Public Function StripTerminator(ByVal Value As String) As String
  Dim Y As Long
  Y = InStr(Value, vbNullChar)
  If Y > 0 Then
    StripTerminator = Left$(Value, Y - 1)
  Else
    StripTerminator = Value
  End If
End Function

' Function UnQuoteString
' Takes a string and removes the surrounding quotation
' marks if they exist
Public Function UnQuoteString(ByVal myStr As String)
  If Left$(myStr, 1) = Chr$(34) And Right$(myStr, 1) = Chr$(34) Then
    UnQuoteString = Mid$(myStr, 2, Len(myStr) - 2)
  Else
    UnQuoteString = myStr
  End If
End Function
