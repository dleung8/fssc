Attribute VB_Name = "basCommon"
' [Common module]
' Common Functions and Procedures
Option Explicit
Option Compare Text

Public Enum RootKeyEnum
  HKEY_CLASSES_ROOT = &H80000000
  HKEY_CURRENT_USER = &H80000001
  HKEY_LOCAL_MACHINE = &H80000002
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

Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Const MF_BYPOSITION = &H400&

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Const INVALID_HANDLE_VALUE = -1

Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Const ERROR_SUCCESS = 0&
Private Const REG_SZ = 1
Private Const REG_DWORD = 4

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Const BaseKey = "Software\Leung"
Public Const GeneralKey = "Software\Leung\General\4.x"

Public AppKey As String
Public CpYear As String

' Function AddDir
' Adds a directory and file together adding "\" if necessary
' Directory = Dir to start with
' File = File to add
' Returns the resulting string
Public Function AddDir(Directory As String, File As String) As String
  AddDir = Directory & IIf(Right$(Directory, 1) <> "\", "\", "") & File
End Function

' Sub CenterForm
' Moves the form slightly above the center of the screen
' X = Form to center
Public Sub CenterForm(X As Form)
  X.Move Screen.Width / 2 - X.Width / 2, (Screen.Height / 2 - X.Height / 2) * 0.85
End Sub

' Sub DialogMenus
' Removes the separator in the dialog boxes
' X = Form to remove menu from
Public Sub DialogMenus(X As Form)
  RemoveMenu GetSystemMenu(X.hwnd, 0), 5, MF_BYPOSITION
End Sub

' Function FileExists
' Check if the file exists
' File = File to check for
Public Function FileExists(File As String) As Boolean
  Dim X As Integer
  On Error Resume Next
  
  X = FreeFile
  Open File For Input As #X
  FileExists = (Err = 0)
  Close #X
  Err = 0
End Function

' Function GetDir
' Retrieves Directory Part of a file name
' File = File Name
' Returns Dir
Public Function GetDir(File As String)
  Dim Y As Integer
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
Public Function GetFileTitle(File As String)
  Dim Y As Integer
  Y = InStrRev(File, "\")
  If Y > 0 Then
    GetFileTitle = Mid$(File, Y + 1)
  Else
    GetFileTitle = File
  End If
End Function

' Function GetRealName
' Retrieves the real name (with correct cases) of a file
' File = UserInputed file name
' Returns corrected file name
Public Function GetRealName(ByVal File As String)
  Dim X As Long, Dat As WIN32_FIND_DATA, Temp As String
  If FileExists(File) Then
    X = FindFirstFile(File, Dat)
    Temp = AddDir(GetRealDirName(GetDir(File)), IIf(X <> INVALID_HANDLE_VALUE, Trim$(Dat.cFileName), File))
    FindClose X
    GetRealName = Left$(Temp, InStr(Temp, Chr$(0)) - 1)
  Else
    GetRealName = File
  End If
End Function

' Retrieve the short pathname version of a path possibly
'   containing long subdirectory and/or file names
Public Function GetShortName(strLongPath As String) As String
  Dim strShortPath As String

  On Error GoTo 0
  strShortPath = String(256, Chr$(0))
  If GetShortPathName(strLongPath, strShortPath, 256) = 0 Then
    GetShortName = ""
  Else
    GetShortName = StripTerminator(strShortPath)
  End If
End Function

' Function GetRealDirName
' Retrives the real name (with correct cases) of a directory
' Path = Path to correct
' Returns corrected path
Public Function GetRealDirName(Path As String)
  Dim Running As String, Temp As String, _
    CurrentDir As String, X As Long, _
    Dat As WIN32_FIND_DATA, Ret As String, _
    Y As Integer
    
  On Error Resume Next
  
  If Path = "" Then GetRealDirName = "": Exit Function
  If Len(Path) <= 3 Then GetRealDirName = UCase$(Left$(Path, 1)) & ":\": Exit Function
  Running = Mid$(Path, 4)
  CurrentDir = UCase$(Left$(Path, 2))
  Do Until Running = ""
    Y = InStr(Running, "\")
    If Y > 0 Then
      Temp = UCase$(Left$(Running, Y - 1))
      Running = Mid$(Running, Y + 1)
    Else
      Temp = UCase$(Running)
      Running = ""
    End If
    
    X = FindFirstFile(AddDir(CurrentDir, Temp), Dat)
    Ret = AddDir(CurrentDir, IIf(X <> INVALID_HANDLE_VALUE, Trim$(Dat.cFileName), Temp))
    FindClose X
    CurrentDir = Left$(Ret, InStr(Ret, Chr$(0)) - 1)
  Loop
  If Err = 0 Then
    GetRealDirName = CurrentDir
  Else
    GetRealDirName = Path
  End If
End Function

' Function MakeFileNameNeat
' Strips the directory and extension off
' File = String to make neat
' Returns neat file name
' Uses: Application Title and Recent File List
Public Function MakeFileNameNeat(File As String) As String
  Dim Temp As String, Y As Integer
  If Len(File) = 0 Then Exit Function
  Temp = GetFileTitle(File)

  Y = InStr(Temp, ".")
  If Len(Temp) <= 12 And UCase$(Temp) = Temp And InStr(Temp, " ") = 0 Then
    If Y > 0 Then
      MakeFileNameNeat = Left$(UCase$(Left$(Temp, 1)) & LCase$(Mid$(Temp, 2)), Y - 1)
    Else
      MakeFileNameNeat = UCase$(Left$(Temp, 1)) & LCase$(Mid$(Temp, 2))
    End If
  Else
    If Y > 0 Then
      MakeFileNameNeat = Left$(Temp, Y - 1)
    Else
      MakeFileNameNeat = Temp
    End If
  End If
End Function

' Reads the next item (separated by a delimiter) from
' a string and delete it from the string
Public Function ReadNext(Str As String, Delimiter As String, Optional TrimStr As Boolean) As String
  Dim Y As Long
  Y = InStr(Str, Delimiter)
  If Y = 0 Then
    ReadNext = Str
    Str = ""
  Else
    ReadNext = Left$(Str, Y - 1)
    Str = Mid$(Str, Y + Len(Delimiter))
    If TrimStr Then Str = Trim$(Str)
  End If
End Function

' Reads the last item (separated by a delimiter) from
' a string and delete it from the string
Public Function ReadLast(Str As String, Delimiter As String) As String
  Dim Y As Long
  Y = InStrRev(Str, Delimiter)
  If Y > 0 Then
    ReadLast = Trim$(Mid$(Str, Y + 1))
    Str = Left$(Str, Y - 1)
  Else
    ReadLast = Str
    Str = ""
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

' Sub SelectText
' Selects the text of a text box
Public Sub SelectText(X As TextBox)
  X.SelStart = 0: X.SelLength = Len(X.Text)
End Sub

' Strip chr$(0)s from a string
Public Function StripTerminator(Tmp As String) As String
  StripTerminator = Tmp
  If InStr(Tmp, Chr$(0)) > 0 Then StripTerminator = Left$(Tmp, InStr(Tmp, Chr$(0)) - 1)
End Function
