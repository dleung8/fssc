VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' [Common Dialog]
' Replicates functionality of the CommonDialog control
' without the control

Option Explicit

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type OPENFILENAME
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  Flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

Private Type ChooseColor
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  rgbResult As Long
  lpCustColors As Long
  Flags As Long
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_LONGNAMES = &H200000                       '  force long names for 3.x modules
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_EXPLORER = &H80000                         '  new look commdlg
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_READONLY = &H1

Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
Private Const CC_ANYCOLOR = &H100
Private Const CC_RGBINIT = &H1
Private Const CC_FULLOPEN = &H2
Private Const CC_SHOWHELP = &H8
Private Const CC_ENABLEHOOK = &H10

Public FilterIndex As Integer
Public DefExt As String

Private CustomColor(15) As Long
Private CustomColorSet As Boolean
Private locFilter As String

' Function ColorDialog
' Shows Common Color Dialog
' Color = Sets/Returns the Color
' Returns True/False if User selected OK
Public Function ColorDialog(ByRef Color As Long, ByVal FullOpen As Boolean) As Boolean
  Dim I As Integer, X As ChooseColor
  
  If Not CustomColorSet Then
    For I = 0 To 15
      CustomColor(I) = vbWhite
    Next I
    CustomColorSet = True
  End If
  
  With X
    .lStructSize = Len(X)
    .hwndOwner = Screen.ActiveForm.hwnd
    .rgbResult = Color
    .lpCustColors = VarPtr(CustomColor(0))
    .Flags = CC_ANYCOLOR Or CC_RGBINIT Or (IIf(FullOpen, CC_FULLOPEN, 0))
    If ChooseColor(X) <> 0 Then
      ColorDialog = True
      Color = .rgbResult
    End If
  End With
End Function

' Changes the Filter
Public Property Let Filter(ByVal NewFilter As String)
  locFilter = Replace(NewFilter, "|", vbNullChar) & vbNullChar & vbNullChar
End Property

' Function OpenDialog
' Shows Common Open Dialog
' Title = Title of Dialog Box
' Returns file name
Public Function OpenDialog(Optional ByVal Title As String, Optional ByVal File As String) As String
  Dim X As OPENFILENAME, FileBuffer As String
  FileBuffer = File & String$(257 - Len(File), 0)
  
  With X
    .lStructSize = Len(X)
    .hwndOwner = Screen.ActiveForm.hwnd
    .hInstance = App.hInstance
    .lpstrFilter = locFilter
    .nFilterIndex = FilterIndex
    .lpstrFile = FileBuffer
    .nMaxFile = Len(FileBuffer) - 1
    .lpstrFileTitle = .lpstrFile
    .nMaxFileTitle = .nMaxFile
    .lpstrInitialDir = GetDir(File)
    .lpstrTitle = Title
    .lpstrDefExt = DefExt
    .Flags = OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST Or OFN_LONGNAMES Or OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_EXPLORER
  
    If GetOpenFileName(X) <> 0 Then
      OpenDialog = StripTerminator(.lpstrFile)
      FilterIndex = .nFilterIndex
    End If
  End With
End Function

' Function SaveDialog
' Shows Common Save Dialog
' File = File Name
' Title = Title of Dialog Box
' FilterIndex = Sets the FilterIndex of the dialog
' Returns file name
Public Function SaveDialog(ByVal File As String, Optional ByVal Title As String) As String
  Dim X As OPENFILENAME, FileBuffer As String
  FileBuffer = File & String$(257 - Len(File), 0)
  
  With X
    .lStructSize = Len(X)
    .hwndOwner = Screen.ActiveForm.hwnd
    .hInstance = App.hInstance
    .lpstrFilter = locFilter
    .nFilterIndex = FilterIndex
    .lpstrFile = FileBuffer
    .nMaxFile = Len(FileBuffer) - 1
    .lpstrFileTitle = .lpstrFile
    .nMaxFileTitle = .nMaxFile
    .lpstrInitialDir = GetDir(File)
    .lpstrTitle = Title
    .lpstrDefExt = DefExt
    .Flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST
    
    If GetSaveFileName(X) <> 0 Then
      SaveDialog = StripTerminator(.lpstrFile)
      FilterIndex = .nFilterIndex
    End If
  End With
End Function

' Sub SelectMultiDialog
' Shows the Open Dialog Box (Multiselect)
' Title = Title of the Dialog
' File = Array to hold file names
' Returns the names of file in File Array Base 1 or
' Array with UBound(0) for empty
Public Function SelectMultiDialog(File() As String, Optional ByVal Title As String) As Boolean
  Dim X As OPENFILENAME, FileBuffer As String, _
    I As Integer, Temp As String, TempDir As String
  ReDim File(0)
  FileBuffer = String$(32767, 0)
  
  With X
    .lStructSize = Len(X)
    .hwndOwner = Screen.ActiveForm.hwnd
    .hInstance = App.hInstance
    .lpstrFilter = locFilter
    .nFilterIndex = FilterIndex
    .lpstrFile = FileBuffer
    .nMaxFile = Len(FileBuffer) - 1
    .lpstrFileTitle = .lpstrFile
    .nMaxFileTitle = .nMaxFile
    .lpstrInitialDir = CurDir$
    .lpstrTitle = Title
    .lpstrDefExt = DefExt
    .Flags = OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST Or OFN_LONGNAMES Or OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_EXPLORER Or OFN_ALLOWMULTISELECT
  End With
  
  If GetOpenFileName(X) <> 0 Then
    Temp = X.lpstrFile
    FilterIndex = X.nFilterIndex
    
    TempDir = ReadNext(Temp, vbNullChar)
    
    If GetAttr(TempDir) And vbDirectory Then
      I = 0
      
      ReDim File(10)
      Do
        If UBound(File) < I Then ReDim Preserve File(I * 2)
        File(I) = AddDir(TempDir, ReadNext(Temp, vbNullChar))
        I = I + 1
      Loop Until StripTerminator(Temp) = ""
      ReDim Preserve File(I - 1)
    Else
      ReDim File(0)
      File(0) = TempDir
    End If
    SelectMultiDialog = True
  End If
End Function
