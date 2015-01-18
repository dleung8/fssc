Attribute VB_Name = "basCDialog"
' [Common Dialog]
' Replicates functionality of the CommonDialog control
' without the control
Option Explicit
Option Compare Text

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

Public FilterIndex As Integer
Public DefExt As String
Private locFilter As String

' Changes the Filter
Public Property Let Filter(ByVal Tmp As String)
  Do Until InStr(Tmp, "|") = 0
    Mid$(Tmp, InStr(Tmp, "|"), 1) = Chr$(0)
  Loop
  locFilter = Tmp & Chr$(0)
End Property

' Function OpenDialog
' Shows Common Open Dialog
' Title = Title of Dialog Box
' Returns file name
Public Function OpenDialog(Optional Title As String, Optional File As String) As String
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
  End With
  
  If GetOpenFileName(X) <> 0 Then
    OpenDialog = StripTerminator(X.lpstrFile)
    FilterIndex = X.nFilterIndex
  End If
End Function
