VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cNCMetrics"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Provides an interface for getting Non-client metrics
Option Explicit

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Const LOGPIXELSY = 90    '  Logical pixels/inch in Y

Private Const SPI_GETICONTITLELOGFONT = 31
Private Const LF_FACESIZE = 32

' Normal log font structure:
Private Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName(LF_FACESIZE) As Byte
End Type

Private Enum CNCMetricsFontWeightConstants
  FW_NORMAL = 400
  FW_REGULAR = 400
  FW_BOLD = 700
  FW_BLACK = 900
End Enum

Public Enum CNCMetricsFontTypes
  IconFont = 1
  CaptionFont = 2
  SMCaptionFont = 3
  MenuFont = 4
  StatusFont = 5
  MessageFont = 6
End Enum

Private Type NMLOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName(LF_FACESIZE - 4) As Byte
End Type

Private Type NONCLIENTMETRICS
  cbSize As Long
  iBorderWidth As Long
  iScrollWidth As Long
  iScrollHeight As Long
  iCaptionWidth As Long
  iCaptionHeight As Long
  lfCaptionFont As NMLOGFONT
  iSMCaptionWidth As Long
  iSMCaptionHeight As Long
  lfSMCaptionFont As NMLOGFONT
  iMenuWidth As Long
  iMenuHeight As Long
  lfMenuFont As NMLOGFONT
  lfStatusFont As NMLOGFONT
  lfMessageFont As NMLOGFONT
End Type

Private Const SPI_GETNONCLIENTMETRICS = 41
Private Const SPI_SETNONCLIENTMETRICS = 42

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private m_tNCM As NONCLIENTMETRICS
Private m_tLF As LOGFONT
Private m_hFont(1 To 6) As Long

Public Property Get BorderWidth() As Long
  BorderWidth = m_tNCM.iBorderWidth
End Property

Public Property Get CaptionHeight() As Long
  CaptionHeight = m_tNCM.iCaptionHeight
End Property

Public Property Get CaptionWidth() As Long
  CaptionWidth = m_tNCM.iCaptionWidth
End Property

Public Function ClearUp()
  Dim I As Integer
  For I = 1 To 6
    If (m_hFont(I) <> 0) Then DeleteObject m_hFont(I)
    m_hFont(I) = 0
  Next I
End Function

Public Property Get Font(ByVal lHDC As Long, ByVal eFontNum As CNCMetricsFontTypes) As IFont
  Dim sFnt As New StdFont
  pLogFontToStdFont m_tNCM.lfMenuFont, lHDC, sFnt
  Set Font = sFnt
End Property

Public Property Get FontHandle(ByVal eFontNum As CNCMetricsFontTypes) As Long
  Dim lR As Long, tLF As LOGFONT
    
  Select Case eFontNum
    Case StatusFont
      CopyMemory tLF, m_tNCM.lfStatusFont, LenB(m_tNCM.lfStatusFont)
    Case SMCaptionFont
      CopyMemory tLF, m_tNCM.lfSMCaptionFont, LenB(m_tNCM.lfSMCaptionFont)
    Case MessageFont
      CopyMemory tLF, m_tNCM.lfMessageFont, LenB(m_tNCM.lfMessageFont)
    Case MenuFont
      CopyMemory tLF, m_tNCM.lfMenuFont, LenB(m_tNCM.lfMenuFont)
    Case IconFont
      CopyMemory tLF, m_tLF, LenB(m_tLF)
    Case CaptionFont
      CopyMemory tLF, m_tNCM.lfCaptionFont, LenB(m_tNCM.lfCaptionFont)
  End Select
    
  FontHandle = plGetFont(eFontNum, tLF)
End Property

Public Function GetMetrics() As Boolean
  Dim lR As Long

  ClearUp
  ' Get Non-client metrics:
  m_tNCM.cbSize = 340 'LenB(m_tNCM) - why doesn't this go?
  lR = SystemParametersInfo(SPI_GETNONCLIENTMETRICS, 0, m_tNCM, 0)
  If (lR <> 0) Then
    ' Get icon font:
    lR = SystemParametersInfo(SPI_GETICONTITLELOGFONT, 0, m_tLF, 0)
    GetMetrics = (lR <> 0)
  End If
End Function

Public Property Get MenuHeight() As Long
  MenuHeight = m_tNCM.iMenuHeight
End Property

Public Property Get MenuWidth() As Long
  MenuWidth = m_tNCM.iMenuWidth
End Property

Private Function plGetFont(ByVal eFontNum As CNCMetricsFontTypes, ByRef tLF As LOGFONT)
  If (m_hFont(eFontNum) = 0) Then m_hFont(eFontNum) = CreateFontIndirect(tLF)
  plGetFont = m_hFont(eFontNum)
End Function

Private Sub pLogFontToStdFont(ByRef tLF As NMLOGFONT, ByVal hdc As Long, ByRef sFnt As StdFont)
  With sFnt
    .Name = StrConv(tLF.lfFaceName, vbUnicode)
    If tLF.lfHeight < 1 Then
      .Size = Abs((72# / GetDeviceCaps(hdc, LOGPIXELSY)) * tLF.lfHeight)
    Else
      .Size = tLF.lfHeight
    End If
    .Charset = tLF.lfCharSet
    .Italic = Not (tLF.lfItalic = 0)
    .Underline = Not (tLF.lfUnderline = 0)
    .Strikethrough = Not (tLF.lfStrikeOut = 0)
    .Bold = (tLF.lfWeight > FW_REGULAR)
  End With
End Sub

Public Property Get ScrollHeight() As Long
  ScrollHeight = m_tNCM.iScrollHeight
End Property

Public Property Get ScrollWidth() As Long
  ScrollWidth = m_tNCM.iScrollWidth
End Property

Public Property Get SMCaptionHeight() As Long
  SMCaptionHeight = m_tNCM.iSMCaptionHeight
End Property

Public Property Get SMCaptionWidth() As Long
  SMCaptionWidth = m_tNCM.iSMCaptionWidth
End Property

Private Sub Class_Terminate()
  ClearUp
End Sub
