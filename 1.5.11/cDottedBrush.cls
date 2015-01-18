VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDottedBrush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Emulates a dotted brush

Option Explicit

Private Type Bitmap
  bmType As Long
  bmWidth As Long
  bmHeight As Long
  bmWidthBytes As Long
  bmPlanes As Integer
  bmBitsPixel As Integer
  bmBits As Long
End Type

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As Any) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateBitmapIndirect Lib "gdi32" (lpBitmap As Bitmap) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long

Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

' The brush:
Private m_hBrush As Long

Public Function Create() As Boolean
  Dim tbm As Bitmap, hBm As Long
  Dim m_lPattern(3) As Long
  Dim I As Integer

  ' The brush's pattern:
  For I = 0 To 3
    m_lPattern(I) = &HAAAA5555
  Next I

  Destroy
     
  ' Create a monochrome bitmap containing the desired pattern:
  With tbm
    .bmType = 0
    .bmWidth = 16
    .bmHeight = 8
    .bmWidthBytes = 2
    .bmPlanes = 1
    .bmBitsPixel = 1
    .bmBits = VarPtr(m_lPattern(0))
  End With
  hBm = CreateBitmapIndirect(tbm)

  ' Make a brush from the bitmap bits
  m_hBrush = CreatePatternBrush(hBm)

  ' Delete the bitmap
  DeleteObject hBm
End Function

Public Sub Destroy()
  If m_hBrush <> 0 Then
    DeleteObject m_hBrush
    m_hBrush = 0
  End If
End Sub

Public Sub Rectangle(ByVal hdc As Long, ByVal xPixels As Long, ByVal yPixels As Long, ByVal widthRectPixels As Long, ByVal heightRectPixels As Long, ByVal lineSizePixels As Long, Optional ByVal oBackColor As OLE_COLOR = -1, Optional ByVal oForeColor As OLE_COLOR = -1)
  Dim lBkColor As Long, lForeColor As Long
  Dim tR As RECT, hOldBrush As Long

  If oBackColor <> -1 Then lBkColor = SetBkColor(hdc, TranslateColor(oBackColor))
  If oForeColor <> -1 Then lForeColor = SetTextColor(hdc, TranslateColor(oForeColor))

  With tR
    .Left = xPixels
    .Top = xPixels
    .Right = .Left + widthRectPixels
    .Bottom = .Top + heightRectPixels
  End With
  FillRect hdc, tR, m_hBrush
End Sub
   
Public Sub SetBrushOrigin(ByVal hdc As Long, ByVal xPixels As Long, ByVal yPixels As Long)
  SetBrushOrgEx hdc, -xPixels And &H7&, -yPixels And &H7&, ByVal 0&
End Sub

Private Sub Class_Terminate()
  Destroy
End Sub
