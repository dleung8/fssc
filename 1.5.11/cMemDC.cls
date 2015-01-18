VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMemDC"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Memory DC for flicker free drawing.
Option Explicit

' ======================================================================================
' Name:     cMemDC.cls
' Author:   Steve McMahon (steve@vbaccelerator.com)
' Date:     20 October 1999
'
' Copyright © 1999 Steve McMahon for vbAccelerator
' --------------------------------------------------------------------------------------
' Visit vbAccelerator - advanced free source code for VB programmers
'    http://vbaccelerator.com
' --------------------------------------------------------------------------------------
'
' FREE SOURCE CODE - ENJOY!
' Do not sell this code.  Credit vbAccelerator.
' ======================================================================================

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal myHdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal myHdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal myHdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal myHdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private myHdc As Long
Private myhBitmap As Long, myhBitmapOld As Long
Private myWidth As Long, myHeight As Long

Public Sub CreateBitmap(ByVal W As Long, ByVal H As Long)
  DestroyBitmap
  myHdc = CreateCompatibleDC(0&)
  myhBitmap = CreateCompatibleBitmap(frmMain.hdc, W, H)
  myhBitmapOld = SelectObject(myHdc, myhBitmap)
  If myhBitmapOld = 0 Then
    DestroyBitmap
  Else
    myWidth = W
    myHeight = H
  End If
End Sub

Private Sub DestroyBitmap()
  If myhBitmapOld <> 0 Then
    SelectObject myHdc, myhBitmapOld
    myhBitmapOld = 0
  End If
  If myhBitmap <> 0 Then
    DeleteObject myhBitmap
    myhBitmap = 0
  End If
  If myHdc <> 0 Then
    DeleteDC myHdc
    myHdc = 0
  End If
  myWidth = 0
  myHeight = 0
End Sub

Public Property Get hdc() As Long
  hdc = myHdc
End Property

Public Property Get Height() As Long
  Height = myHeight
End Property

Public Property Let Height(ByVal H As Long)
  If H > myHeight Then CreateBitmap myWidth, H
End Property

Public Property Get Width() As Long
  Width = myWidth
End Property

Public Property Let Width(ByVal W As Long)
  If W > myWidth Then CreateBitmap W, myHeight
End Property

Private Sub Class_Terminate()
  DestroyBitmap
End Sub
