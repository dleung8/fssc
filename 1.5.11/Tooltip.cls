VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTooltip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Copyright © 1997-1999 Brad Martinez, http://www.mvps.org
'
' Defaults:
' DelayInitial = 500  (1/2 sec)
' DelayAutoPopup = 5000  (5 secs)
' DelayReshow = 100 (1/10 sec)
' MaxTipWidth = 0
' all Margins = 0

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Enum TT_Flags
  TTF_IDISHWND = &H1
  TTF_SUBCLASS = &H10
End Enum

Private Type TOOLINFO
  cbSize As Long
  uFlags As TT_Flags
  hWnd As Long
  uId As Long
  RECT As RECT
  hinst As Long
  lpszText As String   ' Long
  lParam As Long
End Type

Private Enum TT_DelayTime
  TTDT_AUTOMATIC = 0
  TTDT_RESHOW = 1
  TTDT_AUTOPOP = 2
  TTDT_INITIAL = 3
End Enum

Public Enum ttDelayTimeConstants
  ttDelayDefault = TTDT_AUTOMATIC '= 0
  ttDelayInitial = TTDT_INITIAL '= 3
  ttDelayShow = TTDT_AUTOPOP '= 2
  ttDelayReshow = TTDT_RESHOW '= 1
  ttDelayMask = 3
End Enum

' Exported by Comctl32.dll >= v4.00.950
' Ensures that the common control dynamic-link library (DLL) is loaded.
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Const TOOLTIPS_CLASS = "tooltips_class32"

' Styles
Private Const TTS_ALWAYSTIP = &H1

' Messages
Private Const WM_USER = &H400
Private Const TTM_SETDELAYTIME = (WM_USER + 3)
Private Const TTM_GETTOOLCOUNT = (WM_USER + 13)
Private Const TTM_ADDTOOL = (WM_USER + 4)
Private Const TTM_UPDATETIPTEXT = (WM_USER + 12)
Private Const TTM_ENUMTOOLS = (WM_USER + 14)
Private Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)

Private m_hwndTT As Long

Public Function AddTool(Ctrl As Control, Optional ByVal sText As String) As Boolean
  Dim ti As TOOLINFO
  
  If (m_hwndTT = 0) Then Exit Function
  
  If Not GetToolInfo(Ctrl.hWnd, ti) Then
    With ti
      .cbSize = Len(ti)
      ' TTF_IDISHWND must be specified to tell the tooltip control
      ' to retrieve the control's rect from it's hWnd specified in uId.
      .uFlags = TTF_SUBCLASS Or TTF_IDISHWND
      .hWnd = Ctrl.Container.hWnd
      .uId = Ctrl.hWnd
      
      .lpszText = sText
    End With
    
    ' Returns 1 on success, 0 on failure
    AddTool = SendMessage(m_hwndTT, TTM_ADDTOOL, 0, ti)
  End If
End Function

Public Function Create(frm As Form) As Boolean
  If (m_hwndTT = 0) Then
    InitCommonControls
    
    ' Filling the hwndParent param below allows the tooltip window to
    ' be owned by the specified form and be destroyed along with it,
    ' but we'll cleanup in Class_Terminate anyway.
    ' No WS_EX_TOPMOST or TTS_ALWAYSTIP per Win95 UI rules...
    m_hwndTT = CreateWindowEx(0, TOOLTIPS_CLASS, vbNullString, TTS_ALWAYSTIP, 0, 0, 0, 0, frm.hWnd, 0, App.hInstance, ByVal 0)
  End If
  Create = m_hwndTT <> 0
End Function

Public Property Let DelayTime(ByVal dwType As ttDelayTimeConstants, ByVal dwMilliSecs As Long)
  If m_hwndTT = 0 Then Exit Property
  SendMessage m_hwndTT, TTM_SETDELAYTIME, (dwType And ttDelayMask), ByVal dwMilliSecs
End Property

Public Property Let MaxTipWidth(ByVal cX As Long)
  If m_hwndTT = 0 Then Exit Property
  If (cX < 1) Then cX = -1
  SendMessage m_hwndTT, TTM_SETMAXTIPWIDTH, 0, ByVal cX
End Property

Public Property Get ToolCount() As Integer ' read-only
  If m_hwndTT = 0 Then Exit Property
  ToolCount = SendMessage(m_hwndTT, TTM_GETTOOLCOUNT, 0, 0)
End Property

Public Property Let ToolText(Ctrl As Control, ByVal sText As String)
  Dim ti As TOOLINFO
  
  If (m_hwndTT = 0) Then Exit Property
  
  If GetToolInfo(Ctrl.hWnd, ti) Then
    ti.lpszText = sText
    ' The tooltip won't appear for the control if lpszText is an empty string
    SendMessage m_hwndTT, TTM_UPDATETIPTEXT, 0, ti
  End If
End Property

Private Function GetToolInfo(ByVal hwndTool As Long, ti As TOOLINFO) As Boolean
  Dim nItems As Integer, I As Integer
  ti.cbSize = Len(ti)
  nItems = ToolCount
  For I = 0 To nItems - 1
    ' call returns 1 on success, 0 on failure...
    If SendMessage(m_hwndTT, TTM_ENUMTOOLS, I, ti) Then
      If hwndTool = ti.uId Then
        GetToolInfo = True
        Exit Function
      End If
    End If
  Next I
End Function

Private Sub Class_Terminate()
  If m_hwndTT Then DestroyWindow m_hwndTT
End Sub
