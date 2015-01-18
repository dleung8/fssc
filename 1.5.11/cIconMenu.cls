VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIconMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class cIconMenu
' Provides functions for maintaining owner drawn menus
' (with icons)
Option Explicit
Implements ISubclass

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

Private Type MEASUREITEMSTRUCT
  CtlType As Long
  CtlID As Long
  ItemID As Long
  ItemWidth As Long
  ItemHeight As Long
  ItemData As Long
End Type

Private Type DRAWITEMSTRUCT
  CtlType As Long
  CtlID As Long
  ItemID As Long
  itemAction As Long
  ItemState As Long
  hwndItem As Long
  hdc As Long
  rcItem As RECT
  ItemData As Long
End Type

Private Type MENUITEMINFO
  cbSize As Long
  fMask As Long
  fType As Long
  fState As Long
  wID As Long
  hSubMenu As Long
  hbmpChecked As Long
  hbmpUnchecked As Long
  dwItemData As Long
  dwTypeData As Long
  cch As Long
End Type

Private Type MENUITEMINFO_STRINGDATA
  cbSize As Long
  fMask As Long
  fType As Long
  fState As Long
  wID As Long
  hSubMenu As Long
  hbmpChecked As Long
  hbmpUnchecked As Long
  dwItemData As Long
  dwTypeData As String
  cch As Long
End Type

' Application-defined types
Private Type MenuRestoreListType
  hMenu As Long
  iMenu As Long
  tMIIS As MENUITEMINFO_STRINGDATA
  sCaption As String
  sShortcut As String
  sAccelerator As String
  lMenuTextSize As Long
  lMenuShortcutSize As Long
End Type

' Menu flag constants
Private Const MF_CHECKED = &H8&
Private Const MF_GRAYED = &H1&
Private Const MF_HILITE = &H80&
Private Const MF_OWNERDRAW = &H100&
Private Const MF_SEPARATOR = &H800&
Private Const MF_STRING = &H0&
Private Const MF_DISABLED = &H3&
Private Const MF_RADIOCHECK = &H200&

' MenuItemInfo Mask constants
Private Const MIIM_STATE = &H1&
Private Const MIIM_ID = &H2&
Private Const MIIM_TYPE = &H10&
Private Const MIIM_DATA = &H20&

' Owner draw information:
Private Const ODT_MENU = 1

Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function GetMenuItemInfoStr Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Boolean, lpMenuItemInfo As MENUITEMINFO_STRINGDATA) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SetMenuItemInfoStr Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO_STRINGDATA) As Long

Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal fuFlags As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long

' Image type
Private Const DST_ICON = &H3

' State type
Private Const DSS_DISABLED = &H20

Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_CALCRECT = &H400
Private Const DT_VCENTER = &H4
Private Const DT_SINGLELINE = &H20

Private Const TRANSPARENT = 1
Private Const OPAQUE = 2

Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENOUTER = &H2
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_SUNKENINNER = &H8
Private Const BF_LEFT = &H1
Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_BOTTOM = &H8
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal hIml As Long, ByVal I As Long, ByVal hdcDst As Long, ByVal X As Long, ByVal Y As Long, ByVal fStyle As Long) As Long
Private Declare Function ImageList_GetIcon Lib "comctl32.dll" (ByVal hIml As Long, ByVal I As Long, ByVal diIgnore As Long) As Long
Private Declare Function ImageList_GetIconSize Lib "comctl32.dll" (ByVal hIml As Long, cX As Long, cY As Long) As Long

Private Const ILD_TRANSPARENT = 1
Private Const ILD_BLEND25 = 2
Private Const ILD_SELECTED = 4

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

' Window Messages
Private Const WM_DESTROY = &H2
Private Const WM_WININICHANGE = &H1A
Private Const WM_MEASUREITEM = &H2C
Private Const WM_DRAWITEM = &H2B
Private Const WM_INITMENUPOPUP = &H117
Private Const WM_EXITMENULOOP = &H212

Private m_hWnd As Long
Private m_cMemDC As cMemDC
Private m_cNCM As cNCMetrics
Private m_cBrush As cDottedBrush
Private m_fnt As StdFont
Private m_fntSymbol As StdFont

Private m_iRestore As Long
Private RestoreList() As MenuRestoreListType
Private IconImages(1 To 200) As Integer
Private m_iHaveSeenCount As Integer
Private m_hMenuSeen() As Long
Private m_lMenuItemHeight As Integer

Private m_hIml As Long

Private Const m_lIconSize = 16

Public Sub Attach(ByVal lhWnd As Long)
  Detach
  m_hWnd = lhWnd
  AttachMessage Me, m_hWnd, WM_WININICHANGE
  AttachMessage Me, m_hWnd, WM_DRAWITEM
  AttachMessage Me, m_hWnd, WM_MEASUREITEM
  AttachMessage Me, m_hWnd, WM_INITMENUPOPUP
  AttachMessage Me, m_hWnd, WM_EXITMENULOOP
  AttachMessage Me, m_hWnd, WM_DESTROY
End Sub

Private Function DrawItem(ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim tDIS As DRAWITEMSTRUCT, hBr As Long, _
    tR As RECT, tTR As RECT, tWR As RECT, _
    hFntOld As Long, hFntsOld As Long, _
    bRadioCheck As Boolean, bDisabled As Boolean, _
    bChecked As Boolean, bHighlighted As Boolean, _
    tMII As MENUITEMINFO, _
    lHDC As Long, lID As Long, lSelLeft As Long, _
    lIconIndex As Long, TextColor As Long

  CopyMemory tDIS, ByVal lParam, LenB(tDIS)
  If tDIS.CtlType = ODT_MENU Then
    ' tDIS.hWndItem is the menu containing the item, tDIS.itemID is the wID
    
    With m_cMemDC
      .Width = tDIS.rcItem.Right - tDIS.rcItem.Left + 1
      .Height = tDIS.rcItem.Bottom - tDIS.rcItem.Top + 1
      lHDC = .hdc
      hFntOld = SelectObject(lHDC, hFont)
      
      LSet tR = tDIS.rcItem
      OffsetRect tR, -tR.Left, -tR.Top
            
      ' Fill background:
      tTR.Right = .Width
      tTR.Bottom = .Height
    End With
    
    hBr = CreateSolidBrush(TranslateColor(vbMenuBar))
    FillRect lHDC, tTR, hBr
    DeleteObject hBr
    tR.Top = tR.Top + 1
    
    SetBkMode lHDC, TRANSPARENT
    
    ' Draw the text:
    With tMII
      .cbSize = LenB(tMII)
      .fMask = MIIM_TYPE Or MIIM_STATE Or MIIM_DATA
      ReDim B(0 To 128) As Byte
      .dwTypeData = VarPtr(B(0))
    End With
    GetMenuItemInfo tDIS.hwndItem, tDIS.ItemID, False, tMII
    
    If (tMII.fType And MF_SEPARATOR) = MF_SEPARATOR Then
      ' Separator
      LSet tWR = tR
      tWR.Top = (tWR.Bottom - tWR.Top - 2) \ 2 + tWR.Top
      tWR.Bottom = tWR.Top + 2
      InflateRect tWR, -12, 0
      DrawEdge lHDC, tWR, BDR_SUNKENOUTER, BF_TOP Or BF_BOTTOM
    Else
      ' Text item
      With tMII
        bRadioCheck = ((.fType And MF_RADIOCHECK) = MF_RADIOCHECK)
        bDisabled = ((.fState And MF_DISABLED) = MF_DISABLED)
        bChecked = ((.fState And MF_CHECKED) = MF_CHECKED)
        bHighlighted = ((.fState And MF_HILITE) = MF_HILITE)
      End With
       
      If bDisabled Then
        SetTextColor lHDC, TranslateColor(vb3DHighlight)
      Else
        If bHighlighted Then
          SetTextColor lHDC, TranslateColor(vbHighlightText)
        Else
          SetTextColor lHDC, TranslateColor(vbMenuText)
        End If
      End If
       
      lID = tMII.dwItemData

      ' Icon
      On Error Resume Next
      lIconIndex = -1
      lIconIndex = IconImages(RestoreList(lID).tMIIS.wID) - 1
      On Error GoTo 0
       
      If bChecked Or lIconIndex > -1 Then
        lSelLeft = m_lIconSize + 4
      End If
                         
      If bHighlighted And Not bDisabled Then
        LSet tWR = tR
        tWR.Left = lSelLeft
        hBr = CreateSolidBrush(TranslateColor(vbHighlight))
        FillRect lHDC, tWR, hBr
        DeleteObject hBr
      End If
       
      If bDisabled Then
        TextColor = TranslateColor(vb3DHighlight)
      Else
        If bHighlighted Then
          TextColor = TranslateColor(vbHighlightText)
        Else
          TextColor = TranslateColor(vbMenuText)
        End If
      End If
      SetTextColor lHDC, TextColor
       
      ' Get the check/icon space:
      LSet tWR = tR
      tWR.Right = tWR.Left + 19
       
      ' Check:
      If bChecked Then
        ' Color in:
        If Not bHighlighted Then
          SetBkMode lHDC, OPAQUE
          m_cBrush.Rectangle lHDC, tWR.Left + 1, tWR.Top + 1, tWR.Right - tWR.Left - 1, tWR.Bottom - tWR.Top - 1, 1, vbMenuBar, vb3DHighlight
          SetBkMode lHDC, TRANSPARENT
              
          SetTextColor lHDC, TextColor
        End If
        
        DrawEdge lHDC, tWR, BDR_SUNKENOUTER, BF_RECT
                    
        If lIconIndex = -1 Then
          ' Draw the appropriate symbol:
          SelectObject lHDC, hFntOld
          hFntsOld = SelectObject(lHDC, hFontSymbol)
          If bHighlighted Then
            OffsetRect tWR, 1, 1
            pDrawText lHDC, IIf(bRadioCheck, "h", "b"), tWR, DT_VCENTER Or DT_CENTER Or DT_SINGLELINE, bDisabled
            OffsetRect tWR, -1, 0
            pDrawText lHDC, IIf(bRadioCheck, "h", "b"), tWR, DT_VCENTER Or DT_CENTER Or DT_SINGLELINE, bDisabled
            OffsetRect tWR, 0, -1
            SetTextColor lHDC, TranslateColor(vbMenuText)
          End If
          pDrawText lHDC, IIf(bRadioCheck, "h", "b"), tWR, DT_VCENTER Or DT_CENTER Or DT_SINGLELINE, bDisabled
          
          SetTextColor lHDC, TextColor
           
          SelectObject lHDC, hFntsOld
          hFntOld = SelectObject(lHDC, hFont)
        Else
          If bDisabled Then
            ImageListDrawIconDisabled lHDC, m_hIml, lIconIndex, tWR.Left + 1, tWR.Top + 2, m_lIconSize
          Else
            ImageListDrawIcon lHDC, m_hIml, lIconIndex, tWR.Left + 1, tWR.Top + 2
          End If
        End If
      Else
        If lIconIndex > -1 Then
          If bHighlighted And Not bDisabled Then
            DrawEdge lHDC, tWR, BDR_RAISEDINNER, BF_RECT
          End If
          If bDisabled Then
            ImageListDrawIconDisabled lHDC, m_hIml, lIconIndex, tWR.Left + 1, tWR.Top + 2, m_lIconSize
          Else
            ImageListDrawIcon lHDC, m_hIml, lIconIndex, tWR.Left + 1, tWR.Top + 2
          End If
        End If
      End If
       
      ' Draw text
      LSet tWR = tR
      tWR.Left = m_lIconSize + 5
      If 0 < lID And lID <= m_iRestore Then
        pDrawText lHDC, RestoreList(lID).sCaption, tWR, DT_LEFT Or DT_SINGLELINE Or DT_VCENTER, bDisabled
        If Len(RestoreList(lID).sShortcut) > 0 Then
          tWR.Left = tWR.Right - m_lIconSize - 8 - RestoreList(lID).lMenuShortcutSize
          pDrawText lHDC, RestoreList(lID).sShortcut, tWR, DT_LEFT Or DT_SINGLELINE Or DT_VCENTER, bDisabled
        End If
      End If
    End If
    
    SelectObject lHDC, hFntOld
    
    BitBlt tDIS.hdc, tDIS.rcItem.Left, tDIS.rcItem.Top, tDIS.rcItem.Right - tDIS.rcItem.Left + 1, tDIS.rcItem.Bottom - tDIS.rcItem.Top + 1, lHDC, 0, 0, vbSrcCopy
  Else
    DrawItem = CallOldWindowProc(m_hWnd, WM_DRAWITEM, wParam, lParam)
  End If
End Function

Public Sub Detach()
  If m_hWnd <> 0 Then
    DetachMessage Me, m_hWnd, WM_WININICHANGE
    DetachMessage Me, m_hWnd, WM_DRAWITEM
    DetachMessage Me, m_hWnd, WM_MEASUREITEM
    DetachMessage Me, m_hWnd, WM_INITMENUPOPUP
    DetachMessage Me, m_hWnd, WM_EXITMENULOOP
    DetachMessage Me, m_hWnd, WM_DESTROY
  End If
End Sub

Public Property Get Font() As StdFont
  Dim lHDC As Long
  If m_fnt Is Nothing Then
    lHDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
    Set Font = m_cNCM.Font(lHDC, MenuFont)
    DeleteDC lHDC
  Else
    Set Font = m_fnt
  End If
End Property

Private Property Get hFont() As Long
  Dim iFn As IFont
  Set iFn = Font
  hFont = iFn.hFont
End Property

Private Property Get hFontSymbol() As Long
  Dim iFn As IFont
  Set iFn = m_fntSymbol
  hFontSymbol = iFn.hFont
End Property

Public Property Let ImageList(ByVal vImageList As Variant)
  If (VarType(vImageList) = vbObject) Then
    ' Assume a VB image list:
    On Error Resume Next
    ' Get the image list initialized.
    vImageList.ListImages(1).Draw 0, 0, 0, 1
    m_hIml = vImageList.hImageList
    On Error GoTo 0
  Else
    m_hIml = vImageList
  End If
End Property

Private Sub ImageListDrawIcon(ByVal hdc As Long, ByVal hIml As Long, ByVal iIconIndex As Long, ByVal lX As Long, ByVal lY As Long, Optional ByVal bSelected As Boolean = False, Optional ByVal bBlend25 As Boolean = False)
  Dim lFlags As Long

  lFlags = ILD_TRANSPARENT
  If bSelected Then lFlags = lFlags Or ILD_SELECTED
  If bBlend25 Then lFlags = lFlags Or ILD_BLEND25
  ImageList_Draw hIml, iIconIndex, hdc, lX, lY, lFlags
End Sub

Private Sub ImageListDrawIconDisabled(ByVal hdc As Long, ByVal hIml As Long, ByVal iIconIndex As Long, ByVal lX As Long, ByVal lY As Long, ByVal lSize As Long)
  Dim hIcon As Long

  hIcon = ImageList_GetIcon(hIml, iIconIndex, 0)
  DrawState hdc, 0, 0, hIcon, 0, lX, lY, lSize, lSize, DST_ICON Or DSS_DISABLED
  DestroyIcon hIcon
End Sub

Private Function InternalIDForWindowsID(ByVal wID As Long) As Integer
  Dim I As Integer
   ' linear search I'm afraid, but it is only called once
   ' per menu item shown (when WM_MEASUREITEM is fired)
  For I = 1 To m_iRestore
    If RestoreList(I).tMIIS.wID = wID Then
      InternalIDForWindowsID = I
      Exit Function
    End If
  Next I
End Function

Private Function MeasureItem(ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim tMIS As MEASUREITEMSTRUCT, lID As Long
   
  CopyMemory tMIS, ByVal lParam, LenB(tMIS)
  If tMIS.CtlType = ODT_MENU Then
  
    ' because we don't get the popup menu handle
    ' in the tMIS structure, we have to do an internal
    ' lookup to find info about this menu item.
    ' poor implementation of MEASUREITEMSTRUCT - it
    ' should have a .hWndItem field like DRAWITEMSTRUCT
    ' - spm

    lID = InternalIDForWindowsID(tMIS.ItemID)
    
    If lID > 0 And lID <= m_iRestore Then
      ' Width:
      tMIS.ItemWidth = RestoreList(lID).lMenuTextSize + RestoreList(lID).lMenuShortcutSize + 32 ' (icon + padding on right for symmetry)
      
      ' Height:
      If (RestoreList(lID).tMIIS.fType And MF_SEPARATOR) = MF_SEPARATOR Then
        tMIS.ItemHeight = 8
      Else
        ' menu item height is always the same
        tMIS.ItemHeight = m_lMenuItemHeight + 4
      End If
    Else
      MsgBox "MeasureItem error", vbCritical
    End If
    
    CopyMemory ByVal lParam, tMIS, LenB(tMIS)
  Else
    MeasureItem = CallOldWindowProc(m_hWnd, WM_MEASUREITEM, wParam, lParam)
  End If
End Function

Private Sub OwnerDrawMenu(ByVal hMenu As Long)
  Dim tMIIS As MENUITEMINFO_STRINGDATA, _
    tMII As MENUITEMINFO, _
    lC As Integer, I As Integer, _
    sCap As String, sShortcut As String, _
    tR As RECT, iPos As Long, lID As Long, _
    bHaveSeen As Boolean, hFntOld As Long, _
    lMenuTextSize As Long, lMenuShortcutSize As Long
  
  ' Set OD flag on the fly...
  bHaveSeen = pbHaveSeen(hMenu)

  hFntOld = SelectObject(m_cMemDC.hdc, hFont)
  lC = GetMenuItemCount(hMenu)
  For I = 0 To lC - 1
    If Not bHaveSeen Then
      With tMIIS
        .fMask = MIIM_TYPE Or MIIM_DATA Or MIIM_ID
        .cch = 127
        .dwTypeData = String$(128, 0)
        .cbSize = LenB(tMIIS)
      End With
      
      GetMenuItemInfoStr hMenu, I, True, tMIIS

      lID = plAddToRestoreList(hMenu, I, tMIIS)

      If Not (tMIIS.fType And MF_OWNERDRAW) = MF_OWNERDRAW Then
        ' Setting this flag causes tMIIS.dwTypeData to be
        ' overwritten with our own app-defined value:
        With tMII
          .fType = tMIIS.fType Or MF_OWNERDRAW And Not MF_STRING
          .dwItemData = lID
          .cbSize = LenB(tMII)
          .fMask = MIIM_TYPE Or MIIM_DATA
        End With
        SetMenuItemInfo hMenu, I, True, tMII
      End If
    Else
      ReDim B(0 To 128) As Byte
      With tMII
        .fMask = MIIM_TYPE Or MIIM_DATA Or MIIM_ID
        .cbSize = Len(tMIIS)
        .dwTypeData = VarPtr(B(0))
        GetMenuItemInfo hMenu, I, True, tMII
        lID = .dwItemData
      End With
  
      If Not ((tMII.fType And MF_OWNERDRAW) = MF_OWNERDRAW) Then
        lID = plReplaceIndex(hMenu, I)

        With tMIIS
          .fMask = MIIM_TYPE Or MIIM_DATA Or MIIM_ID
          .cch = 127
          .dwTypeData = String$(128, 0)
          .cbSize = LenB(tMIIS)
        End With
        GetMenuItemInfoStr hMenu, I, True, tMIIS

        If lID = 0 Then
          ' New item,just added:
          lID = plAddToRestoreList(hMenu, I, tMIIS)
        Else
          ' replacing existing:
          pReplaceRestoreList lID, hMenu, I, tMIIS
        End If

        ' Setting this flag causes tMIIS.dwTypeData to be
        ' overwritten with our own app-defined value:
        With tMII
          .fType = tMIIS.fType Or MF_OWNERDRAW And Not MF_STRING
          .dwItemData = lID
          .cbSize = LenB(tMII)
          .fMask = MIIM_TYPE Or MIIM_DATA
        End With
        SetMenuItemInfo hMenu, I, True, tMII

        ' Setting this flag causes tMIIS.dwTypeData to be
        ' overwritten with our own app-defined value:
        With tMII
          .fType = tMIIS.fType Or MF_OWNERDRAW And Not MF_STRING
          .dwItemData = lID
          .cbSize = LenB(tMII)
          .fMask = MIIM_TYPE Or MIIM_DATA
        End With
        SetMenuItemInfo hMenu, I, True, tMII
      End If
    End If

    If 0 < lID And lID <= m_iRestore Then
      sCap = RestoreList(lID).sCaption
      sShortcut = RestoreList(lID).sShortcut

      DrawText m_cMemDC.hdc, sCap, -1, tR, DT_LEFT Or DT_SINGLELINE Or DT_CALCRECT
      If tR.Right - tR.Left + 1 > lMenuTextSize Then lMenuTextSize = tR.Right - tR.Left + 1
      If Len(sShortcut) > 0 Then
        DrawText m_cMemDC.hdc, sShortcut, -1, tR, DT_LEFT Or DT_SINGLELINE Or DT_CALCRECT
        If tR.Right - tR.Left + 1 > lMenuShortcutSize Then lMenuShortcutSize = tR.Right - tR.Left + 1
      End If
      m_lMenuItemHeight = tR.Bottom - tR.Top + 2
      If m_lMenuItemHeight < m_lIconSize Then
        m_lMenuItemHeight = m_lIconSize
      End If
    Else
       MsgBox "DrawMenu Error", vbCritical
    End If
  Next I
   
  For I = 1 To m_iRestore
    With RestoreList(I)
      If .hMenu = hMenu Then
        .lMenuTextSize = lMenuTextSize
        .lMenuShortcutSize = lMenuShortcutSize
      End If
    End With
  Next I
  
  SelectObject m_cMemDC.hdc, hFntOld
End Sub

Private Function pbHaveSeen(ByVal hMenu As Long) As Boolean
  ' When WM_INITMENUPOPUP fires, this may or not be
  ' a new menu.  We use an array to store which menus
  ' we've already worked on:

  Dim I As Integer
   
  For I = 1 To m_iHaveSeenCount
    If hMenu = m_hMenuSeen(I) Then
      pbHaveSeen = True
      Exit Function
    End If
  Next I
  m_iHaveSeenCount = m_iHaveSeenCount + 1
  ReDim Preserve m_hMenuSeen(1 To m_iHaveSeenCount)
  m_hMenuSeen(m_iHaveSeenCount) = hMenu
End Function

Private Function plAddToRestoreList(ByVal hMenu As Long, ByVal iMenu As Integer, tMIIS As MENUITEMINFO_STRINGDATA) As Long
  ' Here we store information about a menu item.  When the
  ' menus are closed again we can reset things back to the
  ' way they were using this struct.

  m_iRestore = m_iRestore + 1
  If m_iRestore > UBound(RestoreList) Then ReDim Preserve RestoreList(m_iRestore * 2)
  pReplaceRestoreList m_iRestore, hMenu, iMenu, tMIIS
  plAddToRestoreList = m_iRestore
End Function

Private Function pDrawText(ByVal lHDC As Long, ByVal sText As String, tR As RECT, ByVal dtFlags As Long, ByVal bDisabled As Boolean)
  Dim tWR As RECT
  LSet tWR = tR
  If bDisabled Then
    SetTextColor lHDC, TranslateColor(vb3DHighlight)
    OffsetRect tWR, 1, 1
  End If
  DrawText lHDC, sText, -1, tWR, dtFlags
  If bDisabled Then
    SetTextColor lHDC, TranslateColor(vbButtonShadow)
    OffsetRect tWR, -1, -1
    DrawText lHDC, sText, -1, tWR, dtFlags
  End If
End Function

Private Function plReplaceIndex(ByVal hMenu As Long, ByVal iMenu As Integer)
  Dim I As Integer
  For I = 1 To m_iRestore
    If RestoreList(I).hMenu = hMenu Then
      If RestoreList(I).iMenu = iMenu Then
        plReplaceIndex = I
        Exit Function
      End If
    End If
  Next I
End Function

Private Sub pRestoreList()
  Dim I As Integer
  ' erase the lot:
  For I = 1 To m_iRestore
    SetMenuItemInfoStr RestoreList(I).hMenu, RestoreList(I).iMenu, True, RestoreList(I).tMIIS
  Next I
  m_iRestore = 0
  ReDim RestoreList(10)
  m_iHaveSeenCount = 0
  Erase m_hMenuSeen()
End Sub

Private Sub pReplaceRestoreList(ByVal lIdx As Integer, ByVal hMenu As Long, ByVal iMenu As Integer, tMIIS As MENUITEMINFO_STRINGDATA)
  Dim sCap As String, sShortcut As String, iPos As Long

  With RestoreList(lIdx)
    .hMenu = hMenu
    .iMenu = iMenu
    LSet .tMIIS = tMIIS
    
    sCap = Left$(tMIIS.dwTypeData, tMIIS.cch)
    
    iPos = InStr(sCap, vbTab)
    If iPos > 0 Then
      .sCaption = Left$(sCap, iPos - 1)
      .sShortcut = Mid$(sCap, iPos + 1)
    Else
      .sCaption = sCap
      .sShortcut = ""
    End If
    
    iPos = InStr(.sCaption, "&")
    
    If iPos > 0 And iPos < Len(.sCaption) Then
      .sAccelerator = UCase$(Mid$(.sCaption, iPos + 1, 1))
    End If
  End With
End Sub

Private Sub pSetFont(fntThis As StdFont)
  Set m_fnt = fntThis
  m_fntSymbol.Size = Font.Size * 1.2
End Sub

Public Sub SetIcon(Index As Integer, ListImage As Integer)
  IconImages(Index) = ListImage
End Sub

Private Sub Class_Initialize()
  Set m_cNCM = New cNCMetrics
  m_cNCM.GetMetrics
  
  Set m_cMemDC = New cMemDC
  m_cMemDC.Width = Screen.Width \ Screen.TwipsPerPixelY
  m_cMemDC.Height = 24
  
  Set m_fntSymbol = New StdFont
  m_fntSymbol.Name = "Marlett"
  m_fntSymbol.Size = Font.Size * 1.2
  
  Set m_cBrush = New cDottedBrush
  m_cBrush.Create
  
  ReDim RestoreList(10)
End Sub

Private Sub Class_Terminate()
  Set m_cMemDC = Nothing
  Set m_cNCM = Nothing
  Set m_fntSymbol = Nothing
  Set m_cBrush = Nothing
End Sub

Private Property Get ISubclass_MsgResponse() As EMsgResponse
  If CurrentMessage = WM_DESTROY Or _
     CurrentMessage = WM_INITMENUPOPUP Or _
     CurrentMessage = WM_WININICHANGE Or _
     CurrentMessage = WM_EXITMENULOOP Then
    ISubclass_MsgResponse = emrPreprocess
  Else
    ISubclass_MsgResponse = emrConsume
  End If
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Select Case iMsg
    Case WM_INITMENUPOPUP
      If (lParam And &HFFFF0000) = 0 Then OwnerDrawMenu wParam
    Case WM_MEASUREITEM
      ISubclass_WindowProc = MeasureItem(wParam, lParam)
    Case WM_DRAWITEM
      DrawItem wParam, lParam
    Case WM_WININICHANGE
      m_cNCM.GetMetrics
    Case WM_EXITMENULOOP
      pRestoreList
    Case WM_DESTROY
      Detach
  End Select
End Function
