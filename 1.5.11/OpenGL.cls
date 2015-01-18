VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOpenGL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' [OpenGL class]
' Contains functions to initialize and maintain an
' OpenGL context

Option Explicit

Private Type PIXELFORMATDESCRIPTOR
  nSize As Integer
  nVersion As Integer
  dwFlags As Long
  iPixelType As Byte
  cColorBits As Byte
  cRedBits As Byte
  cRedShift As Byte
  cGreenBits As Byte
  cGreenShift As Byte
  cBlueBits As Byte
  cBlueShift As Byte
  cAlphaBits As Byte
  cAlphaShift As Byte
  cAccumBits As Byte
  cAccumRedBits As Byte
  cAccumGreenBits As Byte
  cAccumBlueBits As Byte
  cAccumAlphaBits As Byte
  cDepthBits As Byte
  cStencilBits As Byte
  cAuxBuffers As Byte
  iLayerType As Byte
  bReserved As Byte
  dwLayerMask As Long
  dwVisibleMask As Long
  dwDamageMask As Long
End Type

Private Type BITMAPINFOHEADER   '40 bytes
  biSize As Long
  biWidth As Long
  biHeight As Long
  biPlanes As Integer
  biBitCount As Integer
  biCompression As Long
  biSizeImage As Long
  biXPelsPerMeter As Long
  biYPelsPerMeter As Long
  biClrUsed As Long
  biClrImportant As Long
End Type

Private Type RGBQUAD
  rgbBlue As Byte
  rgbGreen As Byte
  rgbRed As Byte
  rgbReserved As Byte
End Type

Private Type Bitmap  '24 bytes
  bmType As Long
  bmWidth As Long
  bmHeight As Long
  bmWidthBytes As Long
  bmPlanes As Integer
  bmBitsPixel As Integer
  bmBits As Long
End Type

Private Type BITMAPINFO
  bmiHeader As BITMAPINFOHEADER
  bmiColors As RGBQUAD
End Type

Private Type glBitmapType
  Name As String
  glID As Long
  Count As Integer
  Width As Integer
  Height As Integer
  AntiAlias As Boolean
'  Pixels() As RGBQUAD
End Type

' DC functions
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

' Bitmap functions
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hinst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Const LR_LOADFROMFILE = &H10
Private Const IMAGE_BITMAP = 0&

Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0&

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020

Private Declare Function glGetError Lib "opengl32.dll" () As Long

Private Declare Function ChoosePixelFormat Lib "gdi32" (ByVal hdc As Long, pPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) As Long
Private Declare Function SetPixelFormat Lib "gdi32" (ByVal hdc As Long, ByVal n As Long, pcPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) As Boolean
Private Declare Function GetPixelFormat Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DescribePixelFormat Lib "gdi32" (ByVal hdc As Long, ByVal n As Long, ByVal nBytes As Long, pcPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) As Long
 
Private Declare Function wglCreateContext Lib "opengl32.dll" (ByVal hdc As Long) As Long
Private Declare Function wglDeleteContext Lib "opengl32.dll" (ByVal HGLRC As Long) As Long
Private Declare Function wglMakeCurrent Lib "opengl32.dll" (ByVal hdc As Long, ByVal HGLRC As Long) As Long

Private Declare Sub glViewport Lib "opengl32.dll" (ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long)
Private Declare Sub gluOrtho2D Lib "glu32.dll" (ByVal Left As Double, ByVal Right As Double, ByVal Bottom As Double, ByVal Top As Double)
Private Declare Sub glMatrixMode Lib "opengl32.dll" (ByVal Mode As Long)
Private Const GL_PROJECTION = &H1701

Private Declare Sub glLoadIdentity Lib "opengl32.dll" ()

Private Declare Sub glFlush Lib "opengl32.dll" ()

Private Declare Sub glLineStipple Lib "opengl32.dll" (ByVal Factor As Long, ByVal Pattern As Long)

Private Declare Sub glPixelStorei Lib "opengl32.dll" (ByVal pname As Long, ByVal Param As Long)
Private Const GL_PACK_ALIGNMENT = &HD05
'Private Const GL_UNPACK_ALIGNMENT = &HCF5

Private Declare Sub glTexParameteri Lib "opengl32.dll" (ByVal Target As Long, ByVal pname As Long, ByVal Param As Long)
Private Const GL_TEXTURE_MIN_FILTER = &H2801
Private Const GL_TEXTURE_MAG_FILTER = &H2800
Private Const GL_LINEAR = &H2601
Private Const GL_NEAREST = &H2600

'Private Declare Sub glDrawPixels Lib "opengl32.dll" (ByVal Width As Long, ByVal Height As Long, ByVal Format As Long, ByVal pixelType As Long, Pixels As Any)
'Private Declare Sub glPixelZoom Lib "opengl32.dll" (ByVal xfactor As Single, ByVal yfactor As Single)
'Private Declare Sub glRasterPos2f Lib "opengl32.dll" (ByVal X As Single, ByVal Y As Single)

Private Declare Sub gluScaleImage Lib "glu32.dll" (ByVal Format As Long, ByVal widthIn As Long, ByVal heightIn As Long, ByVal typeIn As Long, dataIn As Any, ByVal widthOut As Long, ByVal heightOut As Long, ByVal typeOut As Long, dataOut As Any)

Private Declare Sub glTexImage2D Lib "opengl32.dll" (ByVal Target As Long, ByVal Level As Long, ByVal InternalFormat As Long, ByVal Width As Long, ByVal Height As Long, ByVal Border As Long, ByVal Format As Long, ByVal pixelType As Long, Pixels As Any)
Private Declare Sub glGetTexImage Lib "opengl32.dll" (ByVal Target As Long, ByVal Level As Long, ByVal Format As Long, ByVal pixelType As Long, Pixels As Any)
Private Const GL_TEXTURE_2D = &HDE1
Private Const GL_UNSIGNED_BYTE = &H1401
Private Const GL_BGRA_EXT As Long = &H80E1&

Private Declare Sub glGenTextures Lib "opengl32.dll" (ByVal n As Long, Arr As Any)
Private Declare Sub glDeleteTextures Lib "opengl32.dll" (ByVal n As Long, Arr As Any)
Private Declare Sub glBindTexture Lib "opengl32.dll" (ByVal Target As Long, ByVal Texture As Long)

Private Declare Function gluProject Lib "glu32.dll" (ByVal objx As Double, ByVal objy As Double, ByVal objz As Double, modelMatrix As Double, projMatrix As Double, viewport As Long, winx As Double, winy As Double, winz As Double) As Long
Private Declare Function gluUnProject Lib "glu32.dll" (ByVal winx As Double, ByVal winy As Double, ByVal winz As Double, modelMatrix As Double, projMatrix As Double, viewport As Long, objx As Double, objy As Double, objz As Double) As Long

Private Declare Sub glGetDoublev Lib "opengl32.dll" (ByVal pname As Long, params As Double)
Private Declare Sub glGetIntegerv Lib "opengl32.dll" (ByVal pname As Long, params As Long)

Private Declare Sub glEnableClientState Lib "opengl32.dll" (ByVal ArrayType As Long)
Private Declare Sub glDisableClientState Lib "opengl32.dll" (ByVal ArrayType As Long)
Private Const GL_VERTEX_ARRAY = &H8074&
Private Const GL_TEXTURE_COORD_ARRAY = &H8078&

Private Const GL_MODELVIEW_MATRIX = &HBA6
Private Const GL_PROJECTION_MATRIX = &HBA7
Private Const GL_VIEWPORT = &HBA2

Private Const PFD_DRAW_TO_WINDOW = &H4
Private Const PFD_DRAW_TO_BITMAP = &H8
Private Const PFD_SUPPORT_OPENGL = &H20

Private Const PFD_TYPE_RGBA = 0
Private Const PFD_MAIN_PLANE = 0

Private hBitmap As Long, hBitmapOrg As Long, _
        hRC As Long

Private mModelMatrix(15) As Double, _
        mProjectMatrix(15) As Double, _
        mViewport(3) As Long

Public myWidth As Long, myHeight As Long, _
       myLeft As Single, myRight As Single, _
       myTop As Single, myBottom As Single

Public hdc As Long

Private glBitmaps() As glBitmapType
Private numBitmaps As Integer

Public Function BitmapName(ByVal TextureID As Integer) As String
  BitmapName = glBitmaps(TextureID).Name
End Function

' Copy the OpenGL image to another hDC
Public Sub CopyTo(ByVal ResulthDC As Long)
  glFlush
  BitBlt ResulthDC, 0, 0, myWidth, myHeight, hdc, 0, 0, SRCCOPY
End Sub

' Decrements the count of the specified bitmap and
' deletes the bitmap from memory only if count = 0
Public Sub DeleteBitmap(ByVal TextureID As Integer)
  With glBitmaps(TextureID)
    .Count = .Count - 1
    If .Count = 0 Then
      .Name = ""
      glDeleteTextures 1, .glID
    End If
  End With
End Sub

' Delete all bitmaps
Public Sub DeleteBitmaps()
  Dim I As Integer
  For I = 0 To numBitmaps - 1
    If glBitmaps(I).Name <> "" Then _
      glDeleteTextures 1, glBitmaps(I).glID
  Next I
  ReDim glBitmaps(10)
  numBitmaps = 0
End Sub

' Selects the specified bitmap into the GL context
Public Sub GetBitmapDimensions(ByVal TextureID As Integer, ByRef Width As Integer, ByRef Height As Integer)
  With glBitmaps(TextureID)
    Width = .Width
    Height = .Height
  End With
End Sub

' Load the named bitmap into GL memory and return an
' identifier (integer). If the bitmap is already in memory,
' returns the identifier to the existing bitmap
' and increments the count
Public Function LoadBitmap(ByVal Filename As String, Optional AntiAlias As Boolean = True) As Integer
  Dim I As Integer, curBitmap As Integer
  Dim hand As Long, oldhand As Long, bmap As Bitmap, _
    srcedibbmap As BITMAPINFO, BytesPerScanLine As Long, _
    PadBytesPerScanLine As Long, hdcNew As Long, _
    glName As Long, BitmapBits() As RGBQUAD, ResizeBits() As RGBQUAD
  Dim TempFileName As String, _
    Width2Up As Long, Height2Up As Long
  
  If Not FileExists(Filename) Then
    LoadBitmap = -1
    Exit Function
  End If
  
  curBitmap = -1
  For I = 0 To numBitmaps - 1
    If StrComp(glBitmaps(I).Name, Filename, vbTextCompare) = 0 Then
      glBitmaps(I).Count = glBitmaps(I).Count + 1
      LoadBitmap = I
      Exit Function
    ElseIf glBitmaps(I).Name = "" And curBitmap = -1 Then
      curBitmap = I
    End If
  Next I
  
  If curBitmap = -1 Then
    If numBitmaps - 1 = UBound(glBitmaps) Then _
      ReDim Preserve glBitmaps(numBitmaps * 2)
    
    curBitmap = numBitmaps
    numBitmaps = numBitmaps + 1
  End If

  LoadTexture Filename, TempFileName

  hand = LoadImage(0, TempFileName, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE)
  
  If FileExists(TempFileName) Then Kill TempFileName
  
  If hand = 0 Then
    LoadBitmap = -1
    Exit Function
  End If
  
  'Fill out the BITMAP structure.
  GetObject hand, Len(bmap), bmap

  'Create a device context compatible with the Desktop.
  hdcNew = CreateCompatibleDC(0&)

  'Select the bitmap handle into the new device context.
  oldhand = SelectObject(hdcNew, hand)

  'srcedibbmap has been dimensioned as BITMAPINFO structure so
  'fill it out to create a template.
  'Two useful equations are those for BytesPerScanLine and
  'PadBytesPerScanLine. They work for any bit depth.
  'PadBytesPerScanLine will always be zero with biBiCount = 32
  'biheight is set negative to invert the "bottom up" scanline
  'reading.

  With srcedibbmap.bmiHeader
    .biSize = 40
    .biWidth = bmap.bmWidth
    .biHeight = -bmap.bmHeight
    .biPlanes = 1
    .biBitCount = 32
    .biCompression = BI_RGB
    BytesPerScanLine = ((((.biWidth * .biBitCount) + 31) \ 32) * 4)
    PadBytesPerScanLine = BytesPerScanLine - (((.biWidth * .biBitCount) + 7) \ 8)
    .biSizeImage = BytesPerScanLine * Abs(.biHeight)
  End With

  'Dimension the array receiving data. COLORQUAD is used just
  'to differentiate it from RGBQUAD.
  ReDim BitmapBits(1 To bmap.bmWidth, 1 To bmap.bmHeight)

  'Get color data from the source into a dib based on the template
  GetDIBits hdcNew, hand, 0, bmap.bmHeight, BitmapBits(1, 1), srcedibbmap, DIB_RGB_COLORS
  
  glGenTextures 1, glName
  glBindTexture GL_TEXTURE_2D, glName
  
  Width2Up = PowerOf2Up(bmap.bmWidth)
  Height2Up = PowerOf2Up(bmap.bmHeight)
  If Width2Up <> bmap.bmWidth Or Height2Up <> bmap.bmHeight Then
    ReDim ResizeBits(1 To Width2Up, 1 To Height2Up)
    SetScreenMousePointer vbHourglass
    gluScaleImage GL_BGRA_EXT, bmap.bmWidth, bmap.bmHeight, GL_UNSIGNED_BYTE, BitmapBits(1, 1), Width2Up, Height2Up, GL_UNSIGNED_BYTE, ResizeBits(1, 1)
    SetScreenMousePointer vbDefault
    glTexImage2D GL_TEXTURE_2D, 0, 4, Width2Up, Height2Up, 0, GL_BGRA_EXT, GL_UNSIGNED_BYTE, ResizeBits(1, 1)
  Else
    glTexImage2D GL_TEXTURE_2D, 0, 4, bmap.bmWidth, bmap.bmHeight, 0, GL_BGRA_EXT, GL_UNSIGNED_BYTE, BitmapBits(1, 1)
  End If

  'Clean up
  SelectObject hdcNew, oldhand
  DeleteObject hand
  DeleteDC hdcNew
  
  With glBitmaps(curBitmap)
    .Name = Filename
    .Count = 1
    .glID = glName
    .Width = bmap.bmWidth
    .Height = bmap.bmHeight
    .AntiAlias = AntiAlias
  End With
  LoadBitmap = curBitmap
End Function

' Resize the physical dimensions of the drawing surfaces
' to the given pixel dimensions
Public Sub PhysicalResize(ByVal Width As Long, ByVal Height As Long)
  Dim hBitmapOld As Long
  
  ' Create new bitmap
  hBitmap = CreateCompatibleBitmap(hdc, Width, Height)
  hBitmapOld = SelectObject(hdc, hBitmap)
  ' Delete old bitmap
  DeleteObject hBitmapOld
  
  myWidth = Width
  myHeight = Height
  
  wglMakeCurrent hdc, hRC
  glMatrixMode GL_PROJECTION
  glLoadIdentity
  glViewport 0, 0, Width, Height

  glGetDoublev GL_MODELVIEW_MATRIX, mModelMatrix(0)
  glGetDoublev GL_PROJECTION_MATRIX, mProjectMatrix(0)
  glGetIntegerv GL_VIEWPORT, mViewport(0)
End Sub

Public Function PixelToScale(ByVal X As Single, ByVal Y As Single, rX As Single, rY As Single)
  Dim resX As Double, resY As Double, resZ As Double

  gluUnProject X, myHeight - Y, 0, mModelMatrix(0), mProjectMatrix(0), mViewport(0), resX, resY, resZ
  rX = resX
  rY = resY
End Function

' Converts Scale to Pixel (Horz)
Public Function PixelX(ByVal Units As Single) As Single
  PixelX = (myWidth / (myRight - myLeft)) * Units
End Function

Public Function PowerOf2Up(ByVal X As Long) As Long
  Dim Value As Long
  Value = 1
  Do Until X <= Value
    Value = Value * 2
  Loop
  If Value > 1024 Then Value = 1024
  PowerOf2Up = Value
End Function

Public Function ScaleToPixel(ByVal X As Single, ByVal Y As Single, rX As Single, rY As Single)
  Dim resX As Double, resY As Double

  gluProject X, Y, 0, mModelMatrix(0), mProjectMatrix(0), mViewport(0), resX, resY, 0
  rX = resX
  rY = myHeight - resY
End Function

Public Function ScaleToPixelLong(ByVal X As Single, ByVal Y As Single, rX As Long, rY As Long)
  Dim resX As Double, resY As Double

  gluProject X, Y, 0, mModelMatrix(0), mProjectMatrix(0), mViewport(0), resX, resY, 0
  rX = resX
  rY = myHeight - resY
End Function

' Change viewport units to scale units
Public Function ScaleX(ByVal Real As Single) As Single
  If myWidth > 0 Then ScaleX = Real / (myWidth / (myRight - myLeft))
End Function

' Selects the specified bitmap into the GL context
Public Sub SelectBitmap(ByVal TextureID As Integer)
  With glBitmaps(TextureID)
    glBindTexture GL_TEXTURE_2D, .glID
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, IIf(.AntiAlias, GL_LINEAR, GL_NEAREST)
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, IIf(.AntiAlias, GL_LINEAR, GL_NEAREST)
  End With
End Sub

' Set the drawing scale of the GL context
Public Sub SetScale(ByVal Left As Single, ByVal Top As Single, ByVal Right As Single, ByVal Bottom As Single)
  myLeft = Left
  myTop = Top
  myRight = Right
  myBottom = Bottom
  gluOrtho2D Left, Right, Bottom, Top

  glGetDoublev GL_MODELVIEW_MATRIX, mModelMatrix(0)
  glGetDoublev GL_PROJECTION_MATRIX, mProjectMatrix(0)
  glGetIntegerv GL_VIEWPORT, mViewport(0)
End Sub

' Select this OpenGL object as the current drawing context
Public Sub StartDraw()
  wglMakeCurrent hdc, hRC
End Sub

Public Sub WriteBitmap(ByVal TextureID As Integer, ByVal Filename As String)
  Dim Bits() As RGBQUAD, BitHeader() As Byte
  Dim FileNum As Integer
  Dim Byte1 As Byte, Byte2 As Byte
  Dim I As Long
  SelectBitmap TextureID
  With glBitmaps(TextureID)
    ReDim Bits(1 To PowerOf2Up(.Width) * PowerOf2Up(.Height))
    glGetTexImage GL_TEXTURE_2D, 0, GL_BGRA_EXT, GL_UNSIGNED_BYTE, Bits(1)
    BitHeader = LoadResData("DATA4", 10)
    If FileExists(Filename) Then Kill Filename
    FileNum = FreeFile
    Open Filename For Binary As #FileNum
    Put #FileNum, , BitHeader
    For I = 1 To UBound(Bits)
      Byte1 = (Bits(I).rgbRed * 31 / 255) + (Bits(I).rgbGreen * 63 / 255 Mod 8) * 32
      Byte2 = (Bits(I).rgbGreen * 63 / 255) \ 8 + (Bits(I).rgbBlue * 31 / 255) * 8
      Put #FileNum, , Byte1
      Put #FileNum, , Byte2
    Next I
    Put #FileNum, 19, PowerOf2Up(.Width)
    Put #FileNum, 23, PowerOf2Up(.Height)
    Close #FileNum
  End With
End Sub

' Initialize the drawing surface
Private Sub Class_Initialize()
  Dim PixFormat As Long, pfd As PIXELFORMATDESCRIPTOR
  
  hdc = CreateCompatibleDC(0&)
  
  ' Create a color bitmap (temporary placeholder)
  hBitmap = CreateCompatibleBitmap(frmMain.hdc, 1, 1)
  hBitmapOrg = SelectObject(hdc, hBitmap)
    
  With pfd
    .nSize = Len(pfd)
    .nVersion = 1
    .dwFlags = PFD_DRAW_TO_WINDOW Or PFD_DRAW_TO_BITMAP Or PFD_SUPPORT_OPENGL
    .iPixelType = PFD_TYPE_RGBA
    .cColorBits = 24
    .cDepthBits = 32
    .iLayerType = PFD_MAIN_PLANE
  End With

  glGetError
  PixFormat = ChoosePixelFormat(hdc, pfd)
  If PixFormat = 0 Then
    MsgBox "OpenGL initialization failed", vbCritical
  End If
  
  SetPixelFormat hdc, PixFormat, pfd
  hRC = wglCreateContext(hdc)
  If hRC = 0 Then
    MsgBox "OpenGL initialization failed", vbCritical
  End If

  wglMakeCurrent hdc, hRC
  glLineStipple 1, 32752
  glPixelStorei GL_PACK_ALIGNMENT, 4
  glEnableClientState GL_VERTEX_ARRAY
  ReDim glBitmaps(10)
  numBitmaps = 0
End Sub

' Delete everything
Private Sub Class_Terminate()
  DeleteBitmaps
  ' Free all resources
  If hBitmapOrg > 0 Then SelectObject hdc, hBitmapOrg
  wglMakeCurrent vbNull, vbNull
  wglDeleteContext hRC
  DeleteObject hBitmap
  DeleteDC hdc
End Sub

