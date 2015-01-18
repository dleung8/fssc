Attribute VB_Name = "basOpenGL"
' [basOpenGL]
' Contains functions to draw in an OpenGL context

Option Explicit

Public Type PointType
  X As Single
  Y As Single
End Type

Private Type RGBColor
  R As Byte
  G As Byte
  B As Byte
  Reserved As Byte
End Type

Public Type TessType
  Mode As Long
  Pts() As PointType
  'EdgeFlags() As Boolean
End Type

Private Declare Function gluNewQuadric Lib "glu32.dll" () As Long
Private Declare Sub gluDeleteQuadric Lib "glu32.dll" (ByVal State As Long)

Private Declare Sub glBegin Lib "opengl32.dll" (ByVal Mode As Long)
Private Declare Sub glEnable Lib "opengl32.dll" (ByVal cap As Long)
Private Declare Sub glDisable Lib "opengl32.dll" (ByVal cap As Long)

Private Declare Sub gluDisk Lib "glu32.dll" (ByVal qobj As Long, ByVal innerRadius As Double, ByVal outerRadius As Double, ByVal Slices As Integer, ByVal Loops As Integer)
Private Declare Sub glVertex2f Lib "opengl32.dll" (ByVal X As Single, ByVal Y As Single)
'Private Declare Sub glVertex3f Lib "opengl32.dll" (ByVal X As Single, ByVal Y As Single, ByVal Z As Single)
Private Declare Sub glClear Lib "opengl32.dll" (ByVal Mask As Long)

Private Declare Sub glLoadIdentity Lib "opengl32.dll" ()

Private Declare Sub glVertexPointer Lib "opengl32.dll" (ByVal Size As Integer, ByVal VarSize As Long, ByVal Stride As Integer, Pointer As Any)
'Private Declare Sub glEdgeFlagPointer Lib "opengl32.dll" (ByVal Stride As Integer, Pointer As Any)
Private Declare Sub glTexCoordPointer Lib "opengl32.dll" (ByVal Size As Integer, ByVal VarSize As Long, ByVal Stride As Integer, Pointer As Any)
Private Declare Sub glDrawArrays Lib "opengl32.dll" (ByVal Mode As Long, ByVal First As Long, ByVal Count As Long)
Private Const GL_FLOAT = &H1406

Private Declare Sub glEnableClientState Lib "opengl32.dll" (ByVal ArrayType As Long)
Private Declare Sub glDisableClientState Lib "opengl32.dll" (ByVal ArrayType As Long)
Private Const GL_VERTEX_ARRAY = &H8074&
Private Const GL_TEXTURE_COORD_ARRAY = &H8078&

Private Declare Sub glTexCoord2f Lib "opengl32.dll" (ByVal X As Single, ByVal Y As Single)
Private Const GL_TEXTURE_2D = &HDE1

Private Declare Function gluNewTess Lib "glu32.dll" () As Long
Private Declare Sub gluTessCallback Lib "glu32.dll" (ByVal Tess As Long, ByVal Which As Long, ByVal fn As Long)
Private Declare Sub gluBeginPolygon Lib "glu32.dll" (ByVal Tess As Long)
Private Declare Sub gluEndPolygon Lib "glu32.dll" (ByVal Tess As Long)
Private Declare Sub gluTessVertex Lib "glu32.dll" (ByVal Tess As Long, Coord As Double, ByVal Data As Long)
Private Declare Sub gluTessProperty Lib "glu32.dll" (ByVal Tess As Long, ByVal Which As Long, ByVal Value As Double)
Private Const GLU_TESS_WINDING_RULE = 100140
Private Const GLU_TESS_WINDING_ODD = 100130
Private Const GLU_TESS_WINDING_NONZERO = 100131

'Private Declare Sub glEdgeFlag Lib "opengl32.dll" (ByVal Value As Boolean)

Private Declare Function gluDeleteTess Lib "glu32.dll" (ByVal Tess As Long) As Long

Private Declare Function gluErrorString Lib "glu32.dll" (ByVal errCode As Long) As String

Private Const GLU_TESS_BEGIN = 100100
Private Const GLU_TESS_VERTEX = 100101
Private Const GLU_TESS_END = 100102
Private Const GLU_TESS_ERROR = 100103
'Private Const GLU_TESS_EDGE_FLAG = 100104
Private Const GLU_TESS_COMBINE = 100105

Private TessOriginal() As PointType
Private TessResult() As TessType
Private TessResultCount As Integer
Private TessSubResultCount As Integer
Private TessOriginalCount As Integer

'Modes
Private Const GL_POINTS = &H0
Private Const GL_LINES = &H1
Private Const GL_LINE_LOOP = &H2
Private Const GL_LINE_STRIP = &H3
Private Const GL_TRIANGLES = &H4
Private Const GL_TRIANGLE_STRIP = &H5
Private Const GL_TRIANGLE_FAN = &H6
Private Const GL_QUADS = &H7&
Private Const GL_QUAD_STRIP = &H8
Private Const GL_POLYGON = &H9

Private Const GL_COLOR_BUFFER_BIT = &H4000
Private Const GL_DEPTH_BUFFER_BIT = &H100

Private Const GL_LINE_STIPPLE = &HB24

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

' Flush the buffer
Public Declare Sub glFlush Lib "opengl32.dll" ()

' Save Matrix attributes
Public Declare Sub glPushMatrix Lib "opengl32.dll" ()
Public Declare Sub glPopMatrix Lib "opengl32.dll" ()

' Translate a drawing
Public Declare Sub glTranslatef Lib "opengl32.dll" (ByVal X As Single, ByVal Y As Single, ByVal Z As Single)

' Rotate objects
Public Declare Sub glRotatef Lib "opengl32.dll" (ByVal Angle As Single, ByVal X As Single, ByVal Y As Single, ByVal Z As Single)

' End a line/polygon draw
Public Declare Sub glEnd Lib "opengl32.dll" ()

' Change the color
Public Declare Sub glColor3f Lib "opengl32.dll" (ByVal Red As Single, ByVal Green As Single, ByVal Blue As Single)
Public Declare Sub glClearColor Lib "opengl32.dll" (ByVal Red As Single, ByVal Green As Single, ByVal Blue As Single, ByVal alpha As Single)

' Change the width of pen
Public Declare Sub glPointSize Lib "opengl32.dll" (ByVal Size As Single)
Public Declare Sub glLineWidth Lib "opengl32.dll" (ByVal Width As Single)

Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long

Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Const HOLLOW_BRUSH = 5
Private Const BLACK_PEN = 7

'Private CurrentEdgeFlag As Boolean

' Lists
'Public Declare Function glGenLists Lib "opengl32.dll" (ByVal Range As Integer) As Integer
'Public Declare Sub glNewList Lib "opengl32.dll" (ByVal List As Integer, ByVal Mode As Long)
'Public Declare Sub glEndList Lib "opengl32.dll" ()
'Public Declare Sub glCallList Lib "opengl32.dll" (ByVal List As Integer)
'Public Declare Sub glDeleteLists Lib "opengl32.dll" (ByVal List As Integer, ByVal Range As Integer)
'Public Const GL_COMPILE = &H1300

' Begin drawing lines
Public Sub glBeginLines()
  glBegin GL_LINES
End Sub

' Begin drawing lines
Public Sub glBeginPoints()
  glBegin GL_POINTS
End Sub

' Clear the drawing surface
Public Sub glCls()
  glClear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
End Sub

' Enable dashed lines
Public Sub glDashLine()
  glEnable GL_LINE_STIPPLE
End Sub

' Draw a circle. Must have called glBeginCircle
Public Sub glDrawCircle(ByVal X As Single, ByVal Y As Single, ByVal Radius As Single, ByVal Color As Long)
'  Dim obj As Long
'  glPushMatrix
'  glTranslatef X, Y, 0
'  glBegin GL_LINE_LOOP
'  obj = gluNewQuadric
'  gluDisk obj, Radius, Radius, 50, 1
'  gluDeleteQuadric obj
'  glEnd
'  glPopMatrix
  
  Dim hPen As Long, oldPen As Long, _
    hBrush As Long, oldBrush As Long
  Dim X1 As Single, Y1 As Single, Radius2 As Single

  glFlush

  ' Create pen and brush
  hPen = CreatePen(vbSolid, 0, Color)
  hBrush = GetStockObject(HOLLOW_BRUSH)
  ' Set the pen and brush
  oldPen = SelectObject(picEditor.hdc, hPen)
  oldBrush = SelectObject(picEditor.hdc, hBrush)

  ' Draw the circle
  picEditor.ScaleToPixel X, Y, X1, Y1
  Radius2 = picEditor.PixelX(Radius)
  Ellipse picEditor.hdc, X1 - Radius2, Y1 - Radius2, X1 + Radius2 + 1, Y1 + Radius2 + 1

  ' Restore defaults and free resources
  SelectObject picEditor.hdc, oldPen
  SelectObject picEditor.hdc, oldBrush
  DeleteObject hPen
End Sub

' Draw a line
Public Sub glDrawLine(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single)
  glVertex2f X1, Y1
  glVertex2f X2, Y2
End Sub

' Draw connected lines
Public Sub glDrawLines(Pts() As PointType)
  glVertexPointer 2, GL_FLOAT, 0, Pts(0).X
  glDrawArrays GL_LINE_STRIP, 0, UBound(Pts) + 1
End Sub

' Draw a point
Public Sub glDrawPoint(ByVal X As Single, ByVal Y As Single)
  glVertex2f X, Y
End Sub

' Draw an outlined polygon
Public Sub glDrawPolygon(Pts() As PointType)
  glVertexPointer 2, GL_FLOAT, 0, Pts(0).X
  glDrawArrays GL_LINE_LOOP, 0, UBound(Pts) + 1
End Sub

' Draw an outlined polygon
Public Sub glDrawPolygon2(Pts() As PointType, ByVal Color As Long, ByVal DefX As Single, ByVal DefY As Single)
  Dim hPen As Long, oldPen As Long
  Dim X1 As Single, Y1 As Single
  
  Dim I As Integer
  
  glFlush
  
  ' Create pen
  hPen = CreatePen(vbSolid, 0, Color)
  ' Set the pen
  oldPen = SelectObject(picEditor.hdc, hPen)
  
  ' Draw
  picEditor.ScaleToPixel Pts(UBound(Pts)).X + DefX, Pts(UBound(Pts)).Y + DefY, X1, Y1
  MoveToEx picEditor.hdc, X1, Y1, 0&
  For I = 0 To UBound(Pts)
    picEditor.ScaleToPixel Pts(I).X + DefX, Pts(I).Y + DefY, X1, Y1
    LineTo picEditor.hdc, X1, Y1
  Next I

  ' Restore defaults and free resources
  SelectObject picEditor.hdc, oldPen
  DeleteObject hPen
End Sub

' Draw an outlined rectangle
Public Sub glDrawRect(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single)
  glBegin GL_LINE_LOOP
  glVertex2f X1, Y1
  glVertex2f X1, Y2
  glVertex2f X2, Y2
  glVertex2f X2, Y1
  glEnd
End Sub

Public Sub glDrawQuads(Pts() As PointType)
  glVertexPointer 2, GL_FLOAT, 0, Pts(0).X
  glDrawArrays GL_QUADS, 0, UBound(Pts) + 1
End Sub

' Paint a region based on tesselated
' coordinates with the forecolor
Public Sub glDrawTesselated(TessPts() As TessType)
  Dim I As Integer
  For I = 0 To UBound(TessPts)
    glVertexPointer 2, GL_FLOAT, 0, TessPts(I).Pts(0)
    glDrawArrays TessPts(I).Mode, 0, UBound(TessPts(I).Pts) + 1
  Next I
End Sub

' Paint a region based on tesselated
' coordinates with the current texture
Public Sub glDrawTesselatedTextured(TessPts() As TessType, ByVal XM As Single, ByVal XB As Single, ByVal YM As Single, ByVal YB As Single)
  Dim I As Integer, J As Integer
  glColor3f 1, 1, 1
  glEnable GL_TEXTURE_2D
  For I = 0 To UBound(TessPts)
    glBegin TessPts(I).Mode
    For J = 0 To UBound(TessPts(I).Pts)
      With TessPts(I).Pts(J)
        ' (Negate Y slope because grid coordinates decreases
        ' as screen coordinates increase)
        glTexCoord2f (.X - XB) * XM, (.Y - YB) * -YM
        glVertex2f .X, .Y
      End With
    Next J
    glEnd
  Next I
  glDisable GL_TEXTURE_2D
End Sub

' Change the color
Public Sub glForeColor(ByVal Color As Long)
  Dim Temp As RGBColor
  CopyMemory Temp, Color, 4
  glColor3f Temp.R / 255, Temp.G / 255, Temp.B / 255
End Sub

' Paint a region with the forecolor
Public Sub glPaintRegion(Pts() As PointType)
  glVertexPointer 2, GL_FLOAT, 0, Pts(0).X
  glDrawArrays GL_POLYGON, 0, UBound(Pts) + 1
End Sub

' Paint a region with the current texture, using a linear texture coordinate interpolation
Public Sub glPaintTexturedRegion1(Pts() As PointType, ByVal XM As Single, ByVal XB As Single, ByVal YM As Single, ByVal YB As Single)
  Dim I As Integer
  glColor3f 1, 1, 1
  glEnable GL_TEXTURE_2D
  glBegin GL_POLYGON
  For I = 0 To UBound(Pts)
    With Pts(I)
      ' (Negate Y slope because grid coordinates decreases
      ' as screen coordinates increase)
      glTexCoord2f (.X - XB) * XM, (.Y - YB) * -YM
      glVertex2f .X, .Y
    End With
  Next I
  glEnd
  glDisable GL_TEXTURE_2D
End Sub

' Paint a region with the current texture using the corresponding texture coordinate
Public Sub glPaintTexturedRegion2(Pts() As PointType, Coords() As PointType)
  glColor3f 1, 1, 1
  glEnable GL_TEXTURE_2D
  glEnableClientState GL_TEXTURE_COORD_ARRAY
  glVertexPointer 2, GL_FLOAT, 0, Pts(0).X
  glTexCoordPointer 2, GL_FLOAT, 0, Coords(0).X
  glDrawArrays GL_POLYGON, 0, UBound(Pts) + 1
  glDisableClientState GL_TEXTURE_COORD_ARRAY
  glDisable GL_TEXTURE_2D
End Sub

' Disable dashed lines
Public Sub glSmoothLine()
  glDisable GL_LINE_STIPPLE
End Sub

' Tesselate a polygon
Public Sub glTesselate(Pts() As PointType, Result() As TessType, ByVal AlternateMode As Boolean)
  Dim TessObj As Long
  Dim I As Integer
  Dim Coords(4) As Double
  
  TessOriginal = Pts
  TessOriginalCount = UBound(Pts)
  
  ReDim TessResult(5)
  TessResultCount = -1
  
  TessObj = gluNewTess()

  gluTessProperty TessObj, GLU_TESS_WINDING_RULE, IIf(AlternateMode, GLU_TESS_WINDING_ODD, GLU_TESS_WINDING_NONZERO)

  ' Set up call backs
  gluTessCallback TessObj, GLU_TESS_BEGIN, AddressOf TessBegin
  gluTessCallback TessObj, GLU_TESS_COMBINE, AddressOf TessCombine
  'gluTessCallback TessObj, GLU_TESS_EDGE_FLAG, AddressOf TessEdgeFlag
  gluTessCallback TessObj, GLU_TESS_END, AddressOf TessEnd
  gluTessCallback TessObj, GLU_TESS_ERROR, AddressOf TessError
  gluTessCallback TessObj, GLU_TESS_VERTEX, AddressOf TessVertex
  
  'CurrentEdgeFlag = True
  
  gluBeginPolygon TessObj
  For I = 0 To UBound(Pts)
    Coords(0) = Pts(I).X
    Coords(1) = Pts(I).Y
    ' Since gluTessVertex takes an opaque pointer, we'll
    ' just send it booby values (i.e. 0-based index of
    ' the array)
    gluTessVertex TessObj, Coords(0), I
  Next I
  gluEndPolygon TessObj
  gluDeleteTess TessObj
  If TessResultCount = -1 Then
    ReDim TessResult(0)
    TessResult(0).Mode = GL_LINE_LOOP
    TessResult(0).Pts = Pts
  Else
    ReDim Preserve TessResult(TessResultCount)
  End If
  Result = TessResult
  Erase TessResult
  Erase TessOriginal
End Sub

Private Sub TessBegin(ByVal PolygonType As Long)
  TessResultCount = TessResultCount + 1
  If TessResultCount > UBound(TessResult) Then _
    ReDim Preserve TessResult(TessResultCount * 2)
  TessResult(TessResultCount).Mode = PolygonType
  ReDim TessResult(TessResultCount).Pts(5)
  'ReDim TessResult(TessResultCount).EdgeFlags(5)
  TessSubResultCount = -1
End Sub

Private Sub TessCombine(ByVal CoordLocation As Long, ByVal VertexLocation As Long, ByVal WeightLocation As Long, Index As Long)
  Dim Coords(2) As Double
  CopyMemory Coords(0), ByVal CoordLocation, 8 * 3
  TessOriginalCount = TessOriginalCount + 1
  If TessOriginalCount > UBound(TessOriginal) Then _
    ReDim Preserve TessOriginal(TessOriginalCount * 2)
  TessOriginal(TessOriginalCount).X = Coords(0)
  TessOriginal(TessOriginalCount).Y = Coords(1)
  Index = TessOriginalCount
End Sub
'
'Private Sub TessEdgeFlag(ByVal Flag As Boolean)
'  CurrentEdgeFlag = Flag
'End Sub

Private Sub TessEnd()
  ReDim Preserve TessResult(TessResultCount).Pts(TessSubResultCount)
  'ReDim Preserve TessResult(TessResultCount).EdgeFlags(TessSubResultCount)
End Sub

Private Sub TessError(ByVal errCode As Long)
  MsgBox "OpenGL tesselation error: " & gluErrorString(errCode), vbInformation
End Sub

Private Sub TessVertex(ByVal Index As Long)
  TessSubResultCount = TessSubResultCount + 1
  With TessResult(TessResultCount)
    If TessSubResultCount > UBound(.Pts) Then
      ReDim Preserve .Pts(TessSubResultCount * 2)
      'ReDim Preserve .EdgeFlags(TessSubResultCount * 2)
    End If
    .Pts(TessSubResultCount) = TessOriginal(Index)
'    .EdgeFlags(TessSubResultCount) = CurrentEdgeFlag
  End With
End Sub
