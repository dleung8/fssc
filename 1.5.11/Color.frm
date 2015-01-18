VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmColor 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   Icon            =   "Color.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   369
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   273
   ShowInTaskbar   =   0   'False
   Tag             =   "2160"
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   61
      Tag             =   "1031"
      Top             =   4920
      WhatsThisHelpID =   1031
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   60
      Tag             =   "1030"
      Top             =   4920
      WhatsThisHelpID =   1030
      Width           =   1215
   End
   Begin VB.CheckBox chkNight 
      Height          =   195
      Left            =   240
      TabIndex        =   59
      Tag             =   "2168"
      Top             =   4560
      WhatsThisHelpID =   2168
      Width           =   3375
   End
   Begin ComctlLib.Slider sldTransparency 
      Height          =   495
      Left            =   840
      TabIndex        =   58
      Top             =   3900
      Visible         =   0   'False
      WhatsThisHelpID =   2166
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      _Version        =   327682
      LargeChange     =   1
      Max             =   15
      SelStart        =   15
      Value           =   15
   End
   Begin VB.PictureBox picRGB 
      BackColor       =   &H00FFFFFF&
      ClipControls    =   0   'False
      Height          =   255
      Left            =   480
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   55
      Top             =   3900
      Visible         =   0   'False
      WhatsThisHelpID =   2165
      Width           =   300
   End
   Begin VB.OptionButton optColor 
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Tag             =   "2163"
      Top             =   2670
      WhatsThisHelpID =   2162
      Width           =   3255
   End
   Begin VB.PictureBox picColorFocus 
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   255
      Left            =   3600
      ScaleHeight     =   255
      ScaleWidth      =   375
      TabIndex        =   4
      Top             =   600
      WhatsThisHelpID =   2162
      Width           =   375
   End
   Begin VB.OptionButton optColor 
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Tag             =   "2162"
      Top             =   630
      WhatsThisHelpID =   2162
      Width           =   3255
   End
   Begin VB.OptionButton optColor 
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Tag             =   "2161"
      Top             =   240
      WhatsThisHelpID =   2161
      Width           =   3255
   End
   Begin VB.OptionButton optColor 
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Tag             =   "2165"
      Top             =   3390
      Visible         =   0   'False
      WhatsThisHelpID =   2165
      Width           =   3255
   End
   Begin VB.Label lbls 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Height          =   195
      Index           =   1
      Left            =   3600
      TabIndex        =   57
      Tag             =   "2167"
      Top             =   3660
      Visible         =   0   'False
      WhatsThisHelpID =   2166
      Width           =   45
   End
   Begin VB.Label lbls 
      AutoSize        =   -1  'True
      Height          =   195
      Index           =   0
      Left            =   960
      TabIndex        =   56
      Tag             =   "2166"
      Top             =   3660
      Visible         =   0   'False
      WhatsThisHelpID =   2166
      Width           =   45
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   22
      Left            =   3120
      TabIndex        =   54
      Top             =   3000
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   21
      Left            =   2745
      TabIndex        =   53
      Top             =   3000
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   20
      Left            =   2370
      TabIndex        =   52
      Top             =   3000
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   19
      Left            =   1995
      TabIndex        =   51
      Top             =   3000
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   18
      Left            =   1620
      TabIndex        =   50
      Top             =   3000
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   17
      Left            =   1245
      TabIndex        =   49
      Top             =   3000
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   16
      Left            =   870
      TabIndex        =   48
      Top             =   3000
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   15
      Left            =   495
      TabIndex        =   47
      Top             =   3000
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   31
      Left            =   3495
      TabIndex        =   46
      Top             =   2280
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   30
      Left            =   3120
      TabIndex        =   45
      Top             =   2280
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   29
      Left            =   2745
      TabIndex        =   44
      Top             =   2280
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   28
      Left            =   2370
      TabIndex        =   43
      Top             =   2280
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   27
      Left            =   1995
      TabIndex        =   42
      Top             =   2280
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   26
      Left            =   1620
      TabIndex        =   41
      Top             =   2280
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   25
      Left            =   1245
      TabIndex        =   40
      Top             =   2280
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   24
      Left            =   870
      TabIndex        =   39
      Top             =   2280
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   23
      Left            =   495
      TabIndex        =   38
      Top             =   2280
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   40
      Left            =   3495
      TabIndex        =   37
      Top             =   1950
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   39
      Left            =   3120
      TabIndex        =   36
      Top             =   1950
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   38
      Left            =   2745
      TabIndex        =   35
      Top             =   1950
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   37
      Left            =   2370
      TabIndex        =   34
      Top             =   1950
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   36
      Left            =   1995
      TabIndex        =   33
      Top             =   1950
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   35
      Left            =   1620
      TabIndex        =   32
      Top             =   1950
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   34
      Left            =   1245
      TabIndex        =   31
      Top             =   1950
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   33
      Left            =   870
      TabIndex        =   30
      Top             =   1950
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   32
      Left            =   495
      TabIndex        =   29
      Top             =   1950
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   49
      Left            =   3495
      TabIndex        =   28
      Top             =   1620
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   48
      Left            =   3120
      TabIndex        =   27
      Top             =   1620
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   47
      Left            =   2745
      TabIndex        =   26
      Top             =   1620
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   46
      Left            =   2370
      TabIndex        =   25
      Top             =   1620
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   45
      Left            =   1995
      TabIndex        =   24
      Top             =   1620
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   44
      Left            =   1620
      TabIndex        =   23
      Top             =   1620
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   43
      Left            =   1245
      TabIndex        =   22
      Top             =   1620
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   42
      Left            =   870
      TabIndex        =   21
      Top             =   1620
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   41
      Left            =   495
      TabIndex        =   20
      Top             =   1620
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   13
      Left            =   3495
      TabIndex        =   19
      Top             =   1290
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   12
      Left            =   3120
      TabIndex        =   18
      Top             =   1290
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   11
      Left            =   2745
      TabIndex        =   17
      Top             =   1290
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   10
      Left            =   2370
      TabIndex        =   16
      Top             =   1290
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   9
      Left            =   1995
      TabIndex        =   15
      Top             =   1290
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   8
      Left            =   1620
      TabIndex        =   14
      Top             =   1290
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   7
      Left            =   1245
      TabIndex        =   13
      Top             =   1290
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   6
      Left            =   870
      TabIndex        =   12
      Top             =   1290
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   5
      Left            =   495
      TabIndex        =   11
      Top             =   1290
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   14
      Left            =   3495
      TabIndex        =   10
      Top             =   960
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   4
      Left            =   1995
      TabIndex        =   9
      Top             =   960
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   3
      Left            =   1620
      TabIndex        =   8
      Top             =   960
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   1245
      TabIndex        =   7
      Top             =   960
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   870
      TabIndex        =   6
      Top             =   960
      WhatsThisHelpID =   2162
      Width           =   300
   End
   Begin VB.Label lblColors 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   495
      TabIndex        =   5
      Top             =   960
      WhatsThisHelpID =   2162
      Width           =   300
   End
End
Attribute VB_Name = "frmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSelection As Integer, _
        mFocus As Integer, _
        mChanged As Boolean

Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long

Private Sub DrawFocus(ByVal Selected As Boolean)
  Dim R As RECT
  If mFocus = 255 Then
    With R
      .Left = picRGB.Left - 3
      .Top = picRGB.Top - 3
      .Right = picRGB.Left + 23
      .Bottom = picRGB.Top + 20
    End With
  ElseIf mFocus > 49 Then
    Exit Sub
  Else
    With R
      .Left = lblColors(mFocus).Left - 3
      .Top = lblColors(mFocus).Top - 3
      .Right = lblColors(mFocus).Left + 23
      .Bottom = lblColors(mFocus).Top + 20
    End With
  End If
  
  Line (R.Left, R.Top)-(R.Right - 1, R.Bottom - 1), vbButtonFace, B
  If Selected Then DrawFocusRect hdc, R
End Sub

Private Sub DrawSelection(ByVal Selected As Boolean)
  Dim X As Integer, Y As Integer
  If mSelection = 254 Then
    Exit Sub
  ElseIf mSelection = 255 Then
    X = picRGB.Left - 1
    Y = picRGB.Top - 1
  Else
    X = lblColors(mSelection).Left - 1
    Y = lblColors(mSelection).Top - 1
  End If
    
  Line (X, Y)-(X + 21, Y + 18), IIf(Selected, vbBlack, BackColor), B
End Sub

' Shows the Color edit dialog box, modifying Data as necessary
Public Function EditData(ByRef Data As Long, Optional ByVal NoExtend As Integer = 0) As Boolean
  Dim I As Integer
  
  Load frmColor
  
  If Options.FSVersion >= Version_FS2K And NoExtend <> 2 Then
    optColor(3).Visible = True
    picRGB.Visible = True
    lbls(0).Visible = True
    lbls(1).Visible = True
    sldTransparency.Visible = True
    If NoExtend = 1 Then chkNight.Visible = False
  Else
    optColor(2).Caption = Lang.GetString(RES_Col_Night)
    chkNight.Visible = False
    cmdOK.Top = 232
    cmdCancel.Top = 232
    Height = 4470
  End If
  CenterForm Me
  
  If (Data And PalMask) > 0 Then
    sldTransparency.Value = 15
    mSelection = Data And PalColormask
    optColor(IIf(Between(mSelection, 15, 22), 2, 1)).Value = True
  ElseIf Data = 0 Or Options.FSVersion < Version_FS2K Then
    optColor(0).Value = True
    mSelection = 254
  Else
    optColor(3).Value = True
    sldTransparency.Value = (Data And TransparentMask) / &H1000000
    mSelection = 255
    picRGB.BackColor = Data And RGBColorMask
  End If
  
  If Options.FSVersion >= Version_FS2K Then
    chkNight.Value = -((Data And NightMask) > 0)
  End If
  
  mFocus = mSelection
  mChanged = False
  Show vbModal, Screen.ActiveForm
  If mChanged Then
    Data = 0
    If chkNight.Value = vbChecked Then Data = NightMask
    If mSelection = 255 Then
      Data = Data Or RGBMask Or picRGB.BackColor Or &H1000000 * sldTransparency.Value
    ElseIf mSelection = 254 Then
      Data = 0
    Else
      Data = Data Or PalMask Or mSelection
    End If
    EditData = True
  End If
  Unload frmColor
End Function

Private Sub cmdCancel_Click()
  Hide
End Sub

Private Sub cmdOK_Click()
  mChanged = True
  Hide
End Sub

Private Sub Form_Load()
  Dim I As Integer
  DialogMenus Me
  Lang.PrepareForm Me
  
  For I = 0 To lblColors.UBound
    lblColors(I).BackColor = FSColors(I)
  Next I
End Sub

Private Sub Form_Paint()
  DrawSelection True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then Cancel = True: Hide
End Sub

Private Sub lblColors_DblClick(Index As Integer)
  cmdOK_Click
End Sub

Private Sub lblColors_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim TestNum As Integer
  picColorFocus.SetFocus
  TestNum = IIf(Between(Index, 15, 22), 2, 1)
  DrawFocus False
  If Not optColor(TestNum).Value Then optColor(TestNum).Value = True
  DrawSelection False
  mSelection = Index
  mFocus = Index
  DrawSelection True
  DrawFocus True
End Sub

Private Sub optColor_Click(Index As Integer)
  If Index = 0 And mSelection <> 254 Then
    DrawSelection False
    mSelection = 254
    mFocus = mSelection
    chkNight.Value = vbUnchecked
  ElseIf Index = 1 And (Between(mSelection, 15, 22) Or (mSelection > 200)) Then
    DrawSelection False
    mSelection = 0
    mFocus = mSelection
  ElseIf Index = 2 And Not Between(mSelection, 15, 22) Then
    DrawSelection False
    mSelection = 15
    mFocus = mSelection
  ElseIf Index = 3 And mSelection <> 255 Then
    DrawSelection False
    mSelection = 255
    mFocus = mSelection
  End If
  chkNight.Enabled = (Index > 0)
  DrawSelection True
End Sub

Private Sub picColorFocus_GotFocus()
  If mSelection < 200 Then
    mFocus = mSelection
  Else
    mFocus = 0
  End If
  DrawFocus True
End Sub

Private Sub picColorFocus_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift And vbAltMask Then Exit Sub
  Select Case KeyCode
    Case vbKeyLeft
      If mFocus <> 0 And mFocus <> 5 And mFocus <> 41 And mFocus <> 32 And mFocus <> 23 And mFocus <> 15 Then
        DrawFocus False
        If mFocus = 14 Then
          mFocus = 4
        Else
          mFocus = mFocus - 1
        End If
        DrawFocus True
      End If
    Case vbKeyRight
      If mFocus <> 14 And mFocus <> 13 And mFocus <> 49 And mFocus <> 40 And mFocus <> 31 And mFocus <> 22 Then
        DrawFocus False
        If mFocus = 4 Then
          mFocus = 14
        Else
          mFocus = mFocus + 1
        End If
        DrawFocus True
      End If
    Case vbKeyUp
      If Not Between(mFocus, 0, 4) And mFocus <> 14 Then
        DrawFocus False
        If Between(mFocus, 5, 9) Then
          mFocus = mFocus - 5
        ElseIf Between(mFocus, 41, 49) Then
          mFocus = mFocus - 36
        ElseIf Between(mFocus, 15, 22) Then
          mFocus = mFocus + 8
        ElseIf mFocus = 10 Or mFocus = 11 Then
          mFocus = 4
        ElseIf mFocus = 12 Or mFocus = 13 Then
          mFocus = 14
        Else
          mFocus = mFocus + 9
        End If
        DrawFocus True
      End If
    Case vbKeyDown
      If Not Between(mFocus, 15, 22) And (optColor(2).Visible Or Not Between(mFocus, 23, 31)) Then
        DrawFocus False
        If Between(mFocus, 0, 4) Then
          mFocus = mFocus + 5
        ElseIf Between(mFocus, 5, 13) Then
          mFocus = mFocus + 36
        ElseIf Between(mFocus, 23, 30) Then
          mFocus = mFocus - 8
        ElseIf mFocus = 14 Then
          mFocus = 13
        ElseIf mFocus = 31 Then
          mFocus = 22
        Else
          mFocus = mFocus - 9
        End If
        DrawFocus True
      End If
    Case vbKeyHome
      DrawFocus False
      mFocus = 0
      DrawFocus True
    Case vbKeyEnd
      DrawFocus False
      mFocus = 31
      DrawFocus True
    Case vbKeySpace
      lblColors_MouseDown (mFocus), 1, 0, 0, 0
  End Select
End Sub

Private Sub picColorFocus_LostFocus()
  DrawFocus False
End Sub

Private Sub picRGB_Click()
  Dim Color As Long
  Color = picRGB.BackColor
  If cDialog.ColorDialog(Color, True) Then
    DrawFocus False
    optColor(3).Value = True
    DrawFocus True
    picRGB.BackColor = Color
  End If
End Sub

Private Sub picRGB_GotFocus()
  mFocus = 255
  DrawFocus True
End Sub

Private Sub picRGB_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then picRGB_Click
End Sub

Private Sub picRGB_LostFocus()
  DrawFocus False
End Sub

Private Sub sldTransparency_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  optColor(3).Value = True
End Sub
