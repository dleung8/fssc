VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   3990
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   6750
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Splash.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picLine 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   75
      Left            =   540
      ScaleHeight     =   75
      ScaleMode       =   0  'User
      ScaleWidth      =   300
      TabIndex        =   0
      Top             =   1800
      Width           =   5670
   End
   Begin VB.Image imgLogo 
      Height          =   1245
      Left            =   788
      Top             =   360
      Width           =   5175
   End
   Begin VB.Label lblContact 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   3240
      Width           =   5535
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   2640
      Width           =   3135
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   2280
      Width           =   3135
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' [frmSplash]
' Shows a quick About window and progress bar while loading
Option Explicit

' Always on top
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_SHOWWINDOW = &H40

Public isShown As Boolean

Private mLow As Integer, mHigh As Integer

' Change the percent indicator
Public Sub LoadPercent(ByVal PercentLoaded As Integer)
  mHigh = 3 * PercentLoaded
  picLine_Paint
  mLow = 0
  DoEvents
End Sub

Private Sub Form_Activate()
  Refresh
End Sub

Private Sub Form_Load()
  Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
  SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
  
  lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
  lblCopyright.Caption = "Copyright © " & CopyrightYear & " by Derek Leung" & vbNewLine & "Leung Software"
  lblContact.Caption = Email & vbCrLf & Webpage

  On Error Resume Next
  imgLogo.Picture = LoadResPicture(1000, vbResBitmap)
  isShown = True
End Sub

Private Sub Form_Paint()
  Line (0, 0)-(ScaleWidth - 15, ScaleHeight - 15), vbBlack, B
  Line (15, 15)-(ScaleWidth - 30, ScaleHeight - 30), vbBlack, B
End Sub

Private Sub Form_Unload(Cancel As Integer)
  isShown = False
End Sub

Private Sub picLine_Paint()
  Dim I As Integer, _
    ColorShift As Integer, _
    Temp As Integer
  
  If mLow > 270 Then
    ColorShift = 192
  ElseIf mLow > 79 Then
    ColorShift = mLow - 79
  End If
    
  For I = mLow To mHigh
    If Between(I, 79, 270) Then ColorShift = ColorShift + 1
    Temp = 192 + ColorShift / 3
    picLine.Line (I, 0)-(I + 1, picLine.ScaleHeight), RGB(192 - ColorShift, Temp, Temp), BF
  Next I
End Sub
