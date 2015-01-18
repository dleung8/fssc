VERSION 5.00
Begin VB.Form frmConvert 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6645
   ClipControls    =   0   'False
   Icon            =   "Convert.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Tag             =   "1030"
      Top             =   2160
      Width           =   1455
   End
   Begin VB.OptionButton optSelection 
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   2
      Tag             =   "2003"
      Top             =   1560
      Width           =   5415
   End
   Begin VB.OptionButton optSelection 
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   1
      Tag             =   "2002"
      Top             =   1080
      Value           =   -1  'True
      Width           =   5415
   End
   Begin VB.Label lbls 
      Height          =   735
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Tag             =   "2001"
      Top             =   240
      Width           =   5415
   End
End
Attribute VB_Name = "frmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long

' Shows the Conversion prompt, and returns 1 if a conversion is desired
Public Function DoDialog() As Integer
  Dim OldValue As MousePointerConstants
  OldValue = Screen.MousePointer
  Screen.MousePointer = vbDefault
  Load frmConvert
  MessageBeep vbQuestion
  Show vbModal, Screen.ActiveForm
  DoDialog = -optSelection(1).Value
  Unload frmConvert
  Screen.MousePointer = OldValue
End Function

Private Sub cmdOK_Click()
  Hide
End Sub

Private Sub Form_Load()
  Caption = App.Title
  CenterForm Me
  DialogMenus Me, True
  Lang.PrepareForm Me
End Sub

Private Sub Form_Paint()
  DrawQuestionIcon hdc, 10, lbls(0).Top / Screen.TwipsPerPixelY
End Sub

