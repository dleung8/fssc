VERSION 5.00
Begin VB.Form frmZoom 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4110
   ClipControls    =   0   'False
   Icon            =   "Zoom.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   Tag             =   "2180"
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Tag             =   "1031"
      Top             =   720
      WhatsThisHelpID =   1031
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Tag             =   "1030"
      Top             =   240
      WhatsThisHelpID =   1030
      Width           =   1215
   End
   Begin VB.TextBox Txts 
      Height          =   285
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Tag             =   "2180"
      Top             =   360
      WhatsThisHelpID =   2180
      Width           =   1215
   End
   Begin VB.Label lbls 
      Height          =   390
      Index           =   0
      Left            =   1680
      TabIndex        =   1
      Tag             =   "1090"
      Top             =   390
      WhatsThisHelpID =   2180
      Width           =   705
   End
End
Attribute VB_Name = "frmZoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' [frmZoom]
' Allows the user to select a Zoom value
Option Explicit

Private mChanged As Boolean

Private mValueCache(0) As Single

' Prompts the user for a Zoom
' NoCancel - If true, the user cannot select the
'               cancel button
Public Function EditData(ZoomLevel As Single)
  Load frmZoom
  Txts(0).Text = MeterToUser(ZoomLevel, "##0.00")
  mChanged = False
  Show vbModal, Screen.ActiveForm
  If mChanged Then
    ZoomLevel = mValueCache(0)
    EditData = True
  End If
  Unload frmZoom
End Function

Private Sub cmdCancel_Click()
  Hide
End Sub

Private Sub cmdOK_Click()
  Dim Msg As String

  If Not Validate(Txts(0), Msg, mValueCache(0)) Then
    MsgBoxEx Me, Msg, vbCritical, RES_ERR_Bound
    Txts(0).SetFocus
  Else
    mChanged = True
    Hide
  End If
End Sub

Private Sub Form_Load()
  CenterForm Me
  DialogMenus Me
  Lang.PrepareForm Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then Cancel = True: Hide
End Sub

Private Sub lstZoom_DblClick()
  cmdOK_Click
End Sub

Private Sub Txts_GotFocus(Index As Integer)
  SmartSelectText Txts(Index)
End Sub
