VERSION 5.00
Begin VB.Form frmTransform 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   Icon            =   "Transform.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   Tag             =   "2190"
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Tag             =   "1030"
      Top             =   240
      WhatsThisHelpID =   1030
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Tag             =   "1031"
      Top             =   720
      WhatsThisHelpID =   1031
      Width           =   1215
   End
   Begin VB.TextBox Txts 
      Height          =   285
      Index           =   2
      Left            =   1680
      TabIndex        =   5
      Tag             =   "2193"
      Top             =   1740
      WhatsThisHelpID =   2193
      Width           =   1095
   End
   Begin VB.TextBox Txts 
      Height          =   285
      Index           =   1
      Left            =   1680
      TabIndex        =   3
      Tag             =   "1044"
      Top             =   1350
      WhatsThisHelpID =   2192
      Width           =   1095
   End
   Begin VB.TextBox Txts 
      Height          =   285
      Index           =   0
      Left            =   1680
      TabIndex        =   1
      Tag             =   "1043"
      Top             =   960
      WhatsThisHelpID =   2192
      Width           =   1095
   End
   Begin VB.OptionButton opts 
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Tag             =   "2192"
      Top             =   480
      WhatsThisHelpID =   2191
      Width           =   2415
   End
   Begin VB.OptionButton opts 
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Tag             =   "2191"
      Top             =   240
      Value           =   -1  'True
      WhatsThisHelpID =   2191
      Width           =   2415
   End
   Begin VB.Label lbls 
      Height          =   390
      Index           =   1
      Left            =   480
      TabIndex        =   4
      Tag             =   "2193"
      Top             =   1770
      WhatsThisHelpID =   2193
      Width           =   1185
   End
   Begin VB.Label lbls 
      Height          =   390
      Index           =   0
      Left            =   480
      TabIndex        =   2
      Tag             =   "1044"
      Top             =   1365
      WhatsThisHelpID =   2192
      Width           =   1185
   End
   Begin VB.Label lbls 
      Height          =   390
      Index           =   9
      Left            =   480
      TabIndex        =   0
      Tag             =   "1043"
      Top             =   990
      WhatsThisHelpID =   2192
      Width           =   1185
   End
End
Attribute VB_Name = "frmTransform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' [frmTransform]
' Allows the user to transform the scenery
Option Explicit

Private mChanged As Boolean

Private mValueCache(2) As Single

' Prompts the user for a transformation
' NoCancel - If true, the user cannot select the
'               cancel button
Public Function EditData(ByRef X As Single, ByRef Y As Single, ByRef Rotation As Single, ByRef All As Boolean)
  Load frmTransform
  mChanged = False
  Txts(0).Text = MeterToUser(0)
  Txts(1).Text = MeterToUser(0)
  Txts(2).Text = Append(0, RES_Unit_AbbrevDeg)
  Show vbModal, Screen.ActiveForm
  If mChanged Then
    X = mValueCache(0)
    Y = mValueCache(1)
    Rotation = mValueCache(2)
    All = opts(0).Value
    EditData = True
  End If
  Unload frmTransform
End Function

Private Sub cmdCancel_Click()
  Hide
End Sub

Private Sub cmdOK_Click()
  Dim I As Integer, Msg As String

  For I = 0 To 2
    If Not Validate(Txts(I), Msg, mValueCache(I)) Then _
      GoTo ValidationError:
  Next I
  mChanged = True
  Hide
  Exit Sub
ValidationError:
  MsgBoxEx Me, Msg, vbCritical, RES_ERR_Bound
  Txts(I).SetFocus
  Exit Sub
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

