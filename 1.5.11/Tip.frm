VERSION 5.00
Begin VB.Form frmTip 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3630
   ClientLeft      =   2355
   ClientTop       =   2385
   ClientWidth     =   5655
   Icon            =   "Tip.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   Tag             =   "2280"
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CheckBox chkShow 
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Tag             =   "2282"
      Top             =   3120
      Width           =   3735
   End
   Begin VB.CommandButton cmdNextTip 
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Tag             =   "2283"
      Top             =   720
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2715
      Left            =   240
      ScaleHeight     =   2655
      ScaleWidth      =   3675
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   3735
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "Tip.frx":000C
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblKnow 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   720
         TabIndex        =   1
         Tag             =   "2281"
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         Height          =   1635
         Left            =   180
         TabIndex        =   2
         Top             =   720
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Tag             =   "1030"
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkShow_Click()
  Options.ShowTips = -chkShow.Value
End Sub

Private Sub cmdNextTip_Click()
  Dim Temp As String
  
  Temp = Lang.GetString(RES_TIP_FIRST + Options.CurrentTip)
  If Temp = "#" Then
    Options.CurrentTip = 0
    Temp = Lang.GetString(RES_TIP_FIRST + Options.CurrentTip)
  End If
  lblTipText.Caption = Temp
  Options.CurrentTip = Options.CurrentTip + 1
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
  CenterForm Me
  DialogMenus Me
  Lang.PrepareForm Me
  chkShow.Value = -Options.ShowTips
  cmdNextTip_Click
End Sub
