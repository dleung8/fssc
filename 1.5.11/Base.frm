VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmBuilding 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   ClipControls    =   0   'False
   Icon            =   "Building.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   4200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Tag             =   "1031"
      Top             =   4200
      WhatsThisHelpID =   1031
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Tag             =   "1030"
      Top             =   4200
      WhatsThisHelpID =   1030
      Width           =   1215
   End
   Begin VB.Frame TabFrame 
      BorderStyle     =   0  'None
      Height          =   3375
      Index           =   2
      Left            =   360
      TabIndex        =   5
      Top             =   600
      Width           =   4815
   End
   Begin VB.Frame TabFrame 
      BorderStyle     =   0  'None
      Height          =   3375
      Index           =   1
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   4815
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   1
         Left            =   1440
         MaxLength       =   64
         TabIndex        =   4
         Top             =   510
         Width           =   2775
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   2
         Top             =   120
         Width           =   2775
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   1
         Left            =   0
         TabIndex        =   3
         Top             =   540
         Width           =   1305
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   150
         Width           =   1305
      End
      Begin VB.Image btnSymbols 
         Height          =   285
         Index           =   1
         Left            =   4320
         Top             =   510
         Width           =   285
      End
      Begin VB.Image btnSymbols 
         Height          =   285
         Index           =   0
         Left            =   4320
         Top             =   120
         Width           =   285
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   3975
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   7011
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Object.Tag             =   "1120"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Object.Tag             =   "1121"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmBuilding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
#If SUBCLASS Then
Implements ISubclass
#End If

Private mChanged As Boolean

Private ValueCache(14) As Single

Public Function EditData(Data As clsBuilding) As Boolean
  Load frmBuilding
  With Data
    Txts(0).Text = .Name
    Changed = False
    Show vbModal, Screen.ActiveForm
    If Changed Then
      EditData = True
    End If
  End With
  Unload frmBuilding
End Function

Private Sub cmdCancel_Click()
  Hide
End Sub

Private Sub cmdOK_Click()
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Handle Ctrl+ [Shift + ] {TAB}s
  If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
    Set TabStrip1.SelectedItem = TabStrip1.Tabs(TabStrip1.SelectedItem.Index Mod TabStrip1.Tabs.Count + 1)
  ElseIf Shift = (vbCtrlMask Or vbShiftMask) And KeyCode = vbKeyTab Then
    Set TabStrip1.SelectedItem = TabStrip1.Tabs((TabStrip1.SelectedItem.Index - 2 + TabStrip1.Tabs.Count) Mod TabStrip1.Tabs.Count + 1)
  End If
End Sub

Private Sub Form_Load()
  Dim I As Integer
  
  CenterForm Me
  If TutorialVisible Then Left = 100
  DialogMenus Me
  Lang.PrepareForm Me
  #If SUBCLASS Then
  AttachMessage Me, Me.hWnd, WM_HELP
  #End If

  For I = 1 To TabFrame.Count
    With TabFrame(I)
      .Move TabFrame(1).Left, TabFrame(1).Top, TabFrame(1).Width, TabFrame(1).Height
      .Enabled = False
      .Visible = False
    End With
  Next I
  
  Set TabStrip1.SelectedItem = TabStrip1.Tabs(2)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then Cancel = True: Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
  #If SUBCLASS Then
  DetachMessage Me, Me.hWnd, WM_HELP
  #End If
End Sub

#If SUBCLASS Then
Private Property Get ISubclass_MsgResponse() As Fsscsubc.EMsgResponse
  ISubclass_MsgResponse = emrConsume
End Property
#End If

#If SUBCLASS Then
Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  'If iMsg = WM_HELP Then
  If MsgBoxHelpID > 0 Then
    ' Help button in Messagebox
    DoErrorHelp Me
  Else
    ' What's This Help
    CallOldWindowProc hWnd, iMsg, wParam, lParam
  End If
  'End If
End Function
#End If

Private Sub TabStrip1_BeforeClick(Cancel As Integer)
  Timer1.Enabled = True
End Sub

Private Sub TabStrip1_Click()
  ' Handle Tab clicks
  Dim SelItem As Integer
  SelItem = TabStrip1.SelectedItem.Index
  
  If TabStrip1.Tag <> "" Then
    If SelItem <> TabStrip1.Tag Then
      With TabFrame(TabStrip1.Tag)
        .Visible = False
        .Enabled = False
      End With
    End If
  End If
  
  With TabFrame(SelItem)
    .Enabled = True
    .Visible = True
  End With
  
  If TabStrip1.Tag <> "" Then
    If Not ActiveControl Is TabStrip1 Then
      If TabStrip1.Tag <> SelItem Then
        Select Case SelItem
          Case 1: Txts(0).SetFocus
'          Case 2: Txts(2).SetFocus
'          Case 3: Txts(8).SetFocus
'          Case 4: chkExclusion(0).SetFocus
'          Case 5: cmbSize.SetFocus
        End Select
      End If
    End If
  End If
  
  TabStrip1.Tag = SelItem
End Sub

Private Sub Timer1_Timer()
  TabStrip1_Click
  Timer1.Enabled = False
End Sub

Private Sub Txts_GotFocus(Index As Integer)
  If Index = 7 Or Between(Index, 10, 13) Then
    SmartSelectText Txts(Index)
  Else
    SelectText Txts(Index)
  End If
  If Index = 0 Then ReturnSymbol 0, 0, 2
End Sub

Private Sub Txts_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Dim Result As Byte
    
  If Index = 0 Then
    If Shift > 1 Then
      Result = ReturnSymbol(KeyCode, Shift)
      If Result > 0 Then Txts(Index).SelText = Chr$(Result): KeyCode = 0
    End If
  End If
End Sub

Private Sub Txts_KeyPress(Index As Integer, KeyAscii As Integer)
  If Index = 0 Then
    ' Suppress "Ding"
    If KeyAscii = 19 Then KeyAscii = 0 ' Ctrl+S
    KeyAscii = ReturnSymbol(KeyAscii, 0, 1)
End Sub

