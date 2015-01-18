VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmCode 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   ClipControls    =   0   'False
   Icon            =   "Code.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   6120
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   5040
      TabIndex        =   24
      Tag             =   "1031"
      Top             =   6120
      WhatsThisHelpID =   1031
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   23
      Tag             =   "1030"
      Top             =   6120
      WhatsThisHelpID =   1030
      Width           =   1215
   End
   Begin VB.Frame TabFrame 
      BorderStyle     =   0  'None
      Height          =   5295
      Index           =   2
      Left            =   5040
      TabIndex        =   21
      Top             =   600
      Width           =   5775
      Begin VB.TextBox txtCode 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5055
         HideSelection   =   0   'False
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   22
         Top             =   120
         Width           =   5655
      End
   End
   Begin VB.Frame TabFrame 
      BorderStyle     =   0  'None
      Height          =   5295
      Index           =   1
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   5745
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   6
         Left            =   1800
         TabIndex        =   20
         Tag             =   "1704"
         Top             =   3510
         WhatsThisHelpID =   1702
         Width           =   1095
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   5
         Left            =   1800
         TabIndex        =   18
         Tag             =   "1703"
         Top             =   3120
         WhatsThisHelpID =   1702
         Width           =   1095
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   4
         Left            =   1800
         TabIndex        =   14
         Top             =   2310
         WhatsThisHelpID =   1045
         Width           =   1935
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   3
         Left            =   1800
         TabIndex        =   11
         Top             =   1920
         WhatsThisHelpID =   1045
         Width           =   1935
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   2
         Left            =   3120
         TabIndex        =   8
         Tag             =   "1044"
         Top             =   1200
         WhatsThisHelpID =   1042
         Width           =   1095
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   6
         Tag             =   "1043"
         Top             =   1200
         WhatsThisHelpID =   1042
         Width           =   1095
      End
      Begin VB.CheckBox chkLocked 
         Height          =   195
         Left            =   0
         TabIndex        =   3
         Tag             =   "1041"
         Top             =   600
         WhatsThisHelpID =   1041
         Width           =   2655
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   2
         Top             =   120
         WhatsThisHelpID =   1040
         Width           =   3135
      End
      Begin VB.Label lbls 
         Caption         =   "= %2"
         Height          =   255
         Index           =   8
         Left            =   3840
         TabIndex        =   15
         Top             =   2340
         Width           =   615
      End
      Begin VB.Label lbls 
         Caption         =   "= %1"
         Height          =   255
         Index           =   6
         Left            =   3840
         TabIndex        =   12
         Top             =   1950
         Width           =   615
      End
      Begin VB.Label lbls 
         Height          =   255
         Index           =   9
         Left            =   0
         TabIndex        =   16
         Tag             =   "1702"
         Top             =   2835
         WhatsThisHelpID =   1702
         Width           =   2655
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   11
         Left            =   360
         TabIndex        =   19
         Tag             =   "1704"
         Top             =   3540
         WhatsThisHelpID =   1702
         Width           =   1305
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   10
         Left            =   360
         TabIndex        =   17
         Tag             =   "1703"
         Top             =   3150
         WhatsThisHelpID =   1702
         Width           =   1305
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   7
         Left            =   600
         TabIndex        =   13
         Tag             =   "1047"
         Top             =   2340
         WhatsThisHelpID =   1045
         Width           =   1065
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   5
         Left            =   600
         TabIndex        =   10
         Tag             =   "1046"
         Top             =   1950
         WhatsThisHelpID =   1045
         Width           =   1065
      End
      Begin VB.Label lbls 
         Height          =   285
         Index           =   4
         Left            =   360
         TabIndex        =   9
         Tag             =   "1045"
         Top             =   1635
         WhatsThisHelpID =   1045
         Width           =   2175
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   3
         Left            =   2640
         TabIndex        =   7
         Tag             =   "1044"
         Top             =   1230
         WhatsThisHelpID =   1042
         Width           =   345
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   2
         Left            =   600
         TabIndex        =   5
         Tag             =   "1043"
         Top             =   1230
         WhatsThisHelpID =   1042
         Width           =   345
      End
      Begin VB.Label lbls 
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Tag             =   "1042"
         Top             =   915
         WhatsThisHelpID =   1042
         Width           =   2175
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Tag             =   "1040"
         Top             =   150
         WhatsThisHelpID =   1040
         Width           =   945
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   5895
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   10398
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "A"
            Object.Tag             =   "1700"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "B"
            Object.Tag             =   "1701"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mChanged As Boolean, mTxtChanged As Boolean, _
        mFocus As Integer

Private mValueCache(6) As Single

Public Function EditData(Data As clsCode) As Boolean
  Dim Offset As Single

  Load frmCode

  With Data
    Txts(0).Tag = .Caption(True)
    Txts(0).Text = .Name
    Txts_Change 0 ' (Changes the caption)
    Txts(1).Text = MeterToUser(.X)
    Txts(2).Text = MeterToUser(.Y)
    chkLocked.Value = -.Locked
    Txts_Validate 1, True ' (Updates Latitude, Longitude)

    Txts(5).Text = MeterToUser(.Horz)
    Txts(6).Text = MeterToUser(.Vert)
    
    txtCode.Text = .Text

    CenterForm Me
    Lang.PrepareForm Me
    mTxtChanged = False
    mChanged = False
    
    If TabValue > 1 Then TabValue = 0
    Set TabStrip1.SelectedItem = TabStrip1.Tabs(TabValue + 1)
    
    Show vbModal, Screen.ActiveForm

    If mChanged Then
      .Name = Txts(0).Text
      .X = mValueCache(1)
      .Y = mValueCache(2)
      .Locked = -chkLocked.Value
      .Horz = mValueCache(5)
      .Vert = mValueCache(6)
      
      .Text = txtCode.Text

      EditData = True
    End If
  End With

  Unload frmCode
End Function

Private Sub chkLocked_Click()
  Dim Value As Boolean
  Value = Not -chkLocked.Value
  SetEnabled Txts(1), Value
  SetEnabled Txts(2), Value
  SetEnabled Txts(3), Value
  SetEnabled Txts(4), Value
End Sub

Private Sub cmdCancel_Click()
  Hide
End Sub

Private Sub cmdOK_Click()
  Dim I As Integer, Msg As String, Cancel As Boolean

  ' Validate event not fired when Enter key pressed
  ' bug workaround
  If TypeOf ActiveControl Is TextBox Then
    Txts_Validate ActiveControl.Index, Cancel
    If Cancel Then Exit Sub
  End If

  For I = 1 To 6
    If Not Validate(Txts(I), Msg, mValueCache(I)) Then _
      GoTo ValidationError:
  Next I
  
  mChanged = True
  Hide
  Exit Sub
ValidationError:
  Set TabStrip1.SelectedItem = TabStrip1.Tabs(Chr$(Txts(I).Container.Index + 64))
  MsgBoxEx Me, Msg, vbCritical, RES_ERR_Bound
  Txts(I).SetFocus
  Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Handle Ctrl+ [Shift + ] {TAB}s
  If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
    Set TabStrip1.SelectedItem = TabStrip1.Tabs(TabStrip1.SelectedItem.Index Mod TabStrip1.Tabs.Count + 1)
  ElseIf Shift = (vbCtrlMask Or vbShiftMask) And KeyCode = vbKeyTab Then
    Set TabStrip1.SelectedItem = TabStrip1.Tabs((TabStrip1.SelectedItem.Index + TabStrip1.Tabs.Count - 2) Mod TabStrip1.Tabs.Count + 1)
  End If
End Sub

Private Sub Form_Load()
  Dim I As Integer
  
  CenterForm Me
  DialogMenus Me
  
  For I = 1 To TabFrame.Count
    With TabFrame(I)
      .Move TabFrame(1).Left, TabFrame(1).Top, TabFrame(1).Width, TabFrame(1).Height
      .Enabled = False
      .Visible = False
    End With
  Next I
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then Cancel = True: Hide
End Sub

Private Sub TabStrip1_BeforeClick(Cancel As Integer)
  Timer1.Enabled = True
End Sub

Private Sub TabStrip1_Click()
  ' Handle Tab clicks
  Dim SelItem As Integer
  SelItem = Asc(TabStrip1.SelectedItem.Key) - 64
  
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
  
  If TabStrip1.Tag <> "" And TabStrip1.Visible Then
    If Not ActiveControl Is TabStrip1 Then
      If TabStrip1.Tag <> SelItem Then
        Select Case SelItem
          Case 1: Txts(0).SetFocus
          Case 2: txtCode.SetFocus
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

Private Sub txtCode_GotFocus()
  cmdOK.Default = False
  cmdOK.TabStop = False
  cmdCancel.TabStop = False
  TabStrip1.TabStop = False
End Sub

Private Sub txtCode_LostFocus()
  cmdOK.Default = True
  cmdOK.TabStop = True
  cmdCancel.TabStop = True
  TabStrip1.TabStop = True
End Sub

Private Sub Txts_Change(Index As Integer)
  Dim TempStr As String
  If Index = 0 Then
    TempStr = Txts(0).Text
    If TempStr = "" Then
      Caption = Txts(0).Tag
    Else
      Caption = TempStr
    End If
  End If
  mTxtChanged = True
End Sub

Private Sub Txts_GotFocus(Index As Integer)
  Select Case Index
    Case 1, 2, 5, 6
      SmartSelectText Txts(Index)
    Case Else
      If Index = 0 Then ReturnSymbol 0, 0, 2
      SelectText Txts(Index)
  End Select
  mTxtChanged = False
End Sub

Private Sub Txts_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Dim Result As Integer
    
  If Index = 0 Then
    If Shift > 1 Then
      Result = ReturnSymbol(KeyCode, Shift)
      If Result > 0 Then Txts(Index).SelText = Chr$(Result): KeyCode = 0
    End If
  End If
End Sub

Private Sub Txts_KeyPress(Index As Integer, KeyAscii As Integer)
  If Index = 0 Then
    KeyAscii = ReturnSymbol(KeyAscii, 0, 1)
  End If
End Sub

Private Sub Txts_Validate(Index As Integer, Cancel As Boolean)
  Dim valX As Single, valY As Single, _
    Distance As Double, Angle As Single, _
    TempLatLon As clsLatLon

  On Error Resume Next
  
  If mTxtChanged = False Then Exit Sub
  mTxtChanged = False
  
  Select Case Index
    Case 1, 2
      If Validate(Txts(1), "", valX) And Validate(Txts(2), "", valY) And Txts(1).Text <> "" And Txts(2).Text <> "" Then
        Set TempLatLon = ReturnPoint(valX, valY)
        Txts(3).Text = TempLatLon.LatitudeUser
        Txts(4).Text = TempLatLon.LongitudeUser
      Else
        Txts(3).Text = ""
        Txts(4).Text = ""
      End If
    Case 3, 4
      Set TempLatLon = New clsLatLon
      TempLatLon.Latitude = Txts(3).Text
      TempLatLon.Longitude = Txts(4).Text
      If TempLatLon.Validate("") Then
        Scenery.Header.Center.CalcDistance TempLatLon, Distance, Angle
        PolarToRect Distance * NmToM, 90 - Angle, valX, valY
        Txts(1).Text = MeterToUser(Round(valX, 2), "0.00")
        Txts(2).Text = MeterToUser(Round(valY, 2), "0.00")
      Else
        Txts(1).Text = ""
        Txts(2).Text = ""
      End If
      Set TempLatLon = Nothing
  End Select
End Sub
