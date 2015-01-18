VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmTower 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5085
   ClipControls    =   0   'False
   Icon            =   "Tower.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame TabFrame 
      BorderStyle     =   0  'None
      Height          =   4815
      Index           =   2
      Left            =   360
      TabIndex        =   16
      Top             =   600
      Width           =   4335
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   17
         Left            =   2640
         TabIndex        =   40
         Tag             =   "1862"
         Top             =   4410
         WhatsThisHelpID =   1851
         Width           =   1335
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   16
         Left            =   2640
         TabIndex        =   38
         Tag             =   "1861"
         Top             =   4020
         WhatsThisHelpID =   1851
         Width           =   1335
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   15
         Left            =   2640
         TabIndex        =   36
         Tag             =   "1860"
         Top             =   3630
         WhatsThisHelpID =   1851
         Width           =   1335
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   14
         Left            =   2640
         TabIndex        =   34
         Tag             =   "1859"
         Top             =   3240
         WhatsThisHelpID =   1851
         Width           =   1335
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   13
         Left            =   2640
         TabIndex        =   32
         Tag             =   "1858"
         Top             =   2850
         WhatsThisHelpID =   1851
         Width           =   1335
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   12
         Left            =   2640
         TabIndex        =   30
         Tag             =   "1857"
         Top             =   2460
         WhatsThisHelpID =   1851
         Width           =   1335
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   11
         Left            =   2640
         TabIndex        =   28
         Tag             =   "1856"
         Top             =   2070
         WhatsThisHelpID =   1851
         Width           =   1335
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   10
         Left            =   2640
         TabIndex        =   26
         Tag             =   "1855"
         Top             =   1680
         WhatsThisHelpID =   1851
         Width           =   1335
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   9
         Left            =   2640
         TabIndex        =   24
         Tag             =   "1854"
         Top             =   1290
         WhatsThisHelpID =   1851
         Width           =   1335
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   8
         Left            =   2640
         TabIndex        =   22
         Tag             =   "1853"
         Top             =   900
         WhatsThisHelpID =   1851
         Width           =   1335
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   7
         Left            =   2640
         TabIndex        =   20
         Tag             =   "1852"
         Top             =   510
         WhatsThisHelpID =   1851
         Width           =   1335
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   6
         Left            =   2640
         TabIndex        =   18
         Tag             =   "1851"
         Top             =   120
         WhatsThisHelpID =   1851
         Width           =   1335
      End
      Begin VB.CheckBox chkFrequency 
         Height          =   195
         Index           =   11
         Left            =   0
         TabIndex        =   39
         Tag             =   "1862"
         Top             =   4440
         WhatsThisHelpID =   1851
         Width           =   2505
      End
      Begin VB.CheckBox chkFrequency 
         Height          =   195
         Index           =   10
         Left            =   0
         TabIndex        =   37
         Tag             =   "1861"
         Top             =   4050
         WhatsThisHelpID =   1851
         Width           =   2505
      End
      Begin VB.CheckBox chkFrequency 
         Height          =   195
         Index           =   9
         Left            =   0
         TabIndex        =   35
         Tag             =   "1860"
         Top             =   3660
         WhatsThisHelpID =   1851
         Width           =   2505
      End
      Begin VB.CheckBox chkFrequency 
         Height          =   195
         Index           =   8
         Left            =   0
         TabIndex        =   33
         Tag             =   "1859"
         Top             =   3270
         WhatsThisHelpID =   1851
         Width           =   2505
      End
      Begin VB.CheckBox chkFrequency 
         Height          =   195
         Index           =   7
         Left            =   0
         TabIndex        =   31
         Tag             =   "1858"
         Top             =   2880
         WhatsThisHelpID =   1851
         Width           =   2505
      End
      Begin VB.CheckBox chkFrequency 
         Height          =   195
         Index           =   6
         Left            =   0
         TabIndex        =   29
         Tag             =   "1857"
         Top             =   2490
         WhatsThisHelpID =   1851
         Width           =   2505
      End
      Begin VB.CheckBox chkFrequency 
         Height          =   195
         Index           =   5
         Left            =   0
         TabIndex        =   27
         Tag             =   "1856"
         Top             =   2100
         WhatsThisHelpID =   1851
         Width           =   2505
      End
      Begin VB.CheckBox chkFrequency 
         Height          =   195
         Index           =   4
         Left            =   0
         TabIndex        =   25
         Tag             =   "1855"
         Top             =   1710
         WhatsThisHelpID =   1851
         Width           =   2505
      End
      Begin VB.CheckBox chkFrequency 
         Height          =   195
         Index           =   3
         Left            =   0
         TabIndex        =   23
         Tag             =   "1854"
         Top             =   1320
         WhatsThisHelpID =   1851
         Width           =   2505
      End
      Begin VB.CheckBox chkFrequency 
         Height          =   195
         Index           =   2
         Left            =   0
         TabIndex        =   21
         Tag             =   "1853"
         Top             =   930
         WhatsThisHelpID =   1851
         Width           =   2505
      End
      Begin VB.CheckBox chkFrequency 
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   19
         Tag             =   "1852"
         Top             =   540
         WhatsThisHelpID =   1851
         Width           =   2505
      End
      Begin VB.CheckBox chkFrequency 
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   17
         Tag             =   "1851"
         Top             =   150
         WhatsThisHelpID =   1851
         Width           =   2505
      End
   End
   Begin VB.Frame TabFrame 
      BorderStyle     =   0  'None
      Height          =   4815
      Index           =   1
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   4440
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   5
         Left            =   1320
         TabIndex        =   15
         Tag             =   "1850"
         Top             =   2700
         WhatsThisHelpID =   1850
         Width           =   1095
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   4
         Left            =   1800
         TabIndex        =   13
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
         Height          =   390
         Index           =   7
         Left            =   0
         TabIndex        =   14
         Tag             =   "1850"
         Top             =   2730
         WhatsThisHelpID =   1850
         Width           =   1185
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   6
         Left            =   600
         TabIndex        =   12
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
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   5640
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   41
      Tag             =   "1031"
      Top             =   5640
      WhatsThisHelpID =   1031
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   42
      Tag             =   "1030"
      Top             =   5640
      WhatsThisHelpID =   1030
      Width           =   1215
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   5415
      Left            =   120
      TabIndex        =   43
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   9551
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "A"
            Object.Tag             =   "1870"
            Object.ToolTipText     =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "B"
            Object.Tag             =   "1871"
            Object.ToolTipText     =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTower"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mChanged As Boolean, mTxtChanged As Boolean, mDontUpdate As Boolean

Private mValueCache(17) As Single

Public Function EditData(Data As clsTower) As Boolean
  Dim I As Integer
  Load frmTower
  
  With Data
    Txts(0).Tag = .Caption(True)
    Txts(0).Text = .Name
    Txts_Change 0 ' (Changes the caption)
    Txts(1).Text = MeterToUser(.X)
    Txts(2).Text = MeterToUser(.Y)
    Txts_Validate 1, False ' (Updates Latitude, Longitude)
    chkLocked.Value = -.Locked
    Txts(5).Text = MeterToUser(.Height)
    
    mDontUpdate = True
    For I = 0 To 11
      chkFrequency(I).Value = IIf(.COMFrequency(I) > 0, vbChecked, vbUnchecked)
      If .COMFrequency(I) > 0 Then
        Txts(I + 6).Text = Append(.COMFrequency(I), RES_Unit_AbbrevMhz, "000.000")
      End If
      chkFrequency_Click I
    Next I
    mDontUpdate = False
    
    Lang.PrepareForm Me
    TabStrip1_Click
    mTxtChanged = False
    mChanged = False
    If TabValue >= TabStrip1.Tabs.Count Then TabValue = 0
    Set TabStrip1.SelectedItem = TabStrip1.Tabs(TabValue + 1)
    Show vbModal, Screen.ActiveForm
    
    If mChanged Then
      .Name = Txts(0).Text
      .X = mValueCache(1)
      .Y = mValueCache(2)
      .Locked = -chkLocked.Value
      .Height = mValueCache(5)
      
      For I = 0 To 11
        .COMFrequency(I) = mValueCache(I + 6) * chkFrequency(I).Value
      Next I
      
      Scenery.AFDRefresh = True
      
      EditData = True
    End If
  End With

  Unload frmTower
End Function

Private Sub chkFrequency_Click(Index As Integer)
  Dim Value As Integer
  Value = chkFrequency(Index).Value
  SetEnabled Txts(Index + 6), Value = vbChecked
  If Not mDontUpdate And Value Then
    Txts(Index + 6).Text = Append(127.3 + Index * 0.1, RES_Unit_AbbrevMhz, "000.000")
  End If
End Sub

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

  For I = 0 To 17
    If Not Txts(I).Locked Then _
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
          Case 2: chkFrequency(0).SetFocus
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

Private Sub Txts_Change(Index As Integer)
  Dim TempStr As String, Conv As Single
  Select Case Index
    Case 0
      TempStr = Txts(0).Text
      If TempStr = "" Then
        Caption = Txts(0).Tag
      Else
        Caption = TempStr
      End If
  End Select
  mTxtChanged = True
End Sub

Private Sub Txts_GotFocus(Index As Integer)
  Select Case Index
    Case 1, 2, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17
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
