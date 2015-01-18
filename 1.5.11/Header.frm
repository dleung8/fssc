VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmHeader 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   ClipControls    =   0   'False
   Icon            =   "Header.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   Tag             =   "1140"
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame TabFrame 
      BorderStyle     =   0  'None
      Height          =   3300
      Index           =   5
      Left            =   360
      TabIndex        =   40
      Top             =   600
      Width           =   4815
      Begin VB.ComboBox cmbSynthetic 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   2745
         WhatsThisHelpID =   1115
         Width           =   2775
      End
      Begin VB.ComboBox cmbSize 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   105
         WhatsThisHelpID =   1114
         Width           =   2775
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   18
         Left            =   0
         TabIndex        =   44
         Tag             =   "1115"
         Top             =   2790
         WhatsThisHelpID =   1115
         Width           =   1305
      End
      Begin VB.Label lbls 
         Height          =   735
         Index           =   17
         Left            =   480
         TabIndex        =   43
         Tag             =   "3023"
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Image imgTexture 
         BorderStyle     =   1  'Fixed Single
         Height          =   2175
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   495
         Width           =   2175
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   16
         Left            =   0
         TabIndex        =   41
         Tag             =   "1114"
         Top             =   150
         WhatsThisHelpID =   1114
         Width           =   1305
      End
   End
   Begin VB.Frame TabFrame 
      BorderStyle     =   0  'None
      Height          =   1695
      Index           =   4
      Left            =   360
      TabIndex        =   35
      Top             =   600
      Width           =   3735
      Begin VB.CheckBox chkExclusion 
         Height          =   195
         Index           =   3
         Left            =   0
         TabIndex        =   39
         Tag             =   "1813"
         Top             =   1200
         WhatsThisHelpID =   1810
         Width           =   3600
      End
      Begin VB.CheckBox chkExclusion 
         Height          =   195
         Index           =   2
         Left            =   0
         TabIndex        =   38
         Tag             =   "1812"
         Top             =   840
         WhatsThisHelpID =   1810
         Width           =   3600
      End
      Begin VB.CheckBox chkExclusion 
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   37
         Tag             =   "1811"
         Top             =   480
         WhatsThisHelpID =   1810
         Width           =   3600
      End
      Begin VB.CheckBox chkExclusion 
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   36
         Tag             =   "1810"
         Top             =   120
         WhatsThisHelpID =   1810
         Width           =   3600
      End
   End
   Begin VB.Frame TabFrame 
      BorderStyle     =   0  'None
      Height          =   3300
      Index           =   3
      Left            =   360
      TabIndex        =   20
      Top             =   600
      Width           =   4815
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   13
         Left            =   1440
         TabIndex        =   34
         Tag             =   "1113"
         Top             =   2820
         WhatsThisHelpID =   1113
         Width           =   1095
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   12
         Left            =   1440
         TabIndex        =   32
         Tag             =   "1112"
         Top             =   2310
         WhatsThisHelpID =   1112
         Width           =   1095
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   11
         Left            =   1440
         TabIndex        =   30
         Tag             =   "1111"
         Top             =   1920
         WhatsThisHelpID =   1110
         Width           =   1095
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   10
         Left            =   1440
         TabIndex        =   28
         Tag             =   "1110"
         Top             =   1530
         WhatsThisHelpID =   1110
         Width           =   1095
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   8
         Left            =   1440
         TabIndex        =   24
         Top             =   510
         WhatsThisHelpID =   1120
         Width           =   1935
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   7
         Left            =   1440
         TabIndex        =   22
         Top             =   120
         WhatsThisHelpID =   1120
         Width           =   1935
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   9
         Left            =   1440
         TabIndex        =   26
         Tag             =   "1116"
         Top             =   900
         WhatsThisHelpID =   1116
         Width           =   1095
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   15
         Left            =   0
         TabIndex        =   33
         Tag             =   "1113"
         Top             =   2850
         WhatsThisHelpID =   1113
         Width           =   1305
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   14
         Left            =   0
         TabIndex        =   31
         Tag             =   "1112"
         Top             =   2340
         WhatsThisHelpID =   1112
         Width           =   1305
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   13
         Left            =   0
         TabIndex        =   29
         Tag             =   "1111"
         Top             =   1950
         WhatsThisHelpID =   1110
         Width           =   1305
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   12
         Left            =   0
         TabIndex        =   27
         Tag             =   "1110"
         Top             =   1560
         WhatsThisHelpID =   1110
         Width           =   1305
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   10
         Left            =   0
         TabIndex        =   23
         Tag             =   "1047"
         Top             =   540
         WhatsThisHelpID =   1120
         Width           =   1305
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   9
         Left            =   0
         TabIndex        =   21
         Tag             =   "1046"
         Top             =   150
         WhatsThisHelpID =   1120
         Width           =   1305
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   11
         Left            =   0
         TabIndex        =   25
         Tag             =   "1116"
         Top             =   930
         WhatsThisHelpID =   1116
         Width           =   1305
      End
   End
   Begin VB.Frame TabFrame 
      BorderStyle     =   0  'None
      Height          =   3300
      Index           =   2
      Left            =   360
      TabIndex        =   5
      Top             =   600
      Width           =   4815
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   6
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   19
         Top             =   2460
         WhatsThisHelpID =   1104
         Width           =   1335
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   5
         Left            =   1440
         TabIndex        =   17
         Top             =   2070
         WhatsThisHelpID =   1104
         Width           =   2775
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   4
         Left            =   1440
         TabIndex        =   15
         Top             =   1680
         WhatsThisHelpID =   1104
         Width           =   2775
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   3
         Left            =   1440
         TabIndex        =   13
         Top             =   1290
         WhatsThisHelpID =   1104
         Width           =   2775
      End
      Begin VB.ComboBox cmbRegion 
         Height          =   315
         Left            =   1440
         TabIndex        =   11
         Top             =   885
         WhatsThisHelpID =   1104
         Width           =   2775
      End
      Begin VB.ComboBox cmbLangCode 
         Height          =   315
         Left            =   1440
         TabIndex        =   9
         Top             =   495
         WhatsThisHelpID =   1103
         Width           =   2775
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   2
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   7
         Top             =   120
         WhatsThisHelpID =   1102
         Width           =   2775
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   8
         Left            =   0
         TabIndex        =   18
         Tag             =   "1108"
         Top             =   2490
         WhatsThisHelpID =   1104
         Width           =   1305
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   7
         Left            =   0
         TabIndex        =   16
         Tag             =   "1107"
         Top             =   2100
         WhatsThisHelpID =   1104
         Width           =   1305
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   6
         Left            =   0
         TabIndex        =   14
         Tag             =   "1106"
         Top             =   1710
         WhatsThisHelpID =   1104
         Width           =   1305
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   5
         Left            =   0
         TabIndex        =   12
         Tag             =   "1105"
         Top             =   1320
         WhatsThisHelpID =   1104
         Width           =   1305
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   4
         Left            =   0
         TabIndex        =   10
         Tag             =   "1104"
         Top             =   930
         WhatsThisHelpID =   1104
         Width           =   1305
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   3
         Left            =   0
         TabIndex        =   8
         Tag             =   "1103"
         Top             =   540
         WhatsThisHelpID =   1103
         Width           =   1305
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   2
         Left            =   0
         TabIndex        =   6
         Tag             =   "1102"
         Top             =   150
         WhatsThisHelpID =   1102
         Width           =   1305
      End
      Begin VB.Image btnSymbols 
         Height          =   285
         Index           =   5
         Left            =   4320
         Top             =   2070
         Width           =   285
      End
      Begin VB.Image btnSymbols 
         Height          =   285
         Index           =   4
         Left            =   4320
         Top             =   1680
         Width           =   285
      End
      Begin VB.Image btnSymbols 
         Height          =   285
         Index           =   3
         Left            =   4320
         Top             =   1290
         Width           =   285
      End
      Begin VB.Image btnSymbols 
         Height          =   285
         Index           =   2
         Left            =   4320
         Top             =   120
         Width           =   285
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   4080
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   47
      Tag             =   "1031"
      Top             =   4080
      WhatsThisHelpID =   1031
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   46
      Tag             =   "1030"
      Top             =   4080
      WhatsThisHelpID =   1030
      Width           =   1215
   End
   Begin VB.Frame TabFrame 
      BorderStyle     =   0  'None
      Height          =   3300
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
         WhatsThisHelpID =   1101
         Width           =   2775
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   2
         Top             =   120
         WhatsThisHelpID =   1100
         Width           =   2775
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   1
         Left            =   0
         TabIndex        =   3
         Tag             =   "1101"
         Top             =   540
         WhatsThisHelpID =   1101
         Width           =   1305
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Tag             =   "1100"
         Top             =   150
         WhatsThisHelpID =   1100
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
      Height          =   3855
      Left            =   120
      TabIndex        =   48
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   6800
      MultiRow        =   -1  'True
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   5
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "A"
            Object.Tag             =   "1120"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "B"
            Object.Tag             =   "1121"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "C"
            Object.Tag             =   "1122"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "D"
            Object.Tag             =   "1123"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "E"
            Object.Tag             =   "1124"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mChanged As Boolean, _
        mDontSelect As Boolean

Private mValueCache(3 To 13) As Single, _
        LatLonCache(1) As Double

Public Function EditData(Data As clsHeader) As Boolean
  Load frmHeader
  With Data
    Txts(0).Text = .Author
    Txts(1).Text = .Copyright
    Txts(2).Text = .Name
    SetComboText cmbLangCode, .LangCode
    SetComboText cmbRegion, .Region
    Txts(3).Text = .Country
    Txts(4).Text = .State
    Txts(5).Text = .City
    Txts(6).Text = .ICAOID
    Txts(7).Text = .Center.LatitudeUser
    Txts(8).Text = .Center.LongitudeUser
    Txts(9).Text = Append(.Rotation, RES_Unit_AbbrevDeg)
    UpdateFields
    Txts(10).Text = MeterToUser(.Horz)
    Txts(11).Text = MeterToUser(.Vert)
    Txts(12).Text = MeterToUser(.Altitude)
    Txts(13).Text = Append(.MagVar, RES_Unit_AbbrevDeg)
    chkExclusion(0).Value = -((.Exclusion And 1) > 0)
    chkExclusion(1).Value = -((.Exclusion And 2) > 0)
    chkExclusion(2).Value = -((.Exclusion And 4) > 0)
    If Options.FSVersion <= Version_CFS1 Then
      chkExclusion(3).Value = -((.Exclusion And 8) > 0)
      cmbSynthetic.ListIndex = .Base
      cmbSize.ListIndex = .Size
    End If
    mChanged = False
    Show vbModal, Screen.ActiveForm
    If mChanged Then
      .Author = Txts(0).Text
      .Copyright = Txts(1).Text
      .Name = Txts(2).Text
      .LangCode = cmbLangCode.Text
      .Region = cmbRegion.Text
      .Country = Txts(3).Text
      .State = Txts(4).Text
      .City = Txts(5).Text
      .ICAOID = Txts(6).Text
      .Center.NumLatitude = LatLonCache(0)
      .Center.NumLongitude = LatLonCache(1)
      .Rotation = mValueCache(9)
      .Horz = mValueCache(10)
      .Vert = mValueCache(11)
      .Altitude = mValueCache(12)
      .MagVar = mValueCache(13)
      .Exclusion = chkExclusion(0).Value * 1 + _
                   chkExclusion(1).Value * 2 + _
                   chkExclusion(2).Value * 4 + _
                   chkExclusion(3).Value * 8
      If Options.FSVersion <= Version_CFS1 Then
        .Size = cmbSize.ListIndex
        .Base = cmbSynthetic.ListIndex
      End If
      
      Scenery.AFDRefresh = True
      EditData = True
    End If
  End With
  Unload frmHeader
End Function

' Change appropriate text labels after Lat/Lon change
Private Sub UpdateFields()
  Dim TempLatLon As clsLatLon, I As Integer, _
    SynTemp As Long

  Set TempLatLon = New clsLatLon
  With TempLatLon
    .Latitude = Txts(7).Text
    .Longitude = Txts(8).Text

    If .NumLatitude > -990 And .NumLongitude > -990 Then
      SynTemp = .SynSize
      Txts(13).Text = Append(TempLatLon.MagVar(), RES_Unit_AbbrevDeg)
    Else
      SynTemp = 0
      Txts(13).Text = ""
    End If
    
    If Options.FSVersion <= Version_CFS1 Then
      If Options.Metric Then
        For I = 1 To 5
          cmbSize.List(I) = Lang.ResolveString(RES_Hdr_MetersWide, SynTemp * (I * 2 - 1), Lang.GetString(RES_Unit_M))
        Next I
      Else
        For I = 1 To 5
          cmbSize.List(I) = Lang.ResolveString(RES_Hdr_MetersWide, SynTemp * (I * 2 - 1) * MToFt, Lang.GetString(RES_Unit_Ft))
        Next I
      End If
    End If
  End With
  Set TempLatLon = Nothing
End Sub

Private Sub btnSymbols_Click(Index As Integer)
  Dim rc As Integer
  rc = frmSymbols.GetSymbol()
  If rc > 0 Then
    Txts(Index).SelText = Chr$(rc)
  End If
  mDontSelect = True
End Sub

Private Sub btnSymbols_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbLeftButton Then btnSymbols(Index).Picture = frmMain.picSymbol(1).Image
End Sub

Private Sub btnSymbols_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Static Status As Boolean
  If Button = vbLeftButton Then
    If X < 0 Or X > btnSymbols(Index).Width Or Y < 0 Or Y > btnSymbols(Index).Height Then
      If Status Then btnSymbols(Index).Picture = frmMain.picSymbol(0).Image: Status = False
    Else
      If Not Status Then btnSymbols(Index).Picture = frmMain.picSymbol(1).Image: Status = True
    End If
  End If
End Sub

Private Sub btnSymbols_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  btnSymbols(Index).Picture = frmMain.picSymbol(0).Image
End Sub

Private Sub cmbLangCode_Change()
  Dim I As Integer, _
      LangIndex As Integer, _
      OldRegion As Integer, _
      OldText As String

  LangIndex = cmbLangCode.ListIndex
  If LangIndex = -1 Then LangIndex = 1
  OldRegion = cmbRegion.ListIndex
  If OldRegion = -1 Then OldText = cmbRegion.Text
  For I = 0 To 8
    cmbRegion.List(I) = Regions(LangIndex).Regions(I)
  Next I
  If OldRegion = -1 Then
    cmbRegion.Text = OldText
  Else
    cmbRegion.ListIndex = OldRegion
  End If
End Sub

Private Sub cmbLangCode_Click()
  cmbLangCode_Change
End Sub

Private Sub cmbSynthetic_Click()
  Dim Result As String, Index As Integer
  Index = cmbSynthetic.ListIndex
  If Index = 0 Then
    imgTexture.Picture = LoadPicture()
  Else
    LoadTexture AddDir(Options.FSPath, "Texture\") & SynNames(Index).File, Result
    If FileExists(Result) Then
      imgTexture.Picture = LoadPicture(Result)
      lbls(17).Visible = False
      Kill Result
    Else
      imgTexture.Picture = LoadPicture()
      lbls(17).Visible = True
    End If
  End If
End Sub

Private Sub cmdCancel_Click()
  Hide
End Sub

Private Sub cmdOK_Click()
  Dim I As Integer, _
    TempLatLon As clsLatLon, _
    Msg As String
  
  Set TempLatLon = New clsLatLon
  TempLatLon.Latitude = Txts(7).Text
  TempLatLon.Longitude = Txts(8).Text
  If Not TempLatLon.Validate(Msg) Then
    I = IIf(TempLatLon.NumLatitude < -990, 7, 8)
    Set TempLatLon = Nothing
    GoTo ValidationError:
  End If
  LatLonCache(0) = TempLatLon.NumLatitude
  LatLonCache(1) = TempLatLon.NumLongitude
  Set TempLatLon = Nothing

  For I = 3 To 13
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
  Dim I As Integer, HeightDiff As Single
  
  DialogMenus Me
  Lang.PrepareForm Me
  
  For I = 0 To 5
    btnSymbols(I).Picture = frmMain.picSymbol(0).Image
  Next I

  For I = 0 To 6
    cmbLangCode.AddItem Regions(I).ID
  Next I

  For I = 0 To 8
    cmbRegion.AddItem ""
  Next I

  If Options.FSVersion >= Version_FS2K Then
    chkExclusion(3).Enabled = False
    TabStrip1.Tabs.Remove 5
  Else
    Lang.AddItems cmbSynthetic, RES_Syn_Transparent, 27
    With cmbSize
      .AddItem Lang.GetString(RES_Hdr_None)
      For I = 1 To 5
        .AddItem I
      Next I
    End With
  End If
  
  HeightDiff = TabStrip1.ClientTop - 450
  
  For I = 1 To TabFrame.Count
    With TabFrame(I)
      .Move TabFrame(1).Left, TabStrip1.ClientTop + 150, TabFrame(1).Width, TabFrame(1).Height
      .Enabled = False
      .Visible = False
    End With
  Next I
  cmdOK.Top = cmdOK.Top + HeightDiff
  cmdCancel.Top = cmdCancel.Top + HeightDiff
  TabStrip1.Height = TabStrip1.Height + HeightDiff
  Height = Height + HeightDiff
  
  CenterForm Me
  
  SetEnabled Txts(13), Options.MagVarMissing
  
  Set TabStrip1.SelectedItem = TabStrip1.Tabs(1)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    Cancel = True
    Hide
  End If
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
          Case 2: Txts(2).SetFocus
          Case 3: Txts(7).SetFocus
          Case 4: chkExclusion(0).SetFocus
          Case 5: cmbSize.SetFocus
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
  If Index = 7 Or Index = 8 Then UpdateFields
End Sub

Private Sub Txts_DblClick(Index As Integer)
  If Between(Index, 9, 13) Then
    SmartSelectText Txts(Index)
  End If
End Sub

Private Sub Txts_GotFocus(Index As Integer)
  If mDontSelect Then
    mDontSelect = False
  Else
    If Between(Index, 9, 13) Then
      SmartSelectText Txts(Index)
    Else
      SelectText Txts(Index)
    End If
  End If
  If Index <= 5 Then ReturnSymbol 0, 0, 2
End Sub

Private Sub Txts_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Dim Result As Integer
    
  If Index <= 5 Then
    If Shift = vbCtrlMask And KeyCode = vbKeyS Then
      ' Ctrl+S = Key Combination to open Symbols menu
      btnSymbols_Click Index
      KeyCode = 0: Shift = 0
    ElseIf Shift > 1 Then
      Result = ReturnSymbol(KeyCode, Shift)
      If Result > 0 Then Txts(Index).SelText = Chr$(Result): KeyCode = 0
    End If
  End If
End Sub

Private Sub Txts_KeyPress(Index As Integer, KeyAscii As Integer)
  If Index <= 5 Then
    ' Suppress "Ding"
    If KeyAscii = 19 Then KeyAscii = 0 ' Ctrl+S
    KeyAscii = ReturnSymbol(KeyAscii, 0, 1)
  ElseIf Index = 6 Then
    If Between(KeyAscii, 97, 122) Then KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
  End If
End Sub
