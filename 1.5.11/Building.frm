VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmBuilding 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   ClipControls    =   0   'False
   Icon            =   "Building.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   6360
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      CausesValidation=   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   49
      Tag             =   "1031"
      Top             =   6360
      WhatsThisHelpID =   1031
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   48
      Tag             =   "1030"
      Top             =   6360
      WhatsThisHelpID =   1030
      Width           =   1215
   End
   Begin VB.Frame TabFrame 
      BorderStyle     =   0  'None
      Height          =   5535
      Index           =   2
      Left            =   360
      TabIndex        =   28
      Top             =   600
      Width           =   4440
      Begin VB.CheckBox chkSynchronize 
         Height          =   195
         Left            =   0
         TabIndex        =   47
         Tag             =   "1410"
         Top             =   5160
         WhatsThisHelpID =   1410
         Width           =   4455
      End
      Begin VB.ComboBox cmbTexture 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   4680
         WhatsThisHelpID =   1407
         Width           =   2655
      End
      Begin VB.PictureBox picPreview 
         Height          =   2655
         Left            =   1080
         ScaleHeight     =   173
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   173
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   1980
         Width           =   2655
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   14
         Left            =   3480
         TabIndex        =   42
         Tag             =   "1408"
         Top             =   1590
         WhatsThisHelpID =   1408
         Width           =   735
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   13
         Left            =   3480
         TabIndex        =   40
         Tag             =   "1408"
         Top             =   1200
         WhatsThisHelpID =   1408
         Width           =   735
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   12
         Left            =   3480
         TabIndex        =   38
         Tag             =   "1408"
         Top             =   810
         WhatsThisHelpID =   1408
         Width           =   735
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   11
         Left            =   1080
         TabIndex        =   36
         Tag             =   "1406"
         Top             =   1590
         WhatsThisHelpID =   1406
         Width           =   1095
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   10
         Left            =   1080
         TabIndex        =   34
         Tag             =   "1412"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   9
         Left            =   1080
         TabIndex        =   32
         Tag             =   "1411"
         Top             =   810
         Width           =   1095
      End
      Begin VB.ComboBox cmbLevel 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   120
         WhatsThisHelpID =   1405
         Width           =   2175
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   22
         Left            =   0
         TabIndex        =   45
         Tag             =   "1407"
         Top             =   4725
         WhatsThisHelpID =   1407
         Width           =   945
      End
      Begin VB.Label lbls 
         Height          =   915
         Index           =   21
         Left            =   0
         TabIndex        =   43
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   20
         Left            =   2400
         TabIndex        =   41
         Tag             =   "1408"
         Top             =   1620
         WhatsThisHelpID =   1408
         Width           =   1065
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   19
         Left            =   2400
         TabIndex        =   39
         Tag             =   "1408"
         Top             =   1230
         WhatsThisHelpID =   1408
         Width           =   1065
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   18
         Left            =   2400
         TabIndex        =   37
         Tag             =   "1408"
         Top             =   840
         WhatsThisHelpID =   1408
         Width           =   1065
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   17
         Left            =   0
         TabIndex        =   35
         Tag             =   "1406"
         Top             =   1620
         WhatsThisHelpID =   1406
         Width           =   945
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   16
         Left            =   0
         TabIndex        =   33
         Tag             =   "1412"
         Top             =   1230
         Width           =   945
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   15
         Left            =   0
         TabIndex        =   31
         Tag             =   "1411"
         Top             =   840
         Width           =   945
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   14
         Left            =   0
         TabIndex        =   29
         Tag             =   "1405"
         Top             =   165
         WhatsThisHelpID =   1405
         Width           =   945
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   0
         X2              =   4320
         Y1              =   615
         Y2              =   615
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   0
         X2              =   4320
         Y1              =   600
         Y2              =   600
      End
   End
   Begin VB.Frame TabFrame 
      BorderStyle     =   0  'None
      Height          =   5535
      Index           =   1
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   4440
      Begin VB.ComboBox Cmbs 
         Height          =   315
         Index           =   2
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   5025
         WhatsThisHelpID =   1049
         Width           =   2415
      End
      Begin VB.ComboBox Cmbs 
         Height          =   315
         Index           =   1
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   4635
         WhatsThisHelpID =   1404
         Width           =   2415
      End
      Begin VB.ComboBox Cmbs 
         Height          =   315
         Index           =   0
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   4245
         WhatsThisHelpID =   1403
         Width           =   2415
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   8
         Left            =   1320
         TabIndex        =   21
         Tag             =   "1402"
         Top             =   3870
         WhatsThisHelpID =   1402
         Width           =   1095
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   7
         Left            =   1320
         TabIndex        =   19
         Tag             =   "1401"
         Top             =   3480
         WhatsThisHelpID =   1400
         Width           =   1095
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   6
         Left            =   1320
         TabIndex        =   17
         Tag             =   "1400"
         Top             =   3090
         WhatsThisHelpID =   1400
         Width           =   1095
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   5
         Left            =   1320
         TabIndex        =   15
         Tag             =   "1048"
         Top             =   2700
         WhatsThisHelpID =   1048
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
         Index           =   13
         Left            =   0
         TabIndex        =   26
         Tag             =   "1049"
         Top             =   5070
         WhatsThisHelpID =   1049
         Width           =   1185
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   12
         Left            =   0
         TabIndex        =   24
         Tag             =   "1404"
         Top             =   4680
         WhatsThisHelpID =   1404
         Width           =   1185
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   11
         Left            =   0
         TabIndex        =   22
         Tag             =   "1403"
         Top             =   4290
         WhatsThisHelpID =   1403
         Width           =   1185
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   10
         Left            =   0
         TabIndex        =   20
         Tag             =   "1402"
         Top             =   3900
         WhatsThisHelpID =   1402
         Width           =   1185
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   9
         Left            =   0
         TabIndex        =   18
         Tag             =   "1401"
         Top             =   3510
         WhatsThisHelpID =   1400
         Width           =   1185
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   8
         Left            =   0
         TabIndex        =   16
         Tag             =   "1400"
         Top             =   3120
         WhatsThisHelpID =   1400
         Width           =   1185
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   7
         Left            =   0
         TabIndex        =   14
         Tag             =   "1048"
         Top             =   2730
         WhatsThisHelpID =   1048
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
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   6135
      Left            =   120
      TabIndex        =   50
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   10821
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "A"
            Object.Tag             =   "1435"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "B"
            Object.Tag             =   "1436"
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

Private mValueCache(10) As Single

' Index 0 for old building, 1-4 for new building
' so we can keep old values even if user switches
' between the two
Private mBuildLevels(4) As clsLevel

Private mPreview As clsOpenGL

Private mTxtChanged As Boolean
Private mChanged As Boolean

Public Function EditData(Data As clsBuilding) As Boolean
  Dim I As Integer
  
  Load frmBuilding
  
  With Data
    If .BuildType > 3 Then
      ' Advanced
      For I = 1 To 4
        .Level(I - 1).CopyTo mBuildLevels(I)
      Next I
      
      ' Default
      mBuildLevels(0).Height = 20
    Else
      .Level(0).CopyTo mBuildLevels(0)
      For I = 1 To 4
        Set mBuildLevels(I) = New clsLevel
        mBuildLevels(I).SetDefault
      Next I
      
      ' Default
      mBuildLevels(1).Height = 20
      mBuildLevels(1).TexID = 8
      mBuildLevels(2).TexID = 5
      mBuildLevels(2).X = 2.5
      mBuildLevels(2).Y = 2.5
      mBuildLevels(3).TexID = 8
      mBuildLevels(4).TexID = 4
    End If
    
    Txts(0).Tag = .Caption(True)
    Txts(0).Text = .Name
    Txts_Change 0 ' (Changes the caption)
    Txts(1).Text = MeterToUser(.X)
    Txts(2).Text = MeterToUser(.Y)
    Txts_Validate 1, False ' (Updates Latitude, Longitude)
    chkLocked.Value = -.Locked
    Txts(5).Text = GeographicToUser(.Rotation)
    Txts(6).Text = MeterToUser(.Length)
    Txts(7).Text = MeterToUser(.Width)
    Txts(8).Text = MeterToUser(.Altitude)
    If .BuildType = 6 Then
      mValueCache(9) = .RoofLength
      mValueCache(10) = .RoofWidth
    End If
    Cmbs(0).ListIndex = .BuildType
    Cmbs(1).ListIndex = .RoofLight
    Cmbs(2).ListIndex = .Complexity
    chkSynchronize.Value = -.Synchronize
  
    mChanged = False
    If TabValue > 1 Then TabValue = 0
    Set TabStrip1.SelectedItem = TabStrip1.Tabs(TabValue + 1)
    Show vbModal, Screen.ActiveForm
    If mChanged Then
      .Name = Txts(0).Text
      .X = mValueCache(1)
      .Y = mValueCache(2)
      .Locked = -chkLocked.Value
      .Rotation = mValueCache(5)
      .Length = mValueCache(6)
      .Width = mValueCache(7)
      .Altitude = mValueCache(8)
      .RoofLength = mValueCache(9)
      .RoofWidth = mValueCache(10)
      .Synchronize = -chkSynchronize.Value
      .BuildType = Cmbs(0).ListIndex
      .RoofLight = Cmbs(1).ListIndex
      .Complexity = Cmbs(2).ListIndex
      
      If .BuildType >= 4 Then
        ' Advanced
        For I = 1 To 4
          mBuildLevels(I).CopyTo .Level(I - 1)
        Next I
      Else
        mBuildLevels(0).CopyTo .Level(0)
      End If
      EditData = True
    End If
  End With

  Unload frmBuilding
End Function

Public Function EditDataM(Multi() As clsObject) As Boolean
  Dim I As Integer, J As Integer, _
    Data As clsBuilding, Temp As clsBuilding
  Dim IgnoreValue(14) As Boolean, _
    IgnoreLevelValue(3, 4) As Boolean
  
  ' Keeps track of which values need to be stored, and
  ' which are skipped
  
  ' When loading to the form, a false value means
  '   fill the control, true means make the control
  '   indeterminate (blank)
  ' When recording changes, a false value means store
  '   the value, true means do not store the value

  Set Data = Multi(1)
  
  For I = 2 To UBound(Multi)
    Set Temp = Multi(I)
    With Temp
      IgnoreValue(0) = IgnoreValue(0) Or Data.Name <> .Name
      IgnoreValue(1) = IgnoreValue(1) Or Data.X <> .X
      IgnoreValue(2) = IgnoreValue(2) Or Data.Y <> .Y
      IgnoreValue(3) = IgnoreValue(3) Or Data.Locked <> .Locked
      IgnoreValue(5) = IgnoreValue(5) Or Data.Rotation <> .Rotation
      IgnoreValue(6) = IgnoreValue(6) Or Data.Length <> .Length
      IgnoreValue(7) = IgnoreValue(7) Or Data.Width <> .Width
      IgnoreValue(8) = IgnoreValue(8) Or Data.Altitude <> .Altitude
      IgnoreValue(9) = IgnoreValue(9) Or Data.BuildType <> .BuildType
      IgnoreValue(10) = IgnoreValue(10) Or Data.RoofLight <> .RoofLight
      IgnoreValue(11) = IgnoreValue(11) Or Data.Complexity <> .Complexity
      IgnoreValue(12) = IgnoreValue(12) Or Data.RoofLength <> .RoofLength
      IgnoreValue(13) = IgnoreValue(13) Or Data.RoofWidth <> .RoofWidth
      IgnoreValue(14) = IgnoreValue(14) Or Data.Synchronize <> .Synchronize

      If Not IgnoreValue(9) Then
        ' 0 to 3 for advanced
        ' 0 for basic
        For J = 0 To IIf(.BuildType >= 4, 3, 0)
          IgnoreLevelValue(J, 0) = IgnoreLevelValue(J, 0) Or Data.Level(J).Height <> .Level(J).Height
          IgnoreLevelValue(J, 1) = IgnoreLevelValue(J, 1) Or Data.Level(J).TexID <> .Level(J).TexID
          IgnoreLevelValue(J, 2) = IgnoreLevelValue(J, 2) Or Data.Level(J).X <> .Level(J).X
          IgnoreLevelValue(J, 3) = IgnoreLevelValue(J, 3) Or Data.Level(J).Y <> .Level(J).Y
          IgnoreLevelValue(J, 4) = IgnoreLevelValue(J, 4) Or Data.Level(J).Z <> .Level(J).Z
        Next J
      End If
    End With
    Set Temp = Nothing
  Next I
  
  Load frmBuilding
  MultiSelection = True
  
  ' Fill Data
  With Data
    If .BuildType >= 4 Then
      ' Advanced
      For I = 1 To 4
        .Level(I - 1).CopyTo mBuildLevels(I)
        If IgnoreLevelValue(I - 1, 0) Then mBuildLevels(I).Height = -999
        If IgnoreLevelValue(I - 1, 1) Then mBuildLevels(I).TexID = 255
        If IgnoreLevelValue(I - 1, 2) Then mBuildLevels(I).X = -999
        If IgnoreLevelValue(I - 1, 3) Then mBuildLevels(I).Y = -999
        If IgnoreLevelValue(I - 1, 4) Then mBuildLevels(I).Z = -999
      Next I
      mBuildLevels(0).Height = 20
    Else
      .Level(0).CopyTo mBuildLevels(0)
      If IgnoreLevelValue(0, 0) Then mBuildLevels(0).Height = -999
      If IgnoreLevelValue(0, 1) Then mBuildLevels(0).TexID = 255
      If IgnoreLevelValue(0, 2) Then mBuildLevels(0).X = -999
      If IgnoreLevelValue(0, 3) Then mBuildLevels(0).Y = -999
      If IgnoreLevelValue(0, 4) Then mBuildLevels(0).Z = -999
      For I = 1 To 4
        Set mBuildLevels(I) = New clsLevel
        mBuildLevels(I).SetDefault
      Next I
      mBuildLevels(1).Height = 20
      mBuildLevels(1).TexID = 8
      mBuildLevels(2).TexID = 5
      mBuildLevels(3).TexID = 8
      mBuildLevels(4).TexID = 4
    End If
    
    If Not IgnoreValue(0) Then Txts(0).Text = .Name
    Txts(0).Tag = Lang.GetString(RES_Obj_Building)
    Txts_Change 0
    If Not IgnoreValue(1) Then Txts(1).Text = MeterToUser(.X)
    If Not IgnoreValue(2) Then Txts(2).Text = MeterToUser(.Y)
    Txts_Validate 1, False
    If Not IgnoreValue(3) Then
      chkLocked.Value = -.Locked
    Else
      chkLocked.Value = vbGrayed
    End If
    If Not IgnoreValue(5) Then Txts(5).Text = GeographicToUser(.Rotation)
    If Not IgnoreValue(6) Then Txts(6).Text = MeterToUser(.Length)
    If Not IgnoreValue(7) Then Txts(7).Text = MeterToUser(.Width)
    If Not IgnoreValue(8) Then Txts(8).Text = MeterToUser(.Altitude)
    If .BuildType = 6 Then
      If Not IgnoreValue(12) Then
        mValueCache(9) = .RoofLength
      Else
        mValueCache(9) = -999
      End If
      
      If Not IgnoreValue(13) Then
        mValueCache(10) = .RoofWidth
      Else
        mValueCache(10) = -999
      End If
    End If
    If Not IgnoreValue(9) Then Cmbs(0).ListIndex = .BuildType
    If Not IgnoreValue(10) Then Cmbs(1).ListIndex = .RoofLight
    If Not IgnoreValue(11) Then Cmbs(2).ListIndex = .Complexity
    If Not IgnoreValue(14) Then
      chkSynchronize.Value = -.Synchronize
    Else
      chkSynchronize.Value = vbGrayed
    End If
  End With

  Cmbs_Click 0
  cmbLevel_Click
  cmbTexture_Click
  
  mChanged = False
  If TabValue > 1 Then TabValue = 0
  Set TabStrip1.SelectedItem = TabStrip1.Tabs(TabValue + 1)
  Show vbModal, Screen.ActiveForm
  If mChanged Then
    IgnoreValue(0) = Txts(0).Text = ""
    IgnoreValue(1) = Txts(1).Text = ""
    IgnoreValue(2) = Txts(2).Text = ""
    IgnoreValue(3) = chkLocked.Value = vbGrayed
    IgnoreValue(5) = Txts(5).Text = ""
    IgnoreValue(6) = Txts(6).Text = ""
    IgnoreValue(7) = Txts(7).Text = ""
    IgnoreValue(8) = Txts(8).Text = ""
    IgnoreValue(9) = Cmbs(0).ListIndex = -1
    IgnoreValue(10) = Cmbs(1).ListIndex = -1
    IgnoreValue(11) = Cmbs(2).ListIndex = -1
    IgnoreValue(12) = mValueCache(9) <= -999
    IgnoreValue(13) = mValueCache(10) <= -999
    IgnoreValue(14) = chkSynchronize.Value = vbGrayed
        
    If Cmbs(0).ListIndex >= 4 Then
      For J = 1 To 4
        IgnoreLevelValue(J - 1, 0) = mBuildLevels(J).Height = -999
        IgnoreLevelValue(J - 1, 1) = mBuildLevels(J).TexID = 255
        IgnoreLevelValue(J - 1, 2) = mBuildLevels(J).X = -999
        IgnoreLevelValue(J - 1, 3) = mBuildLevels(J).Y = -999
        IgnoreLevelValue(J - 1, 4) = mBuildLevels(J).Z = -999
      Next J
    ElseIf Cmbs(0).ListIndex >= 0 Then
      IgnoreLevelValue(0, 0) = mBuildLevels(0).Height = -999
      IgnoreLevelValue(0, 1) = mBuildLevels(0).TexID = 255
      IgnoreLevelValue(0, 2) = mBuildLevels(0).X = -999
      IgnoreLevelValue(0, 3) = mBuildLevels(0).Y = -999
      IgnoreLevelValue(0, 4) = mBuildLevels(0).Z = -999
    Else
      For J = 0 To 3
        IgnoreLevelValue(J, 0) = True
        IgnoreLevelValue(J, 1) = True
        IgnoreLevelValue(J, 2) = True
        IgnoreLevelValue(J, 3) = True
        IgnoreLevelValue(J, 4) = True
      Next J
    End If

    For I = 1 To UBound(Multi)
      Set Temp = Multi(I)
      With Temp
        If Not IgnoreValue(0) Then .Name = Txts(0).Text
        If Not IgnoreValue(1) Then .X = mValueCache(1)
        If Not IgnoreValue(2) Then .Y = mValueCache(2)
        If Not IgnoreValue(3) Then .Locked = -chkLocked.Value
        If Not IgnoreValue(5) Then .Rotation = mValueCache(5)
        If Not IgnoreValue(6) Then .Length = mValueCache(6)
        If Not IgnoreValue(7) Then .Width = mValueCache(7)
        If Not IgnoreValue(8) Then .Altitude = mValueCache(8)
        If Not IgnoreValue(9) Then .BuildType = Cmbs(0).ListIndex
        If Not IgnoreValue(10) Then .RoofLight = Cmbs(1).ListIndex
        If Not IgnoreValue(11) Then .Complexity = Cmbs(2).ListIndex
        If Not IgnoreValue(12) Then .RoofLength = mValueCache(9)
        If Not IgnoreValue(13) Then .RoofWidth = mValueCache(10)
        If Not IgnoreValue(14) Then .Synchronize = -chkSynchronize.Value
        
        If Cmbs(0).ListIndex >= 4 Then
          ' Advanced
          For J = 0 To 3
            If Not IgnoreLevelValue(J, 0) Then .Level(J).Height = mBuildLevels(J + 1).Height
            If Not IgnoreLevelValue(J, 1) Then .Level(J).TexID = mBuildLevels(J + 1).TexID
            If Not IgnoreLevelValue(J, 2) Then .Level(J).X = mBuildLevels(J + 1).X
            If Not IgnoreLevelValue(J, 3) Then .Level(J).Y = mBuildLevels(J + 1).Y
            If Not IgnoreLevelValue(J, 4) Then .Level(J).Z = mBuildLevels(J + 1).Z
          Next J
        Else
          If Not IgnoreLevelValue(0, 0) Then .Level(0).Height = mBuildLevels(0).Height
          If Not IgnoreLevelValue(0, 1) Then .Level(0).TexID = mBuildLevels(0).TexID
          If Not IgnoreLevelValue(0, 2) Then .Level(0).X = mBuildLevels(0).X
          If Not IgnoreLevelValue(0, 3) Then .Level(0).Y = mBuildLevels(0).Y
          If Not IgnoreLevelValue(0, 4) Then .Level(0).Z = mBuildLevels(0).Z
        End If
      End With
    Next I
    EditDataM = True
  End If
  
  Unload frmBuilding
  MultiSelection = False
End Function

Private Sub chkLocked_Click()
  Dim Value As Boolean
  Value = Not -chkLocked.Value
  SetEnabled Txts(1), Value
  SetEnabled Txts(2), Value
  SetEnabled Txts(3), Value
  SetEnabled Txts(4), Value
End Sub

Private Sub chkSynchronize_Click()
  If chkSynchronize.Value = vbChecked Then
    Select Case cmbLevel.ListIndex
      Case 0
        ' Level 1
        mBuildLevels(3).TexID = cmbTexture.ListIndex + 8
      Case 2
        ' Level 3
        mBuildLevels(1).TexID = cmbTexture.ListIndex + 8
      Case 1, 3
        mBuildLevels(3).TexID = mBuildLevels(1).TexID
    End Select
    cmbTexture_Click
  End If
End Sub

Private Sub cmbLevel_Click()
  Dim I As Integer, LevelIndex As Integer
  cmbTexture.Clear
  Select Case Cmbs(0).ListIndex
    Case Is >= 4
      ' Advanced
      LevelIndex = cmbLevel.ListIndex + 1
      
      Select Case LevelIndex
        Case 1, 3
          ' Lobby
          For I = 8 To 85
            cmbTexture.AddItem Building1_3(I)
          Next I
          If mBuildLevels(LevelIndex).TexID = 255 Then
            cmbTexture.ListIndex = -1
          Else
            cmbTexture.ListIndex = mBuildLevels(LevelIndex).TexID - 8
          End If
        Case 2
          For I = 4 To 84
            cmbTexture.AddItem Building2(I)
          Next I
          If mBuildLevels(LevelIndex).TexID = 255 Then
            cmbTexture.ListIndex = -1
          Else
            cmbTexture.ListIndex = mBuildLevels(LevelIndex).TexID - 4
          End If
        Case 4
          For I = 4 To 33
            cmbTexture.AddItem BuildingR(I)
          Next I
          If mBuildLevels(LevelIndex).TexID = 255 Then
            cmbTexture.ListIndex = -1
          Else
            cmbTexture.ListIndex = mBuildLevels(LevelIndex).TexID - 4
          End If
      End Select
      
      With mBuildLevels(LevelIndex)
        If .Height > -999 Then
          Txts(11).Text = MeterToUser(.Height)
        Else
          Txts(11).Text = ""
        End If
        If .X > -999 Then
          Txts(12).Text = CSng(.X)
        Else
          Txts(12).Text = ""
        End If
        If .Y > -999 Then
          Txts(13).Text = CSng(.Y)
        Else
          Txts(13).Text = ""
        End If
        If .Z > -999 Then
          Txts(14).Text = CSng(.Z)
        Else
          Txts(14).Text = ""
        End If
      End With
    Case Is >= 0
      If Options.FSVersion >= Version_FS2K Then
        For I = 0 To 7
          cmbTexture.AddItem Building2(Building(I).Windows)
        Next I
      Else
        Lang.AddItems cmbTexture, RES_Bldg_Texture1, 8
      End If
      With mBuildLevels(0)
        If .Height > -999 Then
          Txts(11).Text = MeterToUser(.Height)
        Else
          Txts(11).Text = ""
        End If
        If .TexID = 255 Then
          cmbTexture.ListIndex = -1
        Else
          cmbTexture.ListIndex = mBuildLevels(0).TexID
        End If
      End With
    Case Else
      cmbTexture.ListIndex = -1
  End Select
  
  If Cmbs(0).ListIndex = 6 And LevelIndex = 4 Then
    ' Roof of slanted building
    If mValueCache(9) > -999 Then
      Txts(9).Text = MeterToUser(mValueCache(9))
    Else
      Txts(9).Text = ""
    End If
    If mValueCache(10) > -999 Then
      Txts(10).Text = MeterToUser(mValueCache(10))
    Else
      Txts(10).Text = ""
    End If
    SetEnabled Txts(9), True
    SetEnabled Txts(10), True
    lbls(15).WhatsThisHelpID = 1411
    lbls(16).WhatsThisHelpID = 1411
    Txts(9).WhatsThisHelpID = 1411
    Txts(10).WhatsThisHelpID = 1411
  Else
    Txts(9).Text = Txts(6).Text
    Txts(10).Text = Txts(7).Text
    SetEnabled Txts(9), False
    SetEnabled Txts(10), False
    lbls(15).WhatsThisHelpID = 1400
    lbls(16).WhatsThisHelpID = 1400
    Txts(9).WhatsThisHelpID = 1400
    Txts(10).WhatsThisHelpID = 1400
  End If
  
  SetEnabled Txts(11), Between(LevelIndex, 0, 3) Or (Cmbs(0).ListIndex = 5) Or (Cmbs(0).ListIndex >= 7)
  SetEnabled Txts(13), Not (Cmbs(0).ListIndex >= 7 And (LevelIndex = 1 Or LevelIndex = 3))
  SetEnabled Txts(14), (LevelIndex = 2 And Cmbs(0).ListIndex <= 6) Or (LevelIndex = 4 And Cmbs(0).ListIndex = 5)
End Sub

Private Sub Cmbs_Click(Index As Integer)
  Dim I As Integer, blnValue As Boolean
  If Index = 0 Then
    cmbLevel.Clear
    Select Case Cmbs(0).ListIndex
      Case Is >= 4
        Lang.AddItems cmbLevel, RES_Bldg_Levels1 + 1, 4
        cmbLevel.ListIndex = 0
        blnValue = True
      Case Is >= 0
        Lang.AddItems cmbLevel, RES_Bldg_Levels1, 1
        cmbLevel.ListIndex = 0
      Case Else
        cmbLevel.ListIndex = -1
    End Select
    cmbLevel_Click
    For I = 12 To 14
      Txts(I).Visible = blnValue
      lbls(I + 6).Visible = blnValue
    Next I
    chkSynchronize.Visible = blnValue
  End If
End Sub

Private Sub cmbTexture_Click()
  Dim Pts(3) As PointType, TexCoords(3) As PointType
  Dim ID As Integer, TexID As Integer, _
    TexErrFlag As Boolean, Value As Integer, _
    File As String
  
  On Error Resume Next

  mPreview.StartDraw
  TexCoords(1).X = 1
  TexCoords(2).X = 1
  TexCoords(2).Y = 1
  TexCoords(3).Y = 1
  glClearColor 0.5, 0.5, 0.5, 0
  glCls
  
  ' We don't delete any of the bitmap handles, just
  ' keep everything in cache so stuff loads faster

  Select Case Cmbs(0).ListIndex
    Case Is >= 4
      ' Advanced
      TexID = cmbTexture.ListIndex
      ' TexID = -1 is the undefined flag. But we still
      ' want to draw the outline of the building
      If TexID > -1 Then
        Select Case cmbLevel.ListIndex + 1
          Case 1
            TexID = TexID + 8
            If chkSynchronize.Value = vbChecked Then mBuildLevels(3).TexID = TexID
          Case 3
            TexID = TexID + 8
            If chkSynchronize.Value = vbChecked Then mBuildLevels(1).TexID = TexID
          Case 2, 4
            TexID = TexID + 4
        End Select
        mBuildLevels(cmbLevel.ListIndex + 1).TexID = TexID
      End If
    
      ' Roof
      Pts(0) = MakePoint(25, 95)
      Pts(1) = MakePoint(95, 85)
      Pts(2) = MakePoint(75, 65)
      Pts(3) = MakePoint(5, 75)
      
      ' Separate lines, so that if TexID = -1, the error
      ' occurs on the first line, and we get a -1 for the
      ' error ID on the second line
      File = BuildingR(mBuildLevels(4).TexID)
      ID = mPreview.LoadBitmap(AddDir(Options.FSPath, "Texture\" & File))
      
      TexCoords(0) = MakePoint(0, 0)
      TexCoords(1) = MakePoint(mBuildLevels(4).X, 0)
      TexCoords(2) = MakePoint(mBuildLevels(4).X, mBuildLevels(4).Y)
      TexCoords(3) = MakePoint(0, mBuildLevels(4).Y)
      GoSub DrawTexture
      
      ' Level 3
      Pts(0) = MakePoint(5, 75)
      Pts(1) = MakePoint(75, 65)
      Pts(2) = MakePoint(75, 50)
      Pts(3) = MakePoint(5, 60)
      
      File = Building1_3(mBuildLevels(3).TexID)
      Value = Val(ReadLast(File, " "))
      ID = mPreview.LoadBitmap(AddDir(Options.FSPath, "Texture\" & File))
      
      If Value = 1 Then
        TexCoords(0) = MakePoint(0, 128 / 255)
        TexCoords(1) = MakePoint(mBuildLevels(3).X, 128 / 255)
        TexCoords(2) = MakePoint(mBuildLevels(3).X, 191 / 255)
        TexCoords(3) = MakePoint(0, 191 / 255)
      Else
        TexCoords(0) = MakePoint(0, 0)
        TexCoords(1) = MakePoint(mBuildLevels(3).X, 0)
        TexCoords(2) = MakePoint(mBuildLevels(3).X, 63 / 255)
        TexCoords(3) = MakePoint(0, 63 / 255)
      End If
      GoSub DrawTexture
  
      Pts(0) = MakePoint(75, 65)
      Pts(1) = MakePoint(95, 85)
      Pts(2) = MakePoint(95, 70)
      Pts(3) = MakePoint(75, 50)
      
      TexCoords(1).X = mBuildLevels(3).Y
      TexCoords(2).X = mBuildLevels(3).Y
      GoSub DrawTexture
  
      ' Level 2
      Pts(0) = MakePoint(5, 60)
      Pts(1) = MakePoint(75, 50)
      Pts(2) = MakePoint(75, 20)
      Pts(3) = MakePoint(5, 30)
      
      TexCoords(0) = MakePoint(0, 0)
      TexCoords(1) = MakePoint(mBuildLevels(2).X, 0)
      TexCoords(2) = MakePoint(mBuildLevels(2).X, mBuildLevels(2).Z)
      TexCoords(3) = MakePoint(0, mBuildLevels(2).Z)
  
      File = Building2(mBuildLevels(2).TexID)
      ID = mPreview.LoadBitmap(AddDir(Options.FSPath, "Texture\" & File))
      GoSub DrawTexture
  
      Pts(0) = MakePoint(75, 50)
      Pts(1) = MakePoint(95, 70)
      Pts(2) = MakePoint(95, 40)
      Pts(3) = MakePoint(75, 20)
      
      TexCoords(1).X = mBuildLevels(2).Y
      TexCoords(2).X = mBuildLevels(2).Y
      GoSub DrawTexture
      
      ' Level 1
      Pts(0) = MakePoint(5, 30)
      Pts(1) = MakePoint(75, 20)
      Pts(2) = MakePoint(75, 5)
      Pts(3) = MakePoint(5, 15)
  
      File = Building1_3(mBuildLevels(1).TexID)
      Value = Val(ReadLast(File, " "))
      ID = mPreview.LoadBitmap(AddDir(Options.FSPath, "Texture\" & File))
      
      If Value = 1 Then
        TexCoords(0) = MakePoint(0, 192 / 255)
        TexCoords(1) = MakePoint(mBuildLevels(1).X, 192 / 255)
        TexCoords(2) = MakePoint(mBuildLevels(1).X, 1)
        TexCoords(3) = MakePoint(0, 1)
      Else
        TexCoords(0) = MakePoint(0, 64 / 255)
        TexCoords(1) = MakePoint(mBuildLevels(1).X, 64 / 255)
        TexCoords(2) = MakePoint(mBuildLevels(1).X, 127 / 255)
        TexCoords(3) = MakePoint(0, 127 / 255)
      End If
          
      GoSub DrawTexture
  
      Pts(0) = MakePoint(75, 20)
      Pts(1) = MakePoint(95, 40)
      Pts(2) = MakePoint(95, 25)
      Pts(3) = MakePoint(75, 5)
      
      TexCoords(1).X = mBuildLevels(1).Y
      TexCoords(2).X = mBuildLevels(1).Y
      GoSub DrawTexture
    Case Is >= 0
      ' Basic
      TexID = cmbTexture.ListIndex
      If TexID > -1 Then mBuildLevels(0).TexID = TexID
        
      ' Roof
      Pts(0) = MakePoint(25, 95)
      Pts(1) = MakePoint(95, 85)
      Pts(2) = MakePoint(75, 65)
      Pts(3) = MakePoint(5, 75)
      If Options.FSVersion >= Version_FS2K And Cmbs(0).ListIndex = 0 Then
        ' See note above on splitting into two lines
        File = BuildingR(Building(TexID).Roof)
        ID = mPreview.LoadBitmap(AddDir(Options.FSPath, "Texture\" & File))
        GoSub DrawTexture
        File = Building2(Building(TexID).Windows)
        ID = mPreview.LoadBitmap(AddDir(Options.FSPath, "Texture\" & File))
      Else
        glForeColor Building(TexID).Color
        glPaintRegion Pts
        ID = mPreview.LoadBitmap(AddDir(Options.FSPath, "Texture\" & Building(TexID).File), IIf(Options.FSVersion >= Version_FS2K, True, False))
        TexCoords(0) = MakePoint(0, 1 / 255)
        TexCoords(1) = MakePoint(17 / 255, 1 / 255)
        TexCoords(2) = MakePoint(17 / 255, 18 / 255)
        TexCoords(3) = MakePoint(0, 18 / 255)
      End If
      
      ' Windows
      Pts(0) = MakePoint(5, 75)
      Pts(1) = MakePoint(75, 65)
      Pts(2) = MakePoint(75, 5)
      Pts(3) = MakePoint(5, 15)
      GoSub DrawTexture
      
      Pts(0) = MakePoint(75, 65)
      Pts(1) = MakePoint(95, 85)
      Pts(2) = MakePoint(95, 25)
      Pts(3) = MakePoint(75, 5)
      GoSub DrawTexture
  End Select
  lbls(21).Caption = Lang.GetString(IIf(TexErrFlag, RES_ERR_TextureError, RES_Bldg_NotToScale))
  mPreview.CopyTo picPreview.hdc
  Exit Sub
DrawTexture:
  If ID = -1 Then
    TexErrFlag = True
    glColor3f 0, 0, 0
    glDrawPolygon Pts
  Else
    mPreview.SelectBitmap ID
    glPaintTexturedRegion2 Pts, TexCoords
  End If
  Return
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
  
  For I = 1 To 8
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

  Lang.PrepareForm Me
  
  If Options.FSVersion >= Version_FS2K Then
    Lang.AddItems Cmbs(0), RES_Bldg_Shape1, 7
    For I = 3 To 32
      Cmbs(0).AddItem Lang.ResolveString(RES_Bldg_ShapeP, I)
    Next I
  Else
    Lang.AddItems Cmbs(0), RES_Bldg_Shape1, 4
  End If
  Lang.AddItems Cmbs(1), RES_Bldg_RoofLight1, 3
  Lang.AddItems Cmbs(2), RES_Complexity1, IIf(Options.FSVersion < Version_FS2K, 5, 6)
  
  Set mPreview = New clsOpenGL
  With mPreview
    .PhysicalResize picPreview.ScaleWidth, picPreview.ScaleHeight
    .SetScale 0, 100, 100, 0
  End With

  For I = 0 To 4
    Set mBuildLevels(I) = New clsLevel
  Next I
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then Cancel = True: Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim I As Integer
  
  For I = 0 To 4
    Set mBuildLevels(I) = Nothing
  Next I
  Set mPreview = Nothing
End Sub

Private Sub picPreview_Paint()
  mPreview.CopyTo picPreview.hdc
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
          Case 2:
            cmbLevel.SetFocus
            cmbLevel_Click
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
  If Index = 1 Or Index = 2 Or Index >= 5 Then
    SmartSelectText Txts(Index)
  Else
    If Index = 0 Then ReturnSymbol 0, 0, 2
    SelectText Txts(Index)
  End If
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
    LevelIndex As Integer, Msg As String, _
    Value As Single, TempLatLon As clsLatLon

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
    Case 9, 10
      If Txts(Index).Locked Then Exit Sub
      Validate Txts(Index - 3), "", valX
      If valX = 0 Then Exit Sub
      If Not Validate(Txts(Index), Msg, mValueCache(Index), , valX) Then
        Cancel = True
        MsgBoxEx Me, Msg, vbCritical, RES_ERR_Bound
      End If
      If Txts(Index).Text = "" Then mValueCache(Index) = -999
    Case 11, 12, 13, 14
      If Txts(Index).Locked Then Exit Sub
      If Cmbs(0).ListIndex <= 3 Then
        LevelIndex = 0
      Else
        LevelIndex = cmbLevel.ListIndex + 1
      End If
      
      If Txts(Index).Text = "" Then
        Value = -999
      Else
        Value = ValEx(Txts(Index).Text)
      End If

      Select Case Index
        Case 11
          If Value > -999 Then
            If Not Validate(Txts(11), Msg, Value) Then
              Cancel = True
              MsgBoxEx Me, Msg, vbCritical, RES_ERR_Bound
            Else
              mBuildLevels(LevelIndex).Height = Value
            End If
          Else
            mBuildLevels(LevelIndex).Height = CInt(Value)
          End If
        Case 12
          If Value > -999 Then
            If Not Validate(Txts(12), Msg, Value) Then
              Cancel = True
              MsgBoxEx Me, Msg, vbCritical, RES_ERR_Bound
            Else
              mBuildLevels(LevelIndex).X = Value
            End If
          Else
             mBuildLevels(LevelIndex).X = CInt(Value)
          End If
        Case 13
          If Value > -999 Then
            If Not Validate(Txts(13), Msg, Value) Then
              Cancel = True
              MsgBoxEx Me, Msg, vbCritical, RES_ERR_Bound
            Else
              mBuildLevels(LevelIndex).Y = Value
            End If
          Else
             mBuildLevels(LevelIndex).Y = CInt(Value)
          End If
        Case 14
          If Value > -999 Then
            If Not Validate(Txts(14), Msg, Value) Then
              Cancel = True
              MsgBoxEx Me, Msg, vbCritical, RES_ERR_Bound
            Else
              mBuildLevels(LevelIndex).Z = Value
            End If
          Else
             mBuildLevels(LevelIndex).Z = CInt(Value)
          End If
      End Select
      If Index <> 11 Then cmbTexture_Click
  End Select
  If Cancel Then
    mTxtChanged = True
    SmartSelectText Txts(Index)
  End If
End Sub
