VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmProperties 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5205
   ClipControls    =   0   'False
   Icon            =   "Properties.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   5400
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   54
      Tag             =   "1031"
      Top             =   5400
      WhatsThisHelpID =   1031
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   53
      Tag             =   "1030"
      Top             =   5400
      WhatsThisHelpID =   1030
      Width           =   1215
   End
   Begin VB.Frame TabFrame 
      BorderStyle     =   0  'None
      Height          =   4575
      Index           =   2
      Left            =   360
      TabIndex        =   18
      Top             =   600
      Width           =   4545
      Begin VB.CommandButton cmdEdit 
         Height          =   375
         Left            =   3120
         TabIndex        =   32
         Tag             =   "1329"
         Top             =   2025
         Visible         =   0   'False
         WhatsThisHelpID =   1329
         Width           =   1215
      End
      Begin VB.CheckBox chkPolyLocked 
         Height          =   195
         Left            =   0
         TabIndex        =   51
         Tag             =   "1317"
         Top             =   4320
         WhatsThisHelpID =   1317
         Width           =   3975
      End
      Begin VB.CheckBox chkNight 
         Height          =   195
         Left            =   0
         TabIndex        =   50
         Tag             =   "1315"
         Top             =   4080
         WhatsThisHelpID =   1315
         Width           =   3975
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   10
         Left            =   2280
         TabIndex        =   47
         Tag             =   "1307"
         Top             =   3630
         WhatsThisHelpID =   1307
         Width           =   855
      End
      Begin VB.CheckBox chkLine 
         Height          =   195
         Left            =   0
         TabIndex        =   45
         Tag             =   "1306"
         Top             =   3660
         WhatsThisHelpID =   1306
         Width           =   1185
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   9
         Left            =   2280
         TabIndex        =   42
         Tag             =   "1305"
         Top             =   3240
         WhatsThisHelpID =   1305
         Width           =   855
      End
      Begin VB.CheckBox chkDot 
         Height          =   195
         Left            =   0
         TabIndex        =   40
         Tag             =   "1304"
         Top             =   3270
         WhatsThisHelpID =   1304
         Width           =   1185
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   12
         Left            =   2280
         TabIndex        =   39
         Tag             =   "1320"
         Top             =   2850
         WhatsThisHelpID =   1320
         Width           =   855
      End
      Begin VB.CheckBox chkVisibility 
         Height          =   195
         Left            =   0
         TabIndex        =   37
         Tag             =   "1319"
         Top             =   2880
         WhatsThisHelpID =   1320
         Width           =   1185
      End
      Begin VB.ComboBox Cmbs 
         Height          =   315
         Index           =   3
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   2445
         WhatsThisHelpID =   1049
         Width           =   2775
      End
      Begin VB.ComboBox Cmbs 
         Height          =   315
         Index           =   2
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   2055
         WhatsThisHelpID =   1308
         Width           =   2775
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   11
         Left            =   1200
         TabIndex        =   29
         Tag             =   "1310"
         Top             =   1290
         WhatsThisHelpID =   1310
         Width           =   975
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   8
         Left            =   1200
         TabIndex        =   26
         Tag             =   "1303"
         Top             =   900
         WhatsThisHelpID =   1303
         Width           =   975
      End
      Begin VB.ComboBox Cmbs 
         Height          =   315
         Index           =   1
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   1665
         WhatsThisHelpID =   1302
         Width           =   2775
      End
      Begin VB.PictureBox picBackground 
         BorderStyle     =   0  'None
         Height          =   3615
         Left            =   4080
         ScaleHeight     =   241
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   420
         Width           =   420
         Begin VB.PictureBox picColor 
            Height          =   255
            Index           =   0
            Left            =   45
            ScaleHeight     =   195
            ScaleWidth      =   240
            TabIndex        =   24
            Top             =   105
            WhatsThisHelpID =   1301
            Width           =   300
         End
         Begin VB.PictureBox picColor 
            Height          =   255
            Index           =   1
            Left            =   45
            ScaleHeight     =   195
            ScaleWidth      =   240
            TabIndex        =   44
            Top             =   2835
            WhatsThisHelpID =   1301
            Width           =   300
         End
         Begin VB.PictureBox picColor 
            Height          =   255
            Index           =   2
            Left            =   45
            ScaleHeight     =   195
            ScaleWidth      =   240
            TabIndex        =   49
            Top             =   3225
            WhatsThisHelpID =   1301
            Width           =   300
         End
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   7
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   510
         WhatsThisHelpID =   1300
         Width           =   1935
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   6
         Left            =   1200
         TabIndex        =   20
         Top             =   120
         WhatsThisHelpID =   1040
         Width           =   3255
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   18
         Left            =   3240
         TabIndex        =   48
         Tag             =   "1301"
         Top             =   3660
         WhatsThisHelpID =   1301
         Width           =   705
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   17
         Left            =   1200
         TabIndex        =   46
         Tag             =   "1307"
         Top             =   3660
         WhatsThisHelpID =   1307
         Width           =   1035
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   16
         Left            =   3240
         TabIndex        =   43
         Tag             =   "1301"
         Top             =   3270
         WhatsThisHelpID =   1301
         Width           =   705
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   15
         Left            =   1200
         TabIndex        =   41
         Tag             =   "1305"
         Top             =   3270
         WhatsThisHelpID =   1305
         Width           =   1035
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   21
         Left            =   1200
         TabIndex        =   38
         Tag             =   "1320"
         Top             =   2880
         WhatsThisHelpID =   1320
         Width           =   1035
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   22
         Left            =   0
         TabIndex        =   35
         Tag             =   "1049"
         Top             =   2490
         WhatsThisHelpID =   1049
         Width           =   1185
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   19
         Left            =   0
         TabIndex        =   33
         Tag             =   "1308"
         Top             =   2100
         WhatsThisHelpID =   1308
         Width           =   1185
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   20
         Left            =   0
         TabIndex        =   28
         Tag             =   "1310"
         Top             =   1320
         WhatsThisHelpID =   1310
         Width           =   1065
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   14
         Left            =   2280
         TabIndex        =   27
         Tag             =   "1090"
         Top             =   930
         WhatsThisHelpID =   1303
         Width           =   1305
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   13
         Left            =   0
         TabIndex        =   25
         Tag             =   "1303"
         Top             =   930
         WhatsThisHelpID =   1303
         Width           =   1065
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   12
         Left            =   0
         TabIndex        =   30
         Tag             =   "1302"
         Top             =   1710
         WhatsThisHelpID =   1302
         Width           =   1185
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   11
         Left            =   3240
         TabIndex        =   23
         Tag             =   "1301"
         Top             =   540
         WhatsThisHelpID =   1301
         Width           =   705
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   10
         Left            =   0
         TabIndex        =   21
         Tag             =   "1300"
         Top             =   540
         WhatsThisHelpID =   1300
         Width           =   1065
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   9
         Left            =   0
         TabIndex        =   19
         Tag             =   "1040"
         Top             =   150
         WhatsThisHelpID =   1040
         Width           =   1065
      End
   End
   Begin VB.Frame TabFrame 
      BorderStyle     =   0  'None
      Height          =   4575
      Index           =   1
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   4545
      Begin VB.ComboBox Cmbs 
         Height          =   315
         Index           =   0
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   3075
         WhatsThisHelpID =   1951
         Width           =   2895
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   5
         Left            =   1320
         TabIndex        =   15
         Tag             =   "1321"
         Top             =   2700
         WhatsThisHelpID =   1321
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
         Width           =   3375
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   8
         Left            =   0
         TabIndex        =   16
         Tag             =   "1951"
         Top             =   3120
         WhatsThisHelpID =   1951
         Width           =   1185
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   7
         Left            =   0
         TabIndex        =   14
         Tag             =   "1321"
         Top             =   2730
         WhatsThisHelpID =   1321
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
      Height          =   5175
      Left            =   120
      TabIndex        =   55
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   9128
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "A"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "B"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long

Private mChanged As Boolean, mTxtChanged As Boolean, _
        mFocus As Integer, mType As ObjectTypeEnum
Private myLinkObject As clsObject

Private mValueCache(12) As Single

Private Sub DrawFocus(ByVal Index As Integer, ByVal Selected As Boolean)
  Dim R As RECT
  With picColor(Index)
    R.Left = .Left - 3
    R.Top = .Top - 3
    R.Right = .Left + 23
    R.Bottom = .Top + 20
  End With
  
  picBackground.Line (R.Left, R.Top)-(R.Right - 1, R.Bottom - 1), vbButtonFace, B
  If Selected Then DrawFocusRect picBackground.hdc, R
End Sub

Public Function EditData(Data As clsPoint) As Boolean
  Dim I As Integer, PointObject As clsPoint
  
  Dim ShapeObject As clsShape
  
  Load frmProperties
  
  With Data
    PrepareForm .Parent.ShapeType

    ' If there are no points (i.e. when shape is first
    ' created), then don't show the Point tab
    If .Parent.NumPoints = -1 Then
      TabStrip1.Tabs.Remove 1
    Else
      TabStrip1.Tabs(1).Caption = "  " & .Caption(True) & "  "
      Txts(0).Text = .Caption(True)
      Txts(1).Text = MeterToUser(.X)
      Txts(2).Text = MeterToUser(.Y)
      Txts(5).Text = MeterToUser(.Z)
      chkLocked.Value = -.Locked
      Txts_Validate 1, True ' (Updates Latitude, Longitude)
      Cmbs(0).ListIndex = .PtType
      If .Parent.ShapeType <> OT_Polygon Then
        If .ObjectIndex - 1 = .Parent.NumPoints Then
          ' This option doesn't make sense for the last point
          Cmbs(0).Locked = True
          Cmbs(0).BackColor = vbButtonFace
        End If
      End If
    End If
    
    With .Parent
      Txts(6).Tag = .Caption(True)
      Txts(6).Text = .Name
      Txts(7).Tag = .Texture
      Txts(7).Text = GetFileTitle(.Texture)
      picColor(0).Tag = .Color
      picColor(0).BackColor = ExtractColor(.Color)
      If Options.FSVersion >= Version_FS2K Then
        Cmbs(1).ListIndex = FS2KLayerToIndex(.Layer)
      Else
        Cmbs(1).ListIndex = .Layer
      End If
      Txts(8).Text = MeterToUser(.MScale, "##0.0#")
      chkDot.Value = -(.DotSpacing > 0)
      Select Case .ShapeType
        Case OT_Polygon, OT_Line
          If .DotSpacing = 0 Then
            Txts(9).Text = MeterToUser(60)
          Else
            Txts(9).Text = MeterToUser(.DotSpacing)
          End If
          picColor(1).Tag = .DotColor
          picColor(1).BackColor = ExtractColor(.DotColor)
          If .LinkObject Is Nothing Then
            Cmbs(2).ListIndex = 0
            cmdEdit.Enabled = False
          ElseIf TypeOf .LinkObject Is clsBuilding Then
            Cmbs(2).ListIndex = 1
            Set myLinkObject = .LinkObject
          ElseIf TypeOf .LinkObject Is clsMacro Then
            Cmbs(2).ListIndex = 2
            Set myLinkObject = .LinkObject
          Else
            Cmbs(2).ListIndex = .LinkObject.ObjectIndex
          End If
        Case OT_Taxiway
          If .DotSpacing = 0 Then
            Txts(9).Text = MeterToUser(60)
          Else
            Txts(9).Text = MeterToUser(.DotSpacing)
          End If
          picColor(1).Tag = .DotColor
          picColor(1).BackColor = ExtractColor(.DotColor)
          
          chkLine.Value = -(.LineWidth > 0)
          If .LineWidth = 0 Then
            Txts(10).Text = MeterToUser(2)
          Else
            Txts(10).Text = MeterToUser(.LineWidth)
          End If
          picColor(2).Tag = .LineColor
          picColor(2).BackColor = ExtractColor(.LineColor)
        Case OT_Road
          If Options.FSVersion >= Version_FS2K Then
            Cmbs(2).ListIndex = .Extra3
          End If
      End Select
      Txts(11).Text = MeterToUser(.Extra1)
      chkVisibility.Value = -(.Visibility = 0)
      Txts(12).Text = MeterToUser(.Visibility)
      Cmbs(3).ListIndex = .Complexity
      chkNight.Value = -.Extra2
      chkPolyLocked.Value = -.Locked
    End With
    
    Lang.PrepareForm Me
    chkDot_Click
    chkLine_Click
    TabStrip1_Click
    Txts_Change 6
    Txts_Change 7
    mTxtChanged = False
    mChanged = False
    If TabValue >= TabStrip1.Tabs.Count Then TabValue = 0
    Set TabStrip1.SelectedItem = TabStrip1.Tabs(TabValue + 1)
    Show vbModal, Screen.ActiveForm
    
    If mChanged Then
      If .Parent.NumPoints > -1 Then
        .X = mValueCache(1)
        .Y = mValueCache(2)
        .Z = mValueCache(5)
        .Locked = -chkLocked.Value
        .PtType = Cmbs(0).ListIndex
      End If

      With .Parent
        .Name = Txts(6).Text
        .Texture = Txts(7).Tag
        .Color = picColor(0).Tag
        If Options.FSVersion >= Version_FS2K Then
          .Layer = FS2KIndexToLayer(Cmbs(1).ListIndex)
        Else
          .Layer = Cmbs(1).ListIndex
        End If
        
        .MScale = mValueCache(8)
        
        Select Case .ShapeType
          Case OT_Polygon
            .DotSpacing = chkDot.Value * mValueCache(9)
            .DotColor = picColor(1).Tag
            If Cmbs(2).ListIndex = 0 Then
              Set .LinkObject = Nothing
            Else
              Set PointObject = Scenery(Cmbs(2).ItemData(Cmbs(2).ListIndex))
              Set .LinkObject = PointObject.Parent
              Set PointObject.Parent.LinkFromObj = Data.Parent
              Set PointObject = Nothing
            End If
          Case OT_Taxiway
            .DotSpacing = chkDot.Value * mValueCache(9)
            .DotColor = picColor(1).Tag
            .LineWidth = chkLine.Value * mValueCache(10)
            .LineColor = picColor(2).Tag
          Case OT_Road
            If Options.FSVersion >= Version_FS2K Then
              .Extra3 = Cmbs(2).ListIndex
            End If
          Case OT_Line
            .DotSpacing = chkDot.Value * mValueCache(9)
            .DotColor = picColor(1).Tag
            Set .LinkObject = Nothing
            If Cmbs(2).ListIndex > 0 Then Set .LinkObject = myLinkObject
        End Select
        
        .Extra1 = mValueCache(11)
        .Complexity = Cmbs(3).ListIndex
        .Extra2 = -chkNight.Value
        .Visibility = mValueCache(12) * Abs(chkVisibility.Value - 1)
        .Locked = -chkPolyLocked.Value
      End With
      
      EditData = True
    End If
  End With
  Set myLinkObject = Nothing

  Unload frmProperties
End Function

Public Function EditDataM(Multi() As clsObject) As Boolean
  Dim I As Integer, Temp As clsPoint, PointObject As clsPoint
  Dim myObjType As ObjectTypeEnum

  ' Keeps track of which values need to be stored, and
  ' which are skipped

  ' When loading to the form, a false value means
  '   fill the control, true means make the control
  '   indeterminate (blank)
  ' When recording changes, a false value means store
  '   the value, true means do not store the value

  Dim IgnoreValue(24) As Boolean
  Dim Data As clsPoint

  Set Data = Multi(1)

  IgnoreValue(22) = Data.Parent.DotSpacing = 0
  IgnoreValue(23) = Data.Parent.LineWidth = 0
  
  myObjType = Data.Parent.ShapeType

  If myObjType = OT_FlatArea Then
    EditDataM = EditMulti(Multi)
    Exit Function
  End If
  
  For I = 2 To UBound(Multi)
    Set Temp = Multi(I)
    If Temp.Parent.ShapeType <> myObjType Then
      EditDataM = EditMulti(Multi)
      Exit Function
    End If
  Next I
  If myObjType = OT_TaxiwayLine Then
    EditDataM = frmFlatArea.EditDataTaxiLineM(Multi)
    Exit Function
  End If

  For I = 2 To UBound(Multi)
    Set Temp = Multi(I)
    With Temp
      IgnoreValue(1) = IgnoreValue(1) Or Data.X <> .X
      IgnoreValue(2) = IgnoreValue(2) Or Data.Y <> .Y
      IgnoreValue(3) = IgnoreValue(3) Or Data.Locked <> .Locked
      IgnoreValue(5) = IgnoreValue(5) Or Data.Z <> .Z
      IgnoreValue(6) = IgnoreValue(6) Or Data.Parent.Name <> .Parent.Name
      IgnoreValue(20) = IgnoreValue(20) Or Not Data.Parent Is .Parent
      IgnoreValue(4) = IgnoreValue(4) Or Data.PtType <> .PtType
      IgnoreValue(7) = IgnoreValue(7) Or Data.Parent.Texture <> .Parent.Texture
      IgnoreValue(12) = IgnoreValue(12) Or Data.Parent.Color <> .Parent.Color
      IgnoreValue(13) = IgnoreValue(13) Or Data.Parent.Layer <> .Parent.Layer
      IgnoreValue(8) = IgnoreValue(8) Or Data.Parent.MScale <> .Parent.MScale
      IgnoreValue(9) = IgnoreValue(9) Or Data.Parent.DotSpacing <> .Parent.DotSpacing
      IgnoreValue(22) = IgnoreValue(22) Or .Parent.DotSpacing = 0
      IgnoreValue(14) = IgnoreValue(14) Or Data.Parent.DotColor <> .Parent.DotColor
      IgnoreValue(10) = IgnoreValue(10) Or Data.Parent.LineWidth <> .Parent.LineWidth
      IgnoreValue(23) = IgnoreValue(23) Or .Parent.LineWidth = 0
      IgnoreValue(15) = IgnoreValue(15) Or Data.Parent.LineColor <> .Parent.LineColor
      IgnoreValue(17) = IgnoreValue(17) Or Data.Parent.Locked <> .Parent.Locked
      IgnoreValue(18) = IgnoreValue(18) Or Data.Parent.Complexity <> .Parent.Complexity
      
      IgnoreValue(11) = IgnoreValue(11) Or Data.Parent.Extra1 <> .Parent.Extra1
      IgnoreValue(16) = IgnoreValue(16) Or Data.Parent.Extra2 <> .Parent.Extra2
      IgnoreValue(21) = IgnoreValue(21) Or Data.Parent.Extra3 <> .Parent.Extra3
      IgnoreValue(24) = IgnoreValue(24) Or Data.Parent.Visibility <> .Parent.Visibility
    End With
    Set Temp = Nothing
  Next I

  MultiSelection = True
  Load frmProperties

  ' Fill Data
  With Data
    PrepareForm .Parent.ShapeType

    TabStrip1.Tabs(1).Caption = "  " & Lang.GetString(RES_Obj_Point) & "  "
    If Not IgnoreValue(1) Then Txts(1).Text = MeterToUser(.X)
    If Not IgnoreValue(2) Then Txts(2).Text = MeterToUser(.Y)
    Txts_Validate 1, True ' (Updates Latitude, Longitude)
    If Not IgnoreValue(3) Then
      chkLocked.Value = -.Locked
    Else
      chkLocked.Value = vbGrayed
    End If
    If Not IgnoreValue(5) Then Txts(5).Text = MeterToUser(.Z)
    If Not IgnoreValue(6) Then
      If Not IgnoreValue(20) Then
        Txts(6).Tag = .Parent.Caption(True)
        Txts(6).Text = .Parent.Name
      Else
        Txts(6).Tag = Lang.GetString(.Parent.ShapeType + RES_Obj_Header)
        Txts(6).Text = .Parent.Name
      End If
    Else
      Txts(6).Tag = Lang.GetString(.Parent.ShapeType + RES_Obj_Header)
    End If
    
    With .Parent
      If Not IgnoreValue(4) Then Cmbs(0).ListIndex = Data.PtType
      If Not IgnoreValue(7) Then
        Txts(7).Tag = .Texture
        Txts(7).Text = GetFileTitle(.Texture)
      Else
        Txts(7).Tag = "z"
      End If
      If Not IgnoreValue(12) Then
        picColor(0).Tag = .Color
        picColor(0).BackColor = ExtractColor(.Color)
      End If
      If Not IgnoreValue(13) Then
        If Options.FSVersion >= Version_FS2K Then
          Cmbs(1).ListIndex = FS2KLayerToIndex(.Layer)
        Else
          Cmbs(1).ListIndex = .Layer
        End If
      End If

      If Not IgnoreValue(8) Then Txts(8).Text = MeterToUser(.MScale, "##0.0#")
      
      If Not IgnoreValue(22) Then
        chkDot.Value = vbChecked
      ElseIf Not IgnoreValue(9) Then
        chkDot.Value = -(.DotSpacing > 0)
      Else
        chkDot.Value = vbGrayed
      End If
      
      Select Case .ShapeType
        Case OT_Polygon, OT_Line
          If Not IgnoreValue(9) Then
            If .DotSpacing = 0 Then
              Txts(9).Text = MeterToUser(60)
            Else
              Txts(9).Text = MeterToUser(.DotSpacing)
            End If
          End If
          
          If Not IgnoreValue(14) Then
            picColor(1).Tag = .DotColor
            picColor(1).BackColor = ExtractColor(.DotColor)
          End If
          
          Cmbs(2).Locked = True
          Cmbs(2).BackColor = vbButtonFace
        Case OT_Taxiway
          If Not IgnoreValue(9) Then
            If .DotSpacing = 0 Then
              Txts(9).Text = MeterToUser(60)
            Else
              Txts(9).Text = MeterToUser(.DotSpacing)
            End If
          End If
          
          If Not IgnoreValue(14) Then
            picColor(1).Tag = .DotColor
            picColor(1).BackColor = ExtractColor(.DotColor)
          End If
          
          If Not IgnoreValue(23) Then
            chkLine.Value = vbChecked
          ElseIf Not IgnoreValue(10) Then
            chkLine.Value = -(.LineWidth > 0)
          Else
            chkLine.Value = vbGrayed
          End If
          
          If Not IgnoreValue(10) Then
            If .LineWidth = 0 Then
              Txts(10).Text = MeterToUser(2)
            Else
              Txts(10).Text = MeterToUser(.LineWidth)
            End If
          End If
          
          If Not IgnoreValue(15) Then
            picColor(2).Tag = .LineColor
            picColor(2).BackColor = ExtractColor(.LineColor)
          End If
        Case OT_Road
          If Options.FSVersion >= Version_FS2K Then
            If Not IgnoreValue(21) Then Cmbs(2).ListIndex = .Extra3
          End If
      End Select
    
      If Not IgnoreValue(17) Then
        chkPolyLocked.Value = -.Locked
      Else
        chkPolyLocked.Value = vbGrayed
      End If
      If Not IgnoreValue(18) Then Cmbs(3).ListIndex = .Complexity
      
      If Not IgnoreValue(11) Then Txts(11).Text = MeterToUser(.Extra1)
      If Not IgnoreValue(16) Then
        chkNight.Value = -.Extra2
      Else
        chkNight.Value = vbGrayed
      End If
    
      If Not IgnoreValue(24) Then
        Txts(12).Text = MeterToUser(.Visibility)
        chkVisibility.Value = -(.Visibility = 0)
      Else
        chkVisibility.Value = vbGrayed
      End If
    End With
    
    Lang.PrepareForm Me
    chkDot_Click
    chkLine_Click
    TabStrip1_Click
    Txts_Change 6
    Txts_Change 7
    mTxtChanged = False
    mChanged = False
    If TabValue >= TabStrip1.Tabs.Count Then TabValue = 0
    Set TabStrip1.SelectedItem = TabStrip1.Tabs(TabValue + 1)
    Show vbModal, Screen.ActiveForm
  
    If mChanged Then
      IgnoreValue(1) = Txts(1).Text = ""
      IgnoreValue(2) = Txts(2).Text = ""
      IgnoreValue(3) = chkLocked.Value = vbGrayed
      IgnoreValue(5) = Txts(5).Text = ""
      IgnoreValue(6) = Txts(6).Text = ""
      
      IgnoreValue(4) = Cmbs(0).ListIndex = -1
      IgnoreValue(7) = Txts(7).Tag = "z"
      IgnoreValue(12) = picColor(0).Tag = ""
      IgnoreValue(13) = Cmbs(1).ListIndex = -1
      IgnoreValue(8) = Txts(8).Text = ""
      IgnoreValue(9) = chkDot.Value = vbGrayed
      
      Select Case .Parent.ShapeType
        Case OT_Polygon, OT_Line
          IgnoreValue(14) = IgnoreValue(9) Or picColor(1).Tag = ""
          IgnoreValue(9) = IgnoreValue(9) Or Txts(9).Text = ""
        Case OT_Taxiway
          IgnoreValue(14) = IgnoreValue(9) Or picColor(1).Tag = ""
          IgnoreValue(9) = IgnoreValue(9) Or Txts(9).Text = ""
                      
          IgnoreValue(10) = chkLine.Value = vbGrayed
          
          IgnoreValue(15) = IgnoreValue(10) Or picColor(2).Tag = ""
          IgnoreValue(10) = IgnoreValue(10) Or Txts(10).Text = ""
        Case OT_Road
          If Options.FSVersion >= Version_FS2K Then
            IgnoreValue(21) = Cmbs(2).ListIndex = -1
          End If
      End Select
      IgnoreValue(17) = chkPolyLocked.Value = vbGrayed
      IgnoreValue(18) = Cmbs(3).ListIndex = -1
      
      IgnoreValue(11) = Txts(11).Text = ""
      IgnoreValue(16) = chkNight.Value = vbGrayed
      IgnoreValue(24) = chkVisibility.Value = vbGrayed
      IgnoreValue(24) = IgnoreValue(24) Or Txts(12).Text = ""

      For I = 1 To UBound(Multi)
        Set Temp = Multi(I)
        With Temp
          If Not IgnoreValue(1) Then .X = mValueCache(1)
          If Not IgnoreValue(2) Then .Y = mValueCache(2)
          If Not IgnoreValue(3) Then .Locked = -chkLocked.Value
          If Not IgnoreValue(5) Then .Z = mValueCache(5)
          If Not IgnoreValue(6) Then .Parent.Name = Txts(6).Text
     
          If .Parent.ShapeType <> OT_Polygon Then
            If .ObjectIndex - 1 <> .Parent.NumPoints Then
              ' last point doesn't have this property
             If Not IgnoreValue(4) Then .PtType = Cmbs(0).ListIndex
            End If
          End If
          
          With .Parent
            If Not IgnoreValue(7) Then .Texture = Txts(7).Tag
            If Not IgnoreValue(12) Then .Color = picColor(0).Tag
            If Not IgnoreValue(13) Then
              If Options.FSVersion >= Version_FS2K Then
                .Layer = FS2KIndexToLayer(Cmbs(1).ListIndex)
              Else
                .Layer = Cmbs(1).ListIndex
              End If
            End If
            If Not IgnoreValue(8) Then .MScale = mValueCache(8)
            Select Case .ShapeType
              Case OT_Polygon, OT_Line
                If Not IgnoreValue(9) Then .DotSpacing = chkDot.Value * mValueCache(9)
                If Not IgnoreValue(14) Then .DotColor = picColor(1).Tag
              Case OT_Taxiway
                If Not IgnoreValue(9) Then .DotSpacing = chkDot.Value * mValueCache(9)
                If Not IgnoreValue(14) Then .DotColor = picColor(1).Tag
                If Not IgnoreValue(10) Then .LineWidth = chkLine.Value * mValueCache(10)
                If Not IgnoreValue(15) Then .LineColor = picColor(2).Tag
              Case OT_Road
                If Options.FSVersion >= Version_FS2K Then
                  If Not IgnoreValue(21) Then .Extra3 = Cmbs(2).ListIndex
                End If
            End Select
            If Not IgnoreValue(17) Then .Locked = -chkPolyLocked.Value
            If Not IgnoreValue(18) Then .Complexity = Cmbs(3).ListIndex
          End With
          
          If Not IgnoreValue(11) Then .Parent.Extra1 = mValueCache(11)
          If Not IgnoreValue(16) Then .Parent.Extra2 = -chkNight.Value
          If Not IgnoreValue(24) Then .Parent.Visibility = mValueCache(12) * Abs(chkVisibility.Value - 1)
        End With
      Next I
      EditDataM = True
    End If
  End With

  Unload frmProperties
  MultiSelection = False
End Function

Public Function EditDataMenuEntry(Data As clsMenuEntry) As Boolean
  Load frmProperties
  
  With Data
    PrepareForm OT_MenuEntry

    Caption = .Caption(True)
    Txts(0).Tag = RES_Rdo_Name
    Txts(0).Text = .Name
    Txts(1).Text = MeterToUser(.X)
    Txts(2).Text = MeterToUser(.Y)
    Txts_Validate 1, False ' (Updates Latitude, Longitude)
    chkLocked.Value = -.Locked
    Txts(5).Text = GeographicToUser(.Rotation)
    
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
      .Rotation = mValueCache(5)
      EditDataMenuEntry = True
    End If
  End With

  Unload frmProperties
End Function

Public Function EditMulti(Multi() As clsObject) As Boolean
  Dim I As Integer, J As Integer, _
    Data As clsObject, Temp As clsObject
  Dim IgnoreValue(6) As Boolean
  
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
      IgnoreValue(1) = IgnoreValue(1) Or Data.PositionX <> .PositionX
      IgnoreValue(2) = IgnoreValue(2) Or Data.PositionY <> .PositionY
      IgnoreValue(3) = IgnoreValue(3) Or Data.Locked <> .Locked
      IgnoreValue(5) = IgnoreValue(5) Or Data.Rotation <> .Rotation
      IgnoreValue(6) = IgnoreValue(6) Or Data.Complexity <> .Complexity
    End With
    Set Temp = Nothing
  Next I
  
  MultiSelection = True
  Load frmProperties
  
  PrepareForm -1
  
  ' Fill Data
  With Data
    If Not IgnoreValue(0) Then Txts(0).Text = .Name
    Txts(0).Tag = Lang.GetString(RES_Tab_Properties)
    Txts_Change 0
    If Not IgnoreValue(1) Then Txts(1).Text = MeterToUser(.PositionX)
    If Not IgnoreValue(2) Then Txts(2).Text = MeterToUser(.PositionY)
    Txts_Validate 1, False
    If Not IgnoreValue(3) Then
      chkLocked.Value = -.Locked
    Else
      chkLocked.Value = vbGrayed
    End If
    If Not IgnoreValue(5) Then Txts(5).Text = GeographicToUser(.Rotation)
    If Not IgnoreValue(6) Then Cmbs(0).ListIndex = .Complexity
  End With
  
  Lang.PrepareForm Me
  TabStrip1_Click
  mTxtChanged = False
  mChanged = False
  
  If TabValue >= TabStrip1.Tabs.Count Then TabValue = 0
  Set TabStrip1.SelectedItem = TabStrip1.Tabs(TabValue + 1)
  Show vbModal, Screen.ActiveForm
  If mChanged Then
    IgnoreValue(0) = Txts(0).Text = ""
    IgnoreValue(1) = Txts(1).Text = ""
    IgnoreValue(2) = Txts(2).Text = ""
    IgnoreValue(3) = chkLocked.Value = vbGrayed
    IgnoreValue(5) = Txts(5).Text = ""
    IgnoreValue(6) = Cmbs(0).ListIndex = -1

    For I = 1 To UBound(Multi)
      Set Temp = Multi(I)
      With Temp
        If Not IgnoreValue(0) Then .Name = Txts(0).Text
        If Not IgnoreValue(1) Then .PositionX = mValueCache(1)
        If Not IgnoreValue(2) Then .PositionY = mValueCache(2)
        If Not IgnoreValue(3) Then .Locked = -chkLocked.Value
        If Not IgnoreValue(5) Then .Rotation = mValueCache(5)
        If Not IgnoreValue(6) Then .Complexity = Cmbs(0).ListIndex
      End With
    Next I
    EditMulti = True
  End If
  
  Unload frmProperties
  MultiSelection = False
End Function

Private Function FS2KLayerToIndex(ByVal Layer As Byte) As Integer
  Select Case Layer
    Case Is < 34
      FS2KLayerToIndex = CInt(Layer / 4)
    Case 40
      FS2KLayerToIndex = 9
    Case 60
      FS2KLayerToIndex = 10
    Case Else
      FS2KLayerToIndex = 0
  End Select
End Function

Private Function FS2KIndexToLayer(ByVal Index As Integer) As Byte
  Select Case Index
    Case Is <= 8
      FS2KIndexToLayer = Index * 4
    Case 9
      FS2KIndexToLayer = 40
    Case 10
      FS2KIndexToLayer = 60
    Case Else
      FS2KIndexToLayer = 0
  End Select
End Function

Public Sub PrepareForm(ByVal TypeToPrepare As Long)
  Dim I As Integer, PointObject As clsPoint
  
  Select Case TypeToPrepare
    Case OT_Polygon
      Lang.AddItems Cmbs(0), RES_Pnt_NormalPoly, 2
      lbls(8).Tag = RES_Pnt_Lighting
      lbls(8).WhatsThisHelpID = RES_Pnt_Lighting
      Cmbs(0).WhatsThisHelpID = RES_Pnt_Lighting
      chkNight.Visible = False
      chkLine.Enabled = False
      Cmbs(2).AddItem Lang.GetString(RES_Shp_CmbNone)
      For I = 0 To Scenery.Count
        If Scenery(I).ObjectType = OT_Point Then
          Set PointObject = Scenery(I)
          If PointObject.ObjectIndex = 1 And PointObject.Parent.ShapeType = OT_Polygon Then _
            Cmbs(2).AddItem PointObject.Parent.Caption
            Cmbs(2).ItemData(Cmbs(2).ListCount - 1) = PointObject.SceneryIndex
        End If
      Next I
      SetEnabled Txts(0), False
      lbls(19).Visible = False
      Cmbs(2).Visible = False
      
      lbls(12).Top = lbls(12).Top + 100
      Cmbs(1).Top = Cmbs(1).Top + 100
      lbls(22).Top = lbls(22).Top - 150
      Cmbs(3).Top = Cmbs(3).Top - 150
    Case OT_Taxiway, OT_Road, OT_River
      Lang.AddItems Cmbs(0), RES_Pnt_NormalLine, 2
      lbls(20).Tag = RES_Shp_Width
      Txts(11).Tag = RES_Shp_Width
      lbls(20).WhatsThisHelpID = RES_Shp_Width
      Txts(11).WhatsThisHelpID = RES_Shp_Width
      If TypeToPrepare = OT_Road And Options.FSVersion >= Version_FS2K Then
        Lang.AddItems Cmbs(2), RES_Shp_Road1, 4
        lbls(19).Tag = RES_Shp_Type
        lbls(19).WhatsThisHelpID = RES_Shp_Type
        Cmbs(2).WhatsThisHelpID = RES_Shp_Type
      Else
        lbls(19).Visible = False
        Cmbs(2).Visible = False
        Cmbs(3).Top = Cmbs(2).Top
        lbls(22).Top = lbls(19).Top
      End If
      chkLine.Enabled = TypeToPrepare = OT_Taxiway
      chkDot.Enabled = TypeToPrepare = OT_Taxiway
      SetEnabled Txts(0), False
    Case OT_Line
      Lang.AddItems Cmbs(0), RES_Pnt_NormalLine, 2
      chkNight.Tag = RES_Shp_NightOnly
      chkNight.WhatsThisHelpID = RES_Shp_NightOnly
      lbls(19).Tag = RES_Shp_Object
      lbls(19).WhatsThisHelpID = RES_Shp_Object
      Cmbs(2).WhatsThisHelpID = RES_Shp_Object
      lbls(20).Tag = RES_Shp_LineObjWidth
      lbls(20).WhatsThisHelpID = RES_Shp_Width
      Txts(11).Tag = RES_Shp_LineObjWidth
      Txts(11).WhatsThisHelpID = RES_Shp_Width
      chkLine.Enabled = False
      Cmbs(2).AddItem Lang.GetString(RES_Shp_CmbNone)
      Cmbs(2).AddItem Lang.GetString(RES_Obj_Building)
      Cmbs(2).AddItem Lang.GetString(RES_Obj_Macro)
      Cmbs(2).Width = 1815
      cmdEdit.Visible = True
      SetEnabled Txts(0), False
    Case OT_MenuEntry
      TabStrip1.Tabs.Remove 2
      TabStrip1.Tabs(1).Tag = RES_Obj_MenuEntry
      lbls(8).Visible = False
      Cmbs(0).Visible = False
      lbls(7).Top = lbls(7).Top + 195
      Txts(5).Top = Txts(5).Top + 195
      lbls(7).Visible = True
      Txts(5).Visible = True
      lbls(7).Tag = RES_LBL_Rotation
      lbls(7).WhatsThisHelpID = RES_LBL_Rotation
      Txts(5).Tag = RES_LBL_Rotation
      Txts(5).WhatsThisHelpID = RES_LBL_Rotation
    Case -1
      TabStrip1.Tabs.Remove 2
      TabStrip1.Tabs(1).Tag = RES_Tab_Properties
      lbls(8).Tag = RES_LBL_Complexity
      Lang.AddItems Cmbs(0), RES_Complexity1, IIf(Options.FSVersion < Version_FS2K, 5, 6)
      lbls(7).Visible = True
      Txts(5).Visible = True
      lbls(7).Tag = RES_LBL_Rotation
      lbls(7).WhatsThisHelpID = RES_LBL_Rotation
      Txts(5).Tag = RES_LBL_Rotation
      Txts(5).WhatsThisHelpID = RES_LBL_Rotation
  End Select
  mType = TypeToPrepare
End Sub

Private Sub chkDot_Click()
  Dim Value As Boolean
  Value = chkDot.Value = vbChecked
  lbls(15).Enabled = Value
  SetEnabled Txts(9), Value
  lbls(16).Enabled = Value
  picColor(1).Enabled = Value
  If mType = OT_Line Then
    SetEnabled Txts(11), Not Value
  End If
End Sub

Private Sub chkLine_Click()
  Dim Value As Boolean
  Value = chkLine.Value = vbChecked
  lbls(17).Enabled = Value
  SetEnabled Txts(10), Value
  lbls(18).Enabled = Value
  picColor(2).Enabled = Value
End Sub

Private Sub chkLocked_Click()
  Dim Value As Boolean
  Value = Not -chkLocked.Value
  SetEnabled Txts(1), Value
  SetEnabled Txts(2), Value
  SetEnabled Txts(3), Value
  SetEnabled Txts(4), Value
End Sub

Private Sub chkVisibility_Click()
  SetEnabled Txts(12), Not -chkVisibility.Value
End Sub

Private Sub Cmbs_Click(Index As Integer)
  Dim PointObject As clsPoint, _
      ShapeObject As clsShape, _
      Num As Integer, _
      TempStr As String, _
      boolVal As Boolean
  
  Dim ToObject As clsShape, FromObject As clsObject
  
  If Index = 2 And Me.Visible Then
    Select Case mType
      Case OT_Polygon
        If Cmbs(2).ListIndex > -1 Then
          Num = Cmbs(2).ItemData(Cmbs(2).ListIndex)
          
          If Num > 0 Then
            Select Case Scenery(Num).ObjectType
              Case OT_Point
                Set PointObject = Scenery(Num)
                Set ShapeObject = PointObject.Parent
                Set ToObject = ShapeObject.LinkFromObj
                Set FromObject = ShapeObject
                Set PointObject = Nothing
                Set ShapeObject = Nothing
            End Select
            If Not ToObject Is Nothing Then
              MsgBoxEx Me, Lang.ResolveString(RES_ERR_Link, ToObject.Caption, FromObject.Caption), vbInformation, RES_ERR_Link
              Cmbs(2).ListIndex = 0
            End If
          End If
        End If
      Case OT_Line
        cmdEdit.Enabled = Cmbs(2).ListIndex > 0
        chkDot.Value = -cmdEdit.Enabled
      Case OT_Road
        If Cmbs(2).ListIndex > -1 Then
          Select Case Cmbs(2).ListIndex
            Case 0
              TempStr = "N " & AddDir(Options.FSPath, "Texture\asphalt.r8")
            Case 1
              TempStr = "N " & AddDir(Options.FSPath, "Texture\v_road_major.bmp")
            Case 2
              TempStr = "N " & AddDir(Options.FSPath, "Texture\v_road_minor.bmp")
            Case 3
              TempStr = "N " & AddDir(Options.FSPath, "Texture\v_railroad.bmp")
          End Select
          Txts(7).Tag = TempStr
          Txts(7).Text = GetFileTitle(TempStr)
        End If
    End Select
  End If
  Set FromObject = Nothing
  Set ToObject = Nothing
End Sub

Private Sub cmdCancel_Click()
  Hide
End Sub

Private Sub cmdEdit_Click()
  Dim boolVal As Boolean, NewObject As clsObject
  
  TabValue = 0
  Select Case Cmbs(2).ListIndex
    Case 1
      Set NewObject = Scenery.CreateNewInstance(OT_Building)
    Case 2
      Set NewObject = Scenery.CreateNewInstance(OT_Macro)
  End Select
  If Not NewObject Is Nothing Then
    If myLinkObject Is Nothing Then
      boolVal = NewObject.Add(0, 0)
    ElseIf NewObject.ObjectType = myLinkObject.ObjectType Then
      Set NewObject = myLinkObject
      boolVal = NewObject.EditProperties()
    Else
      boolVal = NewObject.Add(0, 0)
    End If
    If boolVal Then
      Set myLinkObject = Nothing
      Set myLinkObject = NewObject
      chkDot.Value = vbChecked
    End If
    Set NewObject = Nothing
  End If
End Sub

Private Sub cmdOK_Click()
  Dim I As Integer, Msg As String, Cancel As Boolean
  
  ' Validate event not fired when Enter key pressed
  ' bug workaround
  If TypeOf ActiveControl Is TextBox Then
    Txts_Validate ActiveControl.Index, Cancel
    If Cancel Then Exit Sub
  End If

  If mType = OT_MenuEntry Or mType = -1 Then
    For I = 0 To 5
      If Not Validate(Txts(I), Msg, mValueCache(I)) Then _
        GoTo ValidationError:
    Next I
  Else
    If TabStrip1.Tabs.Count > 1 Then
      For I = 1 To 5
        If Not Validate(Txts(I), Msg, mValueCache(I)) Then _
          GoTo ValidationError:
      Next I
    End If
    
    For I = 8 To 12
      If Not Validate(Txts(I), Msg, mValueCache(I)) Then
        If Not Txts(I).Locked Then _
          GoTo ValidationError:
      End If
    Next I
  End If
  If mType = OT_Line And Cmbs(2).ListIndex > 0 Then
    If chkDot.Value * mValueCache(9) = 0 Then
      Msg = Lang.GetString(RES_ERR_DotSpacingReq)
      I = 9
      GoTo ValidationError:
    End If
  End If
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
  
  Lang.AddItems Cmbs(3), RES_Complexity1, IIf(Options.FSVersion < Version_FS2K, 5, 6)
  
  If Options.FSVersion >= Version_FS2K Then
    For I = 0 To 10
      Cmbs(1).AddItem Lang.GetString(RES_Shp_Layer1 + I)
    Next I
  Else
    For I = 0 To 63
      Cmbs(1).AddItem CStr(I)
    Next I
  End If
  mFocus = -1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then Cancel = True: Hide
End Sub

Private Sub picBackground_Paint()
  If mFocus > -1 Then DrawFocus mFocus, True
End Sub

Private Sub picColor_GotFocus(Index As Integer)
  DrawFocus Index, True
  mFocus = Index
End Sub

Private Sub picColor_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeySpace
      picColor_MouseDown Index, vbLeftButton, 0, 0, 0
    Case vbKeyDelete, vbKeyBack
      picColor(Index).Tag = 0
      picColor(Index).BackColor = ExtractColor(0)
  End Select
End Sub

Private Sub picColor_LostFocus(Index As Integer)
  DrawFocus Index, False
  mFocus = -1
End Sub

Private Sub picColor_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim Color As Long
  Color = Val(picColor(Index).Tag)
  If frmColor.EditData(Color) Then
    picColor(Index).Tag = Color
    picColor(Index).BackColor = ExtractColor(Color)
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
          Case 2: Txts(6).SetFocus
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
      If TempStr = "" And mType <> OT_MenuEntry Then
        Caption = Txts(0).Tag
      Else
        Caption = TempStr
      End If
    Case 6
      TempStr = Txts(6).Text
      If TempStr = "" Then
        TempStr = Txts(6).Tag
        Caption = TempStr
      Else
        Caption = ObjectNames(mType) & ": " & TempStr
      End If
      
      If TabStrip1.Tabs.Count > 1 Then
        TabStrip1.Tabs(2).Caption = "  " & TempStr & "  "
      Else
        TabStrip1.Tabs(1).Caption = "  " & TempStr & "  "
      End If
    
    Case 7
      SetEnabled Txts(8), (Txts(7).Tag <> "")
    Case 11
      If mType <> OT_Polygon Then
        UserToMeter Txts(11).Text, 0, Conv, 0, 0
        If Conv > 0 Then
          If ValEx(Txts(11).Text) * Conv <= 0.1 Then
            Txts(7).BackColor = vbButtonFace
          Else
            Txts(7).BackColor = vbWindowBackground
          End If
        Else
          Txts(7).BackColor = vbWindowBackground
        End If
      End If
  End Select
  mTxtChanged = True
End Sub

Private Sub Txts_Click(Index As Integer)
  Dim X As String
  If Index = 7 Then
    X = Txts(7).Tag
    If frmTexture.EditData(X, Tex_File) Then
      Txts(7).Tag = X
      Txts(7).Text = GetFileTitle(X)
      Txts(7).SelStart = Len(Txts(7).Text)
    End If
  End If
End Sub

Private Sub Txts_GotFocus(Index As Integer)
  Select Case Index
    Case 1, 2, 5, 8, 9, 10, 11, 12
      SmartSelectText Txts(Index)
    Case Else
      If Index = 0 Or Index = 6 Then ReturnSymbol 0, 0, 2
      SelectText Txts(Index)
  End Select
  mTxtChanged = False
End Sub

Private Sub Txts_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Dim Result As Integer
    
  If Index = 0 Or Index = 6 Then
    If Shift > 1 Then
      Result = ReturnSymbol(KeyCode, Shift)
      If Result > 0 Then Txts(Index).SelText = Chr$(Result): KeyCode = 0
    End If
  ElseIf Index = 7 Then
    Select Case KeyCode
      Case vbKeySpace
        Txts_Click 7
      Case vbKeyDelete, vbKeyBack
        Txts(7).Tag = ""
        Txts(7).Text = ""
    End Select
  End If
End Sub

Private Sub Txts_KeyPress(Index As Integer, KeyAscii As Integer)
  If Index = 0 Or Index = 6 Then
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
