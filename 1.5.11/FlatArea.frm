VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmFlatArea 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   ClipControls    =   0   'False
   Icon            =   "FlatArea.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   4080
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   26
      Tag             =   "1031"
      Top             =   4080
      WhatsThisHelpID =   1031
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   25
      Tag             =   "1030"
      Top             =   4080
      WhatsThisHelpID =   1030
      Width           =   1215
   End
   Begin VB.Frame TabFrame 
      BorderStyle     =   0  'None
      Height          =   3255
      Index           =   2
      Left            =   360
      TabIndex        =   16
      Top             =   600
      Width           =   4335
      Begin VB.CheckBox chkLighted 
         Height          =   195
         Left            =   0
         TabIndex        =   23
         Tag             =   "1324"
         Top             =   1560
         WhatsThisHelpID =   1324
         Width           =   2895
      End
      Begin VB.CheckBox chkPolyLocked 
         Height          =   195
         Left            =   0
         TabIndex        =   24
         Tag             =   "1317"
         Top             =   1920
         WhatsThisHelpID =   1317
         Width           =   3615
      End
      Begin VB.ComboBox cmbType 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   765
         WhatsThisHelpID =   1314
         Width           =   3015
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   6
         Left            =   1200
         TabIndex        =   22
         Tag             =   "1318"
         Top             =   1170
         WhatsThisHelpID =   1318
         Width           =   1095
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   5
         Left            =   1200
         TabIndex        =   18
         Top             =   120
         WhatsThisHelpID =   1040
         Width           =   3015
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   9
         Left            =   0
         TabIndex        =   21
         Tag             =   "1318"
         Top             =   1200
         WhatsThisHelpID =   1318
         Width           =   1065
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   8
         Left            =   0
         TabIndex        =   19
         Tag             =   "1314"
         Top             =   810
         WhatsThisHelpID =   1314
         Width           =   1065
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   7
         Left            =   0
         TabIndex        =   17
         Tag             =   "1040"
         Top             =   150
         WhatsThisHelpID =   1040
         Width           =   1065
      End
   End
   Begin VB.Frame TabFrame 
      BorderStyle     =   0  'None
      Height          =   3255
      Index           =   1
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   4440
      Begin VB.ComboBox Cmbs 
         Height          =   315
         Index           =   0
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2760
         WhatsThisHelpID =   1951
         Width           =   2895
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   4
         Left            =   1800
         TabIndex        =   13
         Top             =   2310
         WhatsThisHelpID =   1043
         Width           =   1935
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   3
         Left            =   1800
         TabIndex        =   11
         Top             =   1920
         WhatsThisHelpID =   1043
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
         Index           =   10
         Left            =   0
         TabIndex        =   14
         Tag             =   "1951"
         Top             =   2805
         WhatsThisHelpID =   1951
         Width           =   1185
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   6
         Left            =   600
         TabIndex        =   12
         Tag             =   "1047"
         Top             =   2340
         WhatsThisHelpID =   1043
         Width           =   1065
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   5
         Left            =   600
         TabIndex        =   10
         Tag             =   "1046"
         Top             =   1950
         WhatsThisHelpID =   1043
         Width           =   1065
      End
      Begin VB.Label lbls 
         Height          =   285
         Index           =   4
         Left            =   360
         TabIndex        =   9
         Tag             =   "1045"
         Top             =   1635
         WhatsThisHelpID =   1043
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
      Height          =   3855
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   6800
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "A"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "B"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmFlatArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mChanged As Boolean, mTxtChanged As Boolean, _
        mFocus As Integer, mType As ObjectTypeEnum

Private mValueCache(6) As Single

Public Function EditData(Data As clsPoint) As Boolean
  Load frmFlatArea

  With Data
    PrepareForm OT_FlatArea
    
    SetEnabled Txts(0), False
    
    ' If there are no points (i.e. when shape is first
    ' created), then don't show the Point tab
    If .Parent.NumPoints = -1 Then
      TabStrip1.Tabs.Remove 1
    Else
      TabStrip1.Tabs(1).Caption = "  " & .Caption(True) & "  "
      Txts(0).Text = .Caption(True)
      Txts(1).Text = MeterToUser(.X)
      Txts(2).Text = MeterToUser(.Y)
      chkLocked.Value = -.Locked
      Txts_Validate 1, True ' (Updates Latitude, Longitude)
    End If

    With .Parent
      Txts(5).Tag = .Caption(True)
      Txts(5).Text = .Name
      cmbType.ListIndex = .Extra3
      Txts(6).Text = MeterToUser(.Extra1)
      chkPolyLocked.Value = -.Locked
    End With
    
    Lang.PrepareForm Me
    TabStrip1_Click
    Txts_Change 5
    mTxtChanged = False
    mChanged = False
    If TabValue >= TabStrip1.Tabs.Count Then TabValue = 0
    Set TabStrip1.SelectedItem = TabStrip1.Tabs(TabValue + 1)
    Show vbModal, Screen.ActiveForm

    If mChanged Then
      If .Parent.NumPoints > -1 Then
        .X = mValueCache(1)
        .Y = mValueCache(2)
        .Locked = -chkLocked.Value
      End If

      With .Parent
        .Name = Txts(5).Text
        .Extra3 = cmbType.ListIndex
        .Extra1 = mValueCache(6)
        .Locked = -chkPolyLocked.Value
      End With

      EditData = True
    End If
  End With

  Unload frmFlatArea
End Function

Public Function EditDataTaxiLine(Data As clsPoint) As Boolean
  Load frmFlatArea

  With Data
    PrepareForm OT_TaxiwayLine
    SetEnabled Txts(0), False
    
    ' If there are no points (i.e. when shape is first
    ' created), then don't show the Point tab
    If .Parent.NumPoints = -1 Then
      TabStrip1.Tabs.Remove 1
    Else
      TabStrip1.Tabs(1).Caption = "  " & .Caption(True) & "  "
      Txts(0).Text = .Caption(True)
      Txts(1).Text = MeterToUser(.X)
      Txts(2).Text = MeterToUser(.Y)
      chkLocked.Value = -.Locked
      Txts_Validate 1, True ' (Updates Latitude, Longitude)
    
      If .ObjectIndex - 1 = .Parent.NumPoints Then
        ' This option doesn't make sense for the last point
        Cmbs(0).Locked = True
        Cmbs(0).BackColor = vbButtonFace
      End If
      Cmbs(0).ListIndex = .PtType
    End If

    With .Parent
      Txts(5).Tag = .Caption(True)
      Txts(5).Text = .Name
      cmbType.ListIndex = .Extra3
      Txts(6).Text = MeterToUser(.Extra1)
      chkLighted.Value = -.Extra2
      chkPolyLocked.Value = -.Locked
    End With

    Lang.PrepareForm Me
    TabStrip1_Click
    Txts_Change 5
    mTxtChanged = False
    mChanged = False
    If TabValue >= TabStrip1.Tabs.Count Then TabValue = 0
    Set TabStrip1.SelectedItem = TabStrip1.Tabs(TabValue + 1)
    Show vbModal, Screen.ActiveForm

    If mChanged Then
      If .Parent.NumPoints > -1 Then
        .X = mValueCache(1)
        .Y = mValueCache(2)
        .Locked = -chkLocked.Value
        .PtType = Cmbs(0).ListIndex
      End If

      With .Parent
        .Name = Txts(5).Text
        .Extra3 = cmbType.ListIndex
        .Extra1 = mValueCache(6)
        .Extra2 = -chkLighted.Value
        .Locked = -chkPolyLocked.Value
      End With

      EditDataTaxiLine = True
    End If
  End With

  Unload frmFlatArea
End Function

Public Function EditDataTaxiLineM(Multi() As clsObject) As Boolean
  ' Keeps track of which values need to be stored, and
  ' which are skipped

  ' When loading to the form, a false value means
  '   fill the control, true means make the control
  '   indeterminate (blank)
  ' When recording changes, a false value means store
  '   the value, true means do not store the value

  Dim IgnoreValue(10) As Boolean
  Dim Data As clsPoint
  Dim I As Integer
  Dim Temp As clsPoint
  
  Set Data = Multi(1)
  
  For I = 2 To UBound(Multi)
    Set Temp = Multi(I)
    With Temp
      IgnoreValue(1) = IgnoreValue(1) Or Data.X <> .X
      IgnoreValue(2) = IgnoreValue(2) Or Data.Y <> .Y
      IgnoreValue(3) = IgnoreValue(3) Or Data.Locked <> .Locked
      
      IgnoreValue(5) = IgnoreValue(5) Or Data.Parent.Name <> .Parent.Name
      IgnoreValue(4) = IgnoreValue(4) Or Not Data.Parent Is .Parent
      
      IgnoreValue(6) = IgnoreValue(6) Or Data.Parent.Extra3 <> .Parent.Extra3
      IgnoreValue(7) = IgnoreValue(7) Or Data.Parent.Extra1 <> .Parent.Extra1
      IgnoreValue(8) = IgnoreValue(8) Or Data.Parent.Extra2 <> .Parent.Extra2
      IgnoreValue(9) = IgnoreValue(9) Or Data.Parent.Locked <> .Parent.Locked
    
      IgnoreValue(10) = IgnoreValue(10) Or Data.PtType <> .PtType
    End With
    Set Temp = Nothing
  Next I
    
  MultiSelection = True
  Load frmFlatArea
  PrepareForm OT_TaxiwayLine
  SetEnabled Txts(0), False
    
  With Data
    TabStrip1.Tabs(1).Caption = "  " & Lang.GetString(RES_Obj_Point) & "  "
    If Not IgnoreValue(1) Then Txts(1).Text = MeterToUser(.X)
    If Not IgnoreValue(2) Then Txts(2).Text = MeterToUser(.Y)
    Txts_Validate 1, True ' (Updates Latitude, Longitude)
    If Not IgnoreValue(3) Then
      chkLocked.Value = -.Locked
    Else
      chkLocked.Value = vbGrayed
    End If
    
    If Not IgnoreValue(10) Then Cmbs(0).ListIndex = .PtType
    
    If Not IgnoreValue(5) Then
      If Not IgnoreValue(4) Then
        Txts(5).Tag = .Parent.Caption(True)
      Else
        Txts(5).Tag = Lang.GetString(.Parent.ShapeType + RES_Obj_Header)
      End If
      Txts(5).Text = .Parent.Name
    Else
      Txts(5).Tag = Lang.GetString(.Parent.ShapeType + RES_Obj_Header)
    End If
        
    With .Parent
      If Not IgnoreValue(6) Then cmbType.ListIndex = .Extra3
      If Not IgnoreValue(7) Then Txts(6).Text = MeterToUser(.Extra1)
      If Not IgnoreValue(8) Then
        chkLighted.Value = -.Extra2
      Else
        chkLighted.Value = vbGrayed
      End If
      If Not IgnoreValue(9) Then
        chkPolyLocked.Value = -.Locked
      Else
        chkPolyLocked.Value = vbGrayed
      End If
    End With

    Lang.PrepareForm Me
    TabStrip1_Click
    Txts_Change 5
    mTxtChanged = False
    mChanged = False
    If TabValue >= TabStrip1.Tabs.Count Then TabValue = 0
    Set TabStrip1.SelectedItem = TabStrip1.Tabs(TabValue + 1)
    Show vbModal, Screen.ActiveForm

    If mChanged Then
      IgnoreValue(1) = Txts(1).Text = ""
      IgnoreValue(2) = Txts(2).Text = ""
      IgnoreValue(3) = chkLocked.Value = vbGrayed
      
      IgnoreValue(10) = Cmbs(0).ListIndex = -1

      IgnoreValue(5) = Txts(5).Text = ""
      
      IgnoreValue(6) = cmbType.ListIndex = -1
      IgnoreValue(7) = Txts(6).Text = ""
      IgnoreValue(8) = chkLighted.Value = vbGrayed
      IgnoreValue(9) = chkPolyLocked.Value = vbGrayed
      
      For I = 1 To UBound(Multi)
        Set Temp = Multi(I)
        With Temp
          If Not IgnoreValue(1) Then .X = mValueCache(1)
          If Not IgnoreValue(2) Then .Y = mValueCache(2)
          If Not IgnoreValue(3) Then .Locked = -chkLocked.Value
                    
          If .ObjectIndex - 1 <> .Parent.NumPoints Then
            ' last point doesn't have this property
           If Not IgnoreValue(10) Then .PtType = Cmbs(0).ListIndex
          End If
          
          If Not IgnoreValue(5) Then .Parent.Name = Txts(5).Text
          If Not IgnoreValue(6) Then .Parent.Extra3 = cmbType.ListIndex
          If Not IgnoreValue(7) Then .Parent.Extra1 = mValueCache(6)
          If Not IgnoreValue(8) Then .Parent.Extra2 = -chkLighted.Value
          If Not IgnoreValue(9) Then .Parent.Locked = -chkPolyLocked.Value
        End With
        Set Temp = Nothing
      Next I
      
      EditDataTaxiLineM = True
    End If
  End With

  Unload frmFlatArea
  MultiSelection = False
End Function

Private Sub PrepareForm(ByVal ObjType As ObjectTypeEnum)
  mType = ObjType
  Select Case ObjType
    Case OT_FlatArea
      chkLighted.Visible = False
      Lang.AddItems cmbType, RES_Shp_CmbFlat, 2
      Cmbs(0).Visible = False
      lbls(10).Visible = False
    Case OT_TaxiwayLine
      lbls(8).Tag = RES_Shp_TaxiType
      lbls(9).Tag = RES_Shp_ArcRadius
      Txts(6).Tag = RES_Shp_ArcRadius
      
      cmbType.WhatsThisHelpID = RES_Shp_TaxiType
      lbls(8).WhatsThisHelpID = RES_Shp_TaxiType
      lbls(9).WhatsThisHelpID = RES_Shp_ArcRadius
      Txts(6).WhatsThisHelpID = RES_Shp_ArcRadius
      Lang.AddItems cmbType, RES_Shp_TaxiwayLine1, 7

      Lang.AddItems Cmbs(0), RES_Pnt_NormalLine, 2
  End Select
End Sub

Private Sub chkLocked_Click()
  Dim Value As Boolean
  Value = Not -chkLocked.Value
  SetEnabled Txts(1), Value
  SetEnabled Txts(2), Value
  SetEnabled Txts(3), Value
  SetEnabled Txts(4), Value
End Sub

Private Sub cmbType_Click()
  Dim Value As Boolean
  If mType = OT_TaxiwayLine Then
    Value = (cmbType.ListIndex <= 3)
    chkLighted.Enabled = Value
    SetEnabled Txts(6), Value
  End If
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

  If TabStrip1.Tabs.Count > 1 Then
    For I = 1 To 2
      If Not Validate(Txts(I), Msg, mValueCache(I)) Then _
        GoTo ValidationError:
    Next I
  End If
  I = 6
  If Not Validate(Txts(6), Msg, mValueCache(6)) Then _
    GoTo ValidationError:
  
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
          Case 2: Txts(5).SetFocus
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
  If Index = 5 Then
    TempStr = Txts(5).Text
    If TempStr = "" Then TempStr = Txts(5).Tag
    
    If TabStrip1.Tabs.Count > 1 Then
      TabStrip1.Tabs(2).Caption = "  " & TempStr & "  "
    Else
      TabStrip1.Tabs(1).Caption = "  " & TempStr & "  "
    End If
    Caption = TempStr
  End If
  mTxtChanged = True
End Sub

Private Sub Txts_GotFocus(Index As Integer)
  Select Case Index
    Case 1, 2, 6
      SmartSelectText Txts(Index)
    Case Else
      If Index = 0 Or Index = 5 Then ReturnSymbol 0, 0, 2
      SelectText Txts(Index)
  End Select
  mTxtChanged = False
End Sub

Private Sub Txts_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Dim Result As Integer
    
  If Index = 0 Or Index = 5 Then
    If Shift > 1 Then
      Result = ReturnSymbol(KeyCode, Shift)
      If Result > 0 Then Txts(Index).SelText = Chr$(Result): KeyCode = 0
    End If
  End If
End Sub

Private Sub Txts_KeyPress(Index As Integer, KeyAscii As Integer)
  If Index = 0 Or Index = 5 Then
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
