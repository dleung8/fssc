VERSION 5.00
Begin VB.Form frmExclusion 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   ClipControls    =   0   'False
   Icon            =   "Exclusion.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox Txts 
      Height          =   285
      Index           =   8
      Left            =   1560
      TabIndex        =   22
      Tag             =   "1803"
      Top             =   4590
      Visible         =   0   'False
      WhatsThisHelpID =   1803
      Width           =   1095
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   4170
      Visible         =   0   'False
      WhatsThisHelpID =   1802
      Width           =   2415
   End
   Begin VB.TextBox Txts 
      Height          =   285
      Index           =   7
      Left            =   1560
      TabIndex        =   18
      Tag             =   "1048"
      Top             =   3780
      Visible         =   0   'False
      WhatsThisHelpID =   1048
      Width           =   1095
   End
   Begin VB.CheckBox chks 
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   26
      Tag             =   "1813"
      Top             =   4560
      WhatsThisHelpID =   1810
      Width           =   2655
   End
   Begin VB.CheckBox chks 
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   25
      Tag             =   "1812"
      Top             =   4320
      WhatsThisHelpID =   1810
      Width           =   2655
   End
   Begin VB.CheckBox chks 
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   24
      Tag             =   "1811"
      Top             =   4080
      WhatsThisHelpID =   1810
      Width           =   2655
   End
   Begin VB.CheckBox chks 
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   23
      Tag             =   "1810"
      Top             =   3840
      WhatsThisHelpID =   1810
      Width           =   2655
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   28
      Tag             =   "1031"
      Top             =   5040
      WhatsThisHelpID =   1031
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   27
      Tag             =   "1030"
      Top             =   5040
      WhatsThisHelpID =   1030
      Width           =   1215
   End
   Begin VB.TextBox Txts 
      Height          =   285
      Index           =   6
      Left            =   1560
      TabIndex        =   16
      Tag             =   "1801"
      Top             =   3390
      WhatsThisHelpID =   1800
      Width           =   1095
   End
   Begin VB.TextBox Txts 
      Height          =   285
      Index           =   5
      Left            =   1560
      TabIndex        =   14
      Tag             =   "1800"
      Top             =   3000
      WhatsThisHelpID =   1800
      Width           =   1095
   End
   Begin VB.TextBox Txts 
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   12
      Top             =   2430
      WhatsThisHelpID =   1045
      Width           =   1935
   End
   Begin VB.TextBox Txts 
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   10
      Top             =   2040
      WhatsThisHelpID =   1045
      Width           =   1935
   End
   Begin VB.TextBox Txts 
      Height          =   285
      Index           =   2
      Left            =   3360
      TabIndex        =   7
      Tag             =   "1044"
      Top             =   1320
      WhatsThisHelpID =   1042
      Width           =   1095
   End
   Begin VB.TextBox Txts 
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   5
      Tag             =   "1043"
      Top             =   1320
      WhatsThisHelpID =   1042
      Width           =   1095
   End
   Begin VB.CheckBox chkLocked 
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Tag             =   "1041"
      Top             =   720
      WhatsThisHelpID =   1041
      Width           =   2655
   End
   Begin VB.TextBox Txts 
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      WhatsThisHelpID =   1040
      Width           =   3135
   End
   Begin VB.Label lbls 
      Height          =   390
      Index           =   11
      Left            =   240
      TabIndex        =   21
      Tag             =   "1803"
      Top             =   4620
      Visible         =   0   'False
      WhatsThisHelpID =   1803
      Width           =   1185
   End
   Begin VB.Label lbls 
      Height          =   390
      Index           =   10
      Left            =   240
      TabIndex        =   19
      Tag             =   "1802"
      Top             =   4215
      Visible         =   0   'False
      WhatsThisHelpID =   1802
      Width           =   1185
   End
   Begin VB.Label lbls 
      Height          =   390
      Index           =   9
      Left            =   240
      TabIndex        =   17
      Tag             =   "1048"
      Top             =   3810
      Visible         =   0   'False
      WhatsThisHelpID =   1048
      Width           =   1185
   End
   Begin VB.Label lbls 
      Height          =   390
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Tag             =   "1040"
      Top             =   270
      WhatsThisHelpID =   1040
      Width           =   945
   End
   Begin VB.Label lbls 
      Height          =   390
      Index           =   8
      Left            =   240
      TabIndex        =   15
      Tag             =   "1801"
      Top             =   3420
      WhatsThisHelpID =   1800
      Width           =   1185
   End
   Begin VB.Label lbls 
      Height          =   390
      Index           =   7
      Left            =   240
      TabIndex        =   13
      Tag             =   "1800"
      Top             =   3030
      WhatsThisHelpID =   1800
      Width           =   1185
   End
   Begin VB.Label lbls 
      Height          =   390
      Index           =   6
      Left            =   840
      TabIndex        =   11
      Tag             =   "1047"
      Top             =   2460
      WhatsThisHelpID =   1045
      Width           =   1065
   End
   Begin VB.Label lbls 
      Height          =   390
      Index           =   5
      Left            =   840
      TabIndex        =   9
      Tag             =   "1046"
      Top             =   2070
      WhatsThisHelpID =   1045
      Width           =   1065
   End
   Begin VB.Label lbls 
      Height          =   285
      Index           =   4
      Left            =   600
      TabIndex        =   8
      Tag             =   "1045"
      Top             =   1755
      WhatsThisHelpID =   1045
      Width           =   2175
   End
   Begin VB.Label lbls 
      Height          =   390
      Index           =   3
      Left            =   2880
      TabIndex        =   6
      Tag             =   "1044"
      Top             =   1350
      WhatsThisHelpID =   1042
      Width           =   345
   End
   Begin VB.Label lbls 
      Height          =   390
      Index           =   2
      Left            =   840
      TabIndex        =   4
      Tag             =   "1043"
      Top             =   1350
      WhatsThisHelpID =   1042
      Width           =   345
   End
   Begin VB.Label lbls 
      Height          =   285
      Index           =   1
      Left            =   600
      TabIndex        =   3
      Tag             =   "1042"
      Top             =   1035
      WhatsThisHelpID =   1042
      Width           =   2175
   End
End
Attribute VB_Name = "frmExclusion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mChanged As Boolean, mTxtChanged As Boolean

Private mValueCache(8) As Single

Public Function EditData(Data As clsExclusion) As Boolean
  Dim Offset As Single

  Load frmExclusion

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
    Txts(7).Text = 0
    Txts(8).Text = 0

    chks(0).Value = -((.Exclusion And 1) > 0)
    chks(1).Value = -((.Exclusion And 2) > 0)
    chks(2).Value = -((.Exclusion And 4) > 0)
    chks(3).Value = -((.Exclusion And 8) > 0)

    CenterForm Me
    Lang.PrepareForm Me
    mTxtChanged = False
    mChanged = False
    Show vbModal, Screen.ActiveForm

    If mChanged Then
      .Name = Txts(0).Text
      .X = mValueCache(1)
      .Y = mValueCache(2)
      .Locked = -chkLocked.Value
      .Horz = mValueCache(5)
      .Vert = mValueCache(6)

      .Exclusion = chks(0).Value * 1 + _
                   chks(1).Value * 2 + _
                   chks(2).Value * 4 + _
                   chks(3).Value * 8

      EditData = True
    End If
  End With

  Unload frmExclusion
End Function

Public Function EditDataSurface(Data As clsSurfaceArea) As Boolean
  Dim Offset As Single

  Load frmExclusion
  
  chks(0).Visible = False
  chks(1).Visible = False
  chks(2).Visible = False
  chks(3).Visible = False
  lbls(9).Visible = True
  lbls(10).Visible = True
  lbls(11).Visible = True
  Txts(7).Visible = True
  Txts(8).Visible = True
  cmbType.Visible = True
  Lang.AddItems cmbType, RES_Suf_Type, 4

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
    Txts(7).Text = GeographicToUser(.Rotation)
    Txts(8).Text = MeterToUser(.Height)
    cmbType.ListIndex = .SurfaceType

    CenterForm Me
    Lang.PrepareForm Me
    mTxtChanged = False
    mChanged = False
    Show vbModal, Screen.ActiveForm

    If mChanged Then
      .Name = Txts(0).Text
      .X = mValueCache(1)
      .Y = mValueCache(2)
      .Locked = -chkLocked.Value
      .Horz = mValueCache(5)
      .Vert = mValueCache(6)
      .Rotation = mValueCache(7)
      .SurfaceType = cmbType.ListIndex
      .Height = mValueCache(8)

      EditDataSurface = True
    End If
  End With

  Unload frmExclusion
End Function

Private Sub chkLocked_Click()
  Dim Value As Boolean
  Value = Not -chkLocked.Value
  SetEnabled Txts(1), Value
  SetEnabled Txts(2), Value
  SetEnabled Txts(3), Value
  SetEnabled Txts(4), Value
End Sub

Private Sub cmbType_Click()
  SetEnabled Txts(8), (cmbType.ListIndex = 0)
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

  For I = 0 To 8
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
  DialogMenus Me

  If Options.FSVersion >= Version_FS2K Then
    chks(3).Enabled = False
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then Cancel = True: Hide
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
    Case 1, 2, 5, 6, 7, 8
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
  Select Case Index
    Case 0
      KeyAscii = ReturnSymbol(KeyAscii, 0, 1)
  End Select
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
