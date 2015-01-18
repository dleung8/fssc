VERSION 5.00
Begin VB.Form frmBeacon 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   ClipControls    =   0   'False
   Icon            =   "Beacon.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CheckBox chks 
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   22
      Tag             =   "1623"
      Top             =   5040
      WhatsThisHelpID =   1623
      Width           =   4215
   End
   Begin VB.CheckBox chks 
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   21
      Tag             =   "1622"
      Top             =   4800
      WhatsThisHelpID =   1622
      Width           =   4215
   End
   Begin VB.CheckBox chks 
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   20
      Tag             =   "1621"
      Top             =   4560
      WhatsThisHelpID =   1621
      Width           =   4215
   End
   Begin VB.CheckBox chks 
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   19
      Tag             =   "1620"
      Top             =   4200
      WhatsThisHelpID =   1620
      Width           =   4215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   24
      Tag             =   "1031"
      Top             =   5400
      WhatsThisHelpID =   1031
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   23
      Tag             =   "1030"
      Top             =   5400
      WhatsThisHelpID =   1030
      Width           =   1215
   End
   Begin VB.TextBox Txts 
      Height          =   285
      Index           =   7
      Left            =   1320
      TabIndex        =   18
      Tag             =   "1606"
      Top             =   3750
      WhatsThisHelpID =   1606
      Width           =   1095
   End
   Begin VB.TextBox Txts 
      Height          =   285
      Index           =   6
      Left            =   1320
      TabIndex        =   16
      Tag             =   "1604"
      Top             =   3360
      WhatsThisHelpID =   1603
      Width           =   1095
   End
   Begin VB.TextBox Txts 
      Height          =   285
      Index           =   5
      Left            =   1320
      MaxLength       =   5
      TabIndex        =   3
      Tag             =   "1601"
      Top             =   630
      WhatsThisHelpID =   1601
      Width           =   1095
   End
   Begin VB.TextBox Txts 
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   14
      Top             =   2790
      WhatsThisHelpID =   1045
      Width           =   1935
   End
   Begin VB.TextBox Txts 
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   12
      Top             =   2400
      WhatsThisHelpID =   1045
      Width           =   1935
   End
   Begin VB.TextBox Txts 
      Height          =   285
      Index           =   2
      Left            =   3360
      TabIndex        =   9
      Tag             =   "1044"
      Top             =   1680
      WhatsThisHelpID =   1042
      Width           =   1095
   End
   Begin VB.TextBox Txts 
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   7
      Tag             =   "1043"
      Top             =   1680
      WhatsThisHelpID =   1042
      Width           =   1095
   End
   Begin VB.CheckBox chkLocked 
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Tag             =   "1041"
      Top             =   1080
      WhatsThisHelpID =   1041
      Width           =   2655
   End
   Begin VB.TextBox Txts 
      Height          =   285
      Index           =   0
      Left            =   1320
      MaxLength       =   24
      TabIndex        =   1
      Tag             =   "1600"
      Top             =   240
      WhatsThisHelpID =   1600
      Width           =   3135
   End
   Begin VB.Label lbls 
      Height          =   390
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Tag             =   "1600"
      Top             =   270
      WhatsThisHelpID =   1600
      Width           =   1065
   End
   Begin VB.Label lbls 
      Height          =   390
      Index           =   8
      Left            =   240
      TabIndex        =   17
      Tag             =   "1606"
      Top             =   3780
      WhatsThisHelpID =   1606
      Width           =   945
   End
   Begin VB.Label lbls 
      Height          =   390
      Index           =   9
      Left            =   240
      TabIndex        =   15
      Tag             =   "1604"
      Top             =   3390
      WhatsThisHelpID =   1603
      Width           =   945
   End
   Begin VB.Label lbls 
      Height          =   390
      Index           =   7
      Left            =   240
      TabIndex        =   2
      Tag             =   "1601"
      Top             =   660
      WhatsThisHelpID =   1601
      Width           =   1065
   End
   Begin VB.Label lbls 
      Height          =   390
      Index           =   6
      Left            =   840
      TabIndex        =   13
      Tag             =   "1047"
      Top             =   2820
      WhatsThisHelpID =   1045
      Width           =   1065
   End
   Begin VB.Label lbls 
      Height          =   390
      Index           =   5
      Left            =   840
      TabIndex        =   11
      Tag             =   "1046"
      Top             =   2430
      WhatsThisHelpID =   1045
      Width           =   1065
   End
   Begin VB.Label lbls 
      Height          =   285
      Index           =   4
      Left            =   600
      TabIndex        =   10
      Tag             =   "1045"
      Top             =   2115
      WhatsThisHelpID =   1045
      Width           =   2175
   End
   Begin VB.Label lbls 
      Height          =   390
      Index           =   3
      Left            =   2880
      TabIndex        =   8
      Tag             =   "1044"
      Top             =   1710
      WhatsThisHelpID =   1042
      Width           =   345
   End
   Begin VB.Label lbls 
      Height          =   390
      Index           =   2
      Left            =   840
      TabIndex        =   6
      Tag             =   "1043"
      Top             =   1710
      WhatsThisHelpID =   1042
      Width           =   345
   End
   Begin VB.Label lbls 
      Height          =   285
      Index           =   1
      Left            =   600
      TabIndex        =   5
      Tag             =   "1042"
      Top             =   1395
      WhatsThisHelpID =   1042
      Width           =   2175
   End
End
Attribute VB_Name = "frmBeacon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mChanged As Boolean, mTxtChanged As Boolean

Private mValueCache(7) As Single

Public Function EditDataNDB(Data As clsNDB) As Boolean
  Dim Offset As Single
  
  Load frmBeacon

  With Data
    Txts(0).Text = .Name
    Txts_Change 0 ' (Changes the caption)
    Txts(1).Text = MeterToUser(.X)
    Txts(2).Text = MeterToUser(.Y)
    chkLocked.Value = -.Locked
    Txts_Validate 1, True ' (Updates Latitude, Longitude)

    Txts(5).Text = .ID
    Txts(6).Text = Append(.Frequency, RES_Unit_AbbrevKhz)
    Txts(7).Text = NauticalToUser(.Range)
    
    chks(0).Value = -.AFDEntry
    
    chks(1).Visible = False
    chks(2).Visible = False
    chks(3).Visible = False
    
    Offset = cmdOK.Top - chks(1).Top
    cmdOK.Top = cmdOK.Top - Offset
    cmdCancel.Top = cmdCancel.Top - Offset
    Height = Height - Offset
    
    lbls(9).Tag = RES_Rdo_FrequencyNDB
    Txts(6).Tag = RES_Rdo_FrequencyNDB
    
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
      
      .ID = Txts(5).Text
      .Frequency = mValueCache(6)
      .Range = mValueCache(7)
      
      .AFDEntry = -chks(0).Value
      
      Scenery.AFDRefresh = True
      EditDataNDB = True
    End If
  End With

  Unload frmBeacon
End Function

Public Function EditDataVOR(Data As clsVOR) As Boolean
  Load frmBeacon

  With Data
    Txts(0).Text = .Name
    Txts_Change 0 ' (Changes the caption)
    Txts(1).Text = MeterToUser(.X)
    Txts(2).Text = MeterToUser(.Y)
    chkLocked.Value = -.Locked
    Txts_Validate 1, True ' (Updates Latitude, Longitude)

    Txts(5).Text = .ID
    Txts(6).Text = Append(.Frequency, RES_Unit_AbbrevMhz, "##0.00")
    Txts(7).Text = NauticalToUser(.Range)
    
    chks(0).Value = -((.Flags And 128) > 0)
    chks(1).Value = -((.Flags And 8) = 0)
    chks(2).Value = -((.Flags And 1) > 0)
    chks(3).Value = -((.Flags And 2) > 0)
    
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
      
      .ID = Txts(5).Text
      .Frequency = mValueCache(6)
      .Range = mValueCache(7)
      
      .Flags = chks(0).Value * 128 + _
               Abs(chks(1).Value - 1) * 8 + _
               chks(2).Value * 1 + _
               chks(3).Value * 2

      Scenery.AFDRefresh = True
      EditDataVOR = True
    End If
  End With

  Unload frmBeacon
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

  For I = 0 To 7
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
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then Cancel = True: Hide
End Sub

Private Sub Txts_Change(Index As Integer)
  Caption = Txts(0).Text
  mTxtChanged = True
End Sub

Private Sub Txts_GotFocus(Index As Integer)
  Select Case Index
    Case 1, 2, 6, 7
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
    Case 5
      KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
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
