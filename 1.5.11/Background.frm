VERSION 5.00
Begin VB.Form frmBackground 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8085
   ClipControls    =   0   'False
   Icon            =   "Background.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CheckBox chkVisible 
      Height          =   195
      Left            =   4320
      TabIndex        =   37
      Tag             =   "1756"
      Top             =   4230
      WhatsThisHelpID =   1756
      Width           =   3015
   End
   Begin VB.CheckBox chkLockAspect 
      Height          =   195
      Left            =   4320
      TabIndex        =   18
      Tag             =   "1753"
      Top             =   2115
      WhatsThisHelpID =   1753
      Width           =   2655
   End
   Begin VB.TextBox Txts 
      Height          =   285
      Index           =   12
      Left            =   1800
      TabIndex        =   36
      Tag             =   "1048"
      Top             =   4200
      WhatsThisHelpID =   1048
      Width           =   1455
   End
   Begin VB.TextBox Txts 
      Height          =   285
      Index           =   8
      Left            =   5040
      TabIndex        =   27
      Tag             =   "1043"
      Top             =   2400
      Visible         =   0   'False
      WhatsThisHelpID =   1042
      Width           =   1095
   End
   Begin VB.TextBox Txts 
      Height          =   285
      Index           =   9
      Left            =   6720
      TabIndex        =   29
      Tag             =   "1044"
      Top             =   2400
      Visible         =   0   'False
      WhatsThisHelpID =   1042
      Width           =   1095
   End
   Begin VB.TextBox Txts 
      Height          =   285
      Index           =   10
      Left            =   5760
      TabIndex        =   32
      Top             =   3120
      Visible         =   0   'False
      WhatsThisHelpID =   1045
      Width           =   2055
   End
   Begin VB.TextBox Txts 
      Height          =   285
      Index           =   11
      Left            =   5760
      TabIndex        =   34
      Top             =   3510
      Visible         =   0   'False
      WhatsThisHelpID =   1045
      Width           =   2055
   End
   Begin VB.OptionButton optMethod 
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   7
      Tag             =   "1752"
      Top             =   1800
      WhatsThisHelpID =   1751
      Width           =   3015
   End
   Begin VB.OptionButton optMethod 
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   6
      Tag             =   "1751"
      Top             =   1560
      Value           =   -1  'True
      WhatsThisHelpID =   1751
      Width           =   3015
   End
   Begin VB.CommandButton cmdBrowse 
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Tag             =   "1032"
      Top             =   600
      WhatsThisHelpID =   1032
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   6480
      TabIndex        =   39
      Tag             =   "1031"
      Top             =   4560
      WhatsThisHelpID =   1031
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   5160
      TabIndex        =   38
      Tag             =   "1030"
      Top             =   4560
      WhatsThisHelpID =   1030
      Width           =   1215
   End
   Begin VB.TextBox Txts 
      Height          =   285
      Index           =   7
      Left            =   6000
      TabIndex        =   23
      Tag             =   "1755"
      Top             =   2790
      WhatsThisHelpID =   1754
      Width           =   975
   End
   Begin VB.TextBox Txts 
      Height          =   285
      Index           =   6
      Left            =   6000
      TabIndex        =   20
      Tag             =   "1754"
      Top             =   2400
      WhatsThisHelpID =   1754
      Width           =   975
   End
   Begin VB.TextBox Txts 
      Height          =   285
      Index           =   5
      Left            =   1320
      MaxLength       =   200
      TabIndex        =   3
      Tag             =   "1750"
      Top             =   630
      WhatsThisHelpID =   1750
      Width           =   3855
   End
   Begin VB.TextBox Txts 
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   17
      Top             =   3510
      WhatsThisHelpID =   1045
      Width           =   2055
   End
   Begin VB.TextBox Txts 
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   15
      Top             =   3120
      WhatsThisHelpID =   1045
      Width           =   2055
   End
   Begin VB.TextBox Txts 
      Height          =   285
      Index           =   2
      Left            =   3000
      TabIndex        =   12
      Tag             =   "1044"
      Top             =   2400
      WhatsThisHelpID =   1042
      Width           =   1095
   End
   Begin VB.TextBox Txts 
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   10
      Tag             =   "1043"
      Top             =   2400
      WhatsThisHelpID =   1042
      Width           =   1095
   End
   Begin VB.CheckBox chkLocked 
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Tag             =   "1041"
      Top             =   1200
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
      Width           =   3855
   End
   Begin VB.Label lbls 
      Height          =   390
      Index           =   11
      Left            =   7080
      TabIndex        =   24
      Tag             =   "1090"
      Top             =   2820
      WhatsThisHelpID =   1754
      Width           =   705
   End
   Begin VB.Label lbls 
      Height          =   390
      Index           =   9
      Left            =   7080
      TabIndex        =   21
      Tag             =   "1090"
      Top             =   2430
      WhatsThisHelpID =   1754
      Width           =   705
   End
   Begin VB.Label lbls 
      Height          =   390
      Index           =   18
      Left            =   360
      TabIndex        =   35
      Tag             =   "1048"
      Top             =   4230
      WhatsThisHelpID =   1048
      Width           =   1305
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   240
      X2              =   7800
      Y1              =   4020
      Y2              =   4020
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   240
      X2              =   7800
      Y1              =   4035
      Y2              =   4035
   End
   Begin VB.Label lbls 
      Height          =   285
      Index           =   12
      Left            =   4320
      TabIndex        =   25
      Tag             =   "1042"
      Top             =   2115
      Visible         =   0   'False
      WhatsThisHelpID =   1042
      Width           =   2175
   End
   Begin VB.Label lbls 
      Height          =   390
      Index           =   13
      Left            =   4560
      TabIndex        =   26
      Tag             =   "1043"
      Top             =   2430
      Visible         =   0   'False
      WhatsThisHelpID =   1042
      Width           =   345
   End
   Begin VB.Label lbls 
      Height          =   390
      Index           =   14
      Left            =   6240
      TabIndex        =   28
      Tag             =   "1044"
      Top             =   2430
      Visible         =   0   'False
      WhatsThisHelpID =   1042
      Width           =   345
   End
   Begin VB.Label lbls 
      Height          =   285
      Index           =   15
      Left            =   4320
      TabIndex        =   30
      Tag             =   "1045"
      Top             =   2835
      Visible         =   0   'False
      WhatsThisHelpID =   1045
      Width           =   2175
   End
   Begin VB.Label lbls 
      Height          =   390
      Index           =   16
      Left            =   4560
      TabIndex        =   31
      Tag             =   "1046"
      Top             =   3150
      Visible         =   0   'False
      WhatsThisHelpID =   1045
      Width           =   1065
   End
   Begin VB.Label lbls 
      Height          =   390
      Index           =   17
      Left            =   4560
      TabIndex        =   33
      Tag             =   "1047"
      Top             =   3540
      Visible         =   0   'False
      WhatsThisHelpID =   1045
      Width           =   1065
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
      Index           =   10
      Left            =   4560
      TabIndex        =   22
      Tag             =   "1755"
      Top             =   2820
      WhatsThisHelpID =   1754
      Width           =   1305
   End
   Begin VB.Label lbls 
      Height          =   390
      Index           =   8
      Left            =   4560
      TabIndex        =   19
      Tag             =   "1754"
      Top             =   2430
      WhatsThisHelpID =   1754
      Width           =   1305
   End
   Begin VB.Label lbls 
      Height          =   390
      Index           =   7
      Left            =   240
      TabIndex        =   2
      Tag             =   "1750"
      Top             =   660
      WhatsThisHelpID =   1750
      Width           =   945
   End
   Begin VB.Label lbls 
      Height          =   390
      Index           =   6
      Left            =   840
      TabIndex        =   16
      Tag             =   "1047"
      Top             =   3540
      WhatsThisHelpID =   1045
      Width           =   1065
   End
   Begin VB.Label lbls 
      Height          =   390
      Index           =   5
      Left            =   840
      TabIndex        =   14
      Tag             =   "1046"
      Top             =   3150
      WhatsThisHelpID =   1045
      Width           =   1065
   End
   Begin VB.Label lbls 
      Height          =   285
      Index           =   4
      Left            =   600
      TabIndex        =   13
      Tag             =   "1045"
      Top             =   2835
      WhatsThisHelpID =   1045
      Width           =   2175
   End
   Begin VB.Label lbls 
      Height          =   390
      Index           =   3
      Left            =   2520
      TabIndex        =   11
      Tag             =   "1044"
      Top             =   2430
      WhatsThisHelpID =   1042
      Width           =   345
   End
   Begin VB.Label lbls 
      Height          =   390
      Index           =   2
      Left            =   840
      TabIndex        =   9
      Tag             =   "1043"
      Top             =   2430
      WhatsThisHelpID =   1042
      Width           =   345
   End
   Begin VB.Label lbls 
      Height          =   285
      Index           =   1
      Left            =   600
      TabIndex        =   8
      Tag             =   "1042"
      Top             =   2115
      WhatsThisHelpID =   1042
      Width           =   2175
   End
End
Attribute VB_Name = "frmBackground"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mChanged As Boolean, mTxtChanged As Boolean, _
        mFocus As Integer

Private mValueCache(12) As Single

Public Function EditData(Data As clsBackground) As Boolean
  Load frmBackground

  With Data
    Txts(0).Tag = .Caption(True)
    Txts(0).Text = .Name
    Txts_Change 0 ' (Changes the caption)
    Txts(1).Text = MeterToUser(.X)
    Txts(2).Text = MeterToUser(.Y)
    chkLocked.Value = -.Locked
    Txts_Validate 1, True ' (Updates Latitude, Longitude)
    Txts(5).Text = .File
    Txts(6).Text = MeterToUser(.ZoomX)
    Txts(7).Text = MeterToUser(.ZoomY)
    Txts(12).Text = GeographicToUser(.Rotation)
    chkLockAspect.Value = -.LockAspectRatio
    chkVisible.Value = -.Visible
    chkVisible.Enabled = Options.FSVersion >= Version_FS2K

    Lang.PrepareForm Me
    mTxtChanged = False
    mChanged = False
    Show vbModal, Screen.ActiveForm

    If mChanged Then
      .Name = Txts(0).Text
      .X = mValueCache(1)
      .Y = mValueCache(2)
      .Locked = -chkLocked.Value
      .File = Txts(5).Text
      .ZoomX = mValueCache(6)
      .ZoomY = mValueCache(7)
      .Rotation = mValueCache(12)
      .LockAspectRatio = (chkLockAspect.Value = vbChecked)
      .Visible = (chkVisible.Value = vbChecked)
      EditData = True
    End If
  End With

  Unload frmBackground
End Function

Private Sub GetPictureDimensions(File As String, ByRef picWidth As Long, ByRef picHeight As Long)
  Dim X As String, FileNum As Integer, TempName As String
  Dim Pic As PicType, Result As String

  If Not FileExists(File) Then Exit Sub
  
  LoadTexture File, Result, picWidth, picHeight
  If FileExists(Result) Then
    Kill Result
    Exit Sub
  End If
  
  FileNum = FreeFile
  If FileLen(File) = 65536 Then
    picWidth = 256
    picHeight = 256
    Exit Sub
  ElseIf StrComp(Right$(File, 3), "bmp", vbTextCompare) <> 0 Then
    TempName = AddDir(TempPathName, "fssc.bmp")
    frmMain.picTemp = LoadPicture(File)
    SavePicture frmMain.picTemp.Image, TempName
    Open TempName For Binary As #FileNum
  Else
    Open File For Binary As #FileNum
  End If
  Get #FileNum, 19, picWidth
  Get #FileNum, 23, picHeight
  Close #FileNum
  
  If TempName <> "" Then Kill TempName
End Sub

Private Sub chkLocked_Click()
  Dim Value As Boolean, I As Integer
  Value = Not -chkLocked.Value
  For I = 1 To 11
    If I <> 5 Then SetEnabled Txts(I), Value
  Next I
  optMethod(0).Enabled = Value
  optMethod(1).Enabled = Value
  chkLockAspect.Enabled = Value
End Sub

Private Sub chkLockAspect_Click()
  Dim Value As Boolean
  Value = (chkLockAspect.Value = vbChecked)
  SetEnabled Txts(7), Not Value
  If Value Then
    Txts(7).Text = Txts(6).Text
  End If
End Sub

Private Sub cmdBrowse_Click()
  Dim DataStr As String
  DataStr = Txts(5).Text
  If frmTexture.EditData(DataStr, TEX_Background) Then
    Txts(5).Text = DataStr
    Txts(5).SelStart = Len(DataStr)
  End If
End Sub

Private Sub cmdCancel_Click()
  Hide
End Sub

Private Sub cmdOK_Click()
  Dim I As Integer, Msg As String, Cancel As Boolean
  Dim picWidth As Long, picHeight As Long, Width2Up As Integer, Height2Up As Integer
  ' Validate event not fired when Enter key pressed
  ' bug workaround
  If TypeOf ActiveControl Is TextBox Then
    Txts_Validate ActiveControl.Index, Cancel
    If Cancel Then Exit Sub
  End If
  
  If optMethod(1).Value Then
    For I = 8 To 9
      If Not Validate(Txts(I), Msg, mValueCache(I)) Then
        GoTo ValidationError:
      End If
    Next I
    optMethod(0).Value = True
  End If

  For I = 1 To 12
    If Txts(I).Visible Then
      If Not Validate(Txts(I), Msg, mValueCache(I)) Then
        GoTo ValidationError:
      End If
    End If
  Next I

  GetPictureDimensions Txts(5).Text, picWidth, picHeight
  If picWidth = 0 Or picHeight = 0 Then
    ' error
    Msg = Lang.GetString(RES_ERR_FileExists)
    I = 5
    GoTo ValidationError:
  End If
  mChanged = True
  Hide
  Exit Sub
ValidationError:
  MsgBoxEx Me, Msg, vbCritical, RES_ERR_Bound
  Txts(I).SetFocus
  Exit Sub
End Sub

Private Sub Form_Load()
  CenterForm Me
  DialogMenus Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then Cancel = True: Hide
End Sub

Private Sub optMethod_Click(Index As Integer)
  Dim Value As Boolean, I As Integer
  Value = Txts(6).Visible
  
  Txts(6).Visible = Not Value
  Txts(7).Visible = Not Value
  chkLockAspect.Visible = Not Value
  For I = 8 To 11
    lbls(I).Visible = Not Value
  Next I
  
  For I = 8 To 11
    Txts(I).Visible = Value
  Next I
  For I = 12 To 17
    lbls(I).Visible = Value
  Next I
  
  Dim Pt As PointType, Pt1 As PointType, Pt2 As PointType, _
    Delta As PointType, Zoom As PointType, _
    picWidth As Long, picHeight As Long, _
    Angle As Single

  GetPictureDimensions Txts(5).Text, picWidth, picHeight
  If picWidth = 0 Or picHeight = 0 Then
    ' error
  End If
  
  If Not Validate(Txts(12), "", Angle) Then
    ' error
  End If
  
  If Not (Validate(Txts(6), "", Zoom.X) And Validate(Txts(7), "", Zoom.Y)) Then
    ' error
  End If
  
  On Error Resume Next
  If Value Then
    ' From Center
    If Validate(Txts(1), "", Pt.X) And Validate(Txts(2), "", Pt.Y) Then
      Delta = MakePoint(picWidth * Zoom.X / 2, picHeight * Zoom.Y / 2)
      Rotate Delta, -Angle
      
      Txts(1).Text = MeterToUser(Pt.X - Delta.X, "0.0")
      Txts(2).Text = MeterToUser(Pt.Y + Delta.Y, "0.0")
      Txts_Validate 1, False
      Txts(8).Text = MeterToUser(Pt.X + Delta.X, "0.0")
      Txts(9).Text = MeterToUser(Pt.Y - Delta.Y, "0.0")
      Txts_Validate 8, False
    Else
      'error
    End If
  Else
    ' From Corners
    If Validate(Txts(1), "", Pt1.X) And Validate(Txts(2), "", Pt1.Y) And Validate(Txts(8), "", Pt2.X) And Validate(Txts(9), "", Pt2.Y) Then
      Pt = MakePoint((Pt1.X + Pt2.X) / 2, (Pt1.Y + Pt2.Y) / 2)
      Txts(1).Text = MeterToUser(Round(Pt.X, 2), "0.00")
      Txts(2).Text = MeterToUser(Round(Pt.Y, 2), "0.00")
      Pt1 = MakePoint(Abs(Pt1.X - Pt.X), Abs(Pt1.Y - Pt.Y))
      Pt2 = MakePoint(Abs(Pt2.X - Pt.X), Abs(Pt2.Y - Pt.Y))
      
      Txts_Validate 1, False
      Rotate Pt1, Angle
      Rotate Pt2, Angle
      Zoom.X = Round((Pt1.X + Pt2.X) / picWidth, 4)
      Zoom.Y = Round((Pt1.X + Pt2.Y) / picHeight, 4)
      If Zoom.X - Zoom.Y > 0.0001 And chkLockAspect.Value = vbChecked Then
        chkLockAspect.Value = vbUnchecked
      End If
      Txts(6).Text = MeterToUser(Zoom.X)
      Txts(7).Text = MeterToUser(Zoom.Y)
    Else
      'error
    End If
  End If
End Sub

Private Sub Txts_Change(Index As Integer)
  Dim TempStr As String
  mTxtChanged = True
  Select Case Index
    Case 0
      TempStr = Txts(0).Text
      If TempStr = "" Then
        Caption = Txts(0).Tag
      Else
        Caption = TempStr
      End If
    Case 6
      If chkLockAspect.Value = vbChecked Then
        Txts(7).Text = Txts(6).Text
      End If
  End Select
End Sub

Private Sub Txts_GotFocus(Index As Integer)
  Select Case Index
    Case 1, 2, 6, 7, 8, 9, 12
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
    TempLatLon As clsLatLon, _
    BottomIndex As Integer

  On Error Resume Next
  
  If mTxtChanged = False Then Exit Sub
  mTxtChanged = False
  
  Select Case Index
    Case 1, 2, 8, 9
      If Index = 2 Or Index = 9 Then BottomIndex = Index - 1 Else BottomIndex = Index
    
      If Validate(Txts(BottomIndex), "", valX) And Validate(Txts(BottomIndex + 1), "", valY) And Txts(BottomIndex).Text <> "" And Txts(BottomIndex + 1).Text <> "" Then
        Set TempLatLon = ReturnPoint(valX, valY)
        Txts(BottomIndex + 2).Text = TempLatLon.LatitudeUser
        Txts(BottomIndex + 3).Text = TempLatLon.LongitudeUser
      Else
        Txts(BottomIndex + 2).Text = ""
        Txts(BottomIndex + 3).Text = ""
      End If
    Case 3, 4, 10, 11
      If Index = 4 Or Index = 11 Then BottomIndex = Index - 1 Else BottomIndex = Index
      Set TempLatLon = New clsLatLon
      TempLatLon.Latitude = Txts(BottomIndex).Text
      TempLatLon.Longitude = Txts(BottomIndex + 1).Text
      If TempLatLon.Validate("") Then
        Scenery.Header.Center.CalcDistance TempLatLon, Distance, Angle
        PolarToRect Distance * NmToM, 90 - Angle, valX, valY
        Txts(BottomIndex - 2).Text = MeterToUser(Round(valX, 2), "0.00")
        Txts(BottomIndex - 1).Text = MeterToUser(Round(valY, 2), "0.00")
      Else
        Txts(BottomIndex - 2).Text = ""
        Txts(BottomIndex - 1).Text = ""
      End If
      Set TempLatLon = Nothing
  End Select
End Sub
