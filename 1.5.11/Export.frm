VERSION 5.00
Begin VB.Form frmExport 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6510
   Icon            =   "Export.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   Tag             =   "2240"
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   5040
      TabIndex        =   12
      Tag             =   "1031"
      Top             =   5160
      WhatsThisHelpID =   1031
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   11
      Tag             =   "1030"
      Top             =   5160
      WhatsThisHelpID =   1030
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Tag             =   "2259"
      Top             =   1560
      WhatsThisHelpID =   2259
      Width           =   1215
   End
   Begin VB.CommandButton cmdBrowse 
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Tag             =   "1032"
      Top             =   300
      WhatsThisHelpID =   1032
      Width           =   1215
   End
   Begin VB.ComboBox Cmbs 
      Height          =   315
      Index           =   1
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Tag             =   "0"
      Top             =   4680
      WhatsThisHelpID =   2246
      Width           =   2775
   End
   Begin VB.ListBox lstFiles 
      Height          =   2985
      Left            =   240
      Style           =   1  'Checkbox
      TabIndex        =   7
      Top             =   1560
      WhatsThisHelpID =   2245
      Width           =   4575
   End
   Begin VB.ComboBox Cmbs 
      Height          =   315
      Index           =   0
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   750
      WhatsThisHelpID =   2244
      Width           =   3135
   End
   Begin VB.TextBox Txts 
      Height          =   285
      Index           =   0
      Left            =   1680
      TabIndex        =   2
      Top             =   360
      WhatsThisHelpID =   2243
      Width           =   3135
   End
   Begin VB.OptionButton optMethod 
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Tag             =   "2241"
      Top             =   240
      Visible         =   0   'False
      WhatsThisHelpID =   2241
      Width           =   3615
   End
   Begin VB.Label lbls 
      Height          =   390
      Index           =   3
      Left            =   240
      TabIndex        =   9
      Tag             =   "2246"
      Top             =   4725
      WhatsThisHelpID =   2246
      Width           =   1695
   End
   Begin VB.Label lbls 
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   6
      Tag             =   "2245"
      Top             =   1320
      WhatsThisHelpID =   2245
      Width           =   2775
   End
   Begin VB.Label lbls 
      Height          =   390
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Tag             =   "2244"
      Top             =   795
      WhatsThisHelpID =   2244
      Width           =   1335
   End
   Begin VB.Label lbls 
      Height          =   390
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Tag             =   "2243"
      Top             =   390
      WhatsThisHelpID =   2243
      Width           =   1335
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' [frmExport]
' Export Wizard
Option Explicit

Private mFiles() As String
Private mSelected() As Boolean
Private mFileCount As Integer

Private mAFDIndex As Integer

Private mChanged As Boolean
Private mDontUpdate As Boolean

Private Type SHFILEOPSTRUCT
   hwnd As Long
   wFunc As Long
   pFrom As String
   pTo As String
   fFlags As Long
   fAnyOperationsAborted As Long
   hNameMappings As Long
   lpszProgressTitle As String
End Type

Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Private Const FO_COPY As Long = &H2
Private Const FOF_MULTIDESTFILES As Long = &H1
Private Const FOF_NOCONFIRMATION As Long = &H10
Private Const FOF_NOCONFIRMMKDIR As Long = &H200

Public Sub CopyFiles()
  Dim I As Integer
  Dim Result As Long
  Dim FileOp As SHFILEOPSTRUCT
  Dim PathIndex As Integer
  Dim DestPath As String

  If MakePath(Scenery.ExportPath) Then
    If MakePath(AddDir(Scenery.ExportPath, "Scenery")) Then
      Scenery.Compile ((mAFDIndex > 0) And Cmbs(0).ListIndex <> 1) And (Cmbs(0).ListIndex <> 0 Or Scenery.AFDRefresh)
      
      ' Progress bar
      With frmMain
        If .Statusbar.Visible Then
          .lblStatus.Caption = Lang.GetString(RES_Main_CopyingFiles)
          .barProgress.Max = 1
          .barProgress.Value = 1
          .Statusbar.Visible = False
          .picProgress.Visible = True
          .picProgress.Refresh
        End If
      End With
      
      With FileOp
        .hwnd = frmMain.hwnd
        .wFunc = FO_COPY
        .fFlags = FO_COPY Or FOF_MULTIDESTFILES Or FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR
        For I = 0 To mFileCount - 1
          If mSelected(I) Then
            PathIndex = CInt(Left$(mFiles(I), 1))
            Select Case PathIndex
              Case 0
                DestPath = AddDir(Scenery.ExportPath, "")
              Case 1
                DestPath = AddDir(Scenery.ExportPath, "Scenery\")
              Case 2
                DestPath = AddDir(Scenery.ExportPath, "Texture\")
              Case 3
                DestPath = AddDir(Options.FSPath, "Scenery\")
              Case 4
                DestPath = AddDir(Options.FSPath, "Texture\")
              Case Else
                DestPath = ""
            End Select
            If DestPath <> "" Then
              If FileExists(Mid$(mFiles(I), 2)) Then
                If StrComp(Mid$(mFiles(I), 2), DestPath & GetFileTitle(Mid$(mFiles(I), 2)), vbTextCompare) <> 0 Then
                  .pFrom = .pFrom & Mid$(mFiles(I), 2) & vbNullChar
                  .pTo = .pTo & DestPath & GetFileTitle(Mid$(mFiles(I), 2)) & vbNullChar
                End If
              End If
            End If
          End If
        Next I
        .pFrom = .pFrom & vbNullChar
        .pTo = .pTo & vbNullChar
        If Len(.pFrom) > 1 Then
          Result = SHFileOperation(FileOp)
        
          If Result <> 0 Or FileOp.fAnyOperationsAborted <> 0 Then
            MsgBoxEx frmMain, Lang.GetString(RES_ERR_CopyFail), vbCritical, RES_ERR_CopyFail
          End If
        End If
      End With
  
      ' Progress bar
      With frmMain
        If .picProgress.Visible Then
          .picProgress.Visible = False
          .Statusbar.Visible = True
        End If
      End With
    End If
  End If
End Sub

Public Sub DoCompile()
  Load frmExport
  CopyFiles
  Unload frmExport
End Sub

Public Sub DoWizard()
  Dim I As Integer
  Load frmExport
  mChanged = False
  Show vbModal, Screen.ActiveForm
  If mChanged Then
    For I = 0 To mFileCount - 1
      mSelected(I) = lstFiles.Selected(I)
    Next I
    CopyFiles
  End If
  Unload frmExport
End Sub

Private Sub Cmbs_Click(Index As Integer)
  If Index = 0 And Not mDontUpdate Then
    If mAFDIndex > 0 Then
      Select Case Cmbs(0).ListIndex
        Case 0
          lstFiles.Selected(mAFDIndex) = Scenery.AFDRefresh
        Case 1
          lstFiles.Selected(mAFDIndex) = False
        Case 2
          lstFiles.Selected(mAFDIndex) = True
      End Select
      lstFiles.ListIndex = 0
    End If
  ElseIf Index = 1 Then
    lstFiles.ItemData(lstFiles.ListIndex) = Cmbs(1).ListIndex
  End If
End Sub

Private Sub cmdAdd_Click()
  Dim Files() As String, I As Integer
  With cDialog
    .Filter = Lang.GetString(RES_Dist_AddFileFilter)
    .FilterIndex = 1
    .DefExt = ""
    If .SelectMultiDialog(Files(), Lang.GetString(RES_Dist_AddFileCaption)) Then
      With lstFiles
        For I = 0 To UBound(Files)
          If mFileCount > UBound(Files) Then
            ReDim Preserve mFiles(mFileCount * 2)
            ReDim Preserve mSelected(mFileCount * 2)
          End If
          mFiles(mFileCount) = "0" & Files(I)
          mFileCount = mFileCount + 1
          .AddItem GetFileTitle(Files(I))
          .ItemData(.ListCount - 1) = 0
          .Selected(.ListCount - 1) = True
          mSelected(.ListCount - 1) = True
        Next I
      End With
    End If
  End With
End Sub

Private Sub cmdBrowse_Click()
  Dim myPath As String
  
  myPath = Txts(0).Text
  
  If frmPath.EditData(myPath, Lang.GetString(RES_Dist_ChangeFolder), False) Then
    Txts(0).Text = myPath
  End If
End Sub

Private Sub cmdCancel_Click()
  Hide
End Sub

Private Sub cmdOK_Click()
  Dim I As Integer
  Scenery.ExportPath = Txts(0).Text
  For I = 0 To mFileCount - 1
    Mid$(mFiles(I), 1, 1) = CStr(lstFiles.ItemData(I))
  Next I
  mChanged = True
  Hide
  Exit Sub
End Sub

Private Sub Form_Load()
  Dim I As Integer
  Dim PointObject As clsPoint, _
      MacroObject As clsMacro, _
      BackgroundObject As clsBackground, _
      RunwayObject As clsRunway, _
      NDBObject As clsNDB, _
      VORObject As clsVOR
  Dim FileTitle As String, Filename As String, _
    FileBase As String

  Dim PathIndex As Long
  Dim DoArea16N As Boolean
  Dim ExcludeCount As Integer
  
  mFileCount = 0
  mAFDIndex = 0
  ReDim mFiles(50)
  
  CenterForm Me
  DialogMenus Me
  Lang.PrepareForm Me
  Lang.AddItems Cmbs(0), RES_CMB_Method, 3
  Lang.AddItems Cmbs(1), RES_CMB_Destination, 5
  
  If Scenery.ExportPath = "" Then
    Txts(0).Text = AddDir(Options.FSPath, "Scenery\" & MakeFileNameNeat(Scenery.File))
  Else
    Txts(0).Text = Scenery.ExportPath
  End If
  
  If Scenery.Header.Exclusion > 0 Then
    ExcludeCount = ExcludeCount + 1
  End If
  
  For I = 1 To Scenery.Count
    Select Case Scenery(I).ObjectType
      Case OT_Point
        If Scenery(I).ObjectIndex = 1 Then
          Set PointObject = Scenery(I)
          If PointObject.Parent.ShapeType = OT_FlatArea Then
            DoArea16N = DoArea16N Or PointObject.Parent.Extra3 = 1
          Else
            PointObject.Parent.ScanTextures mFiles(), mFileCount
          End If
          Set PointObject = Nothing
        End If
      Case OT_Macro
        Set MacroObject = Scenery(I)
        MacroObject.ScanTextures mFiles(), mFileCount
        Set MacroObject = Nothing
      Case OT_Background
        Set BackgroundObject = Scenery(I)
        BackgroundObject.ScanTextures mFiles(), mFileCount
        Set BackgroundObject = Nothing
      Case OT_Runway
        If mAFDIndex = 0 Then
          Set RunwayObject = Scenery(I)
          mAFDIndex = -RunwayObject.AFDEntry
          Set RunwayObject = Nothing
        End If
      Case OT_NDB
        If mAFDIndex = 0 Then
          Set NDBObject = Scenery(I)
          mAFDIndex = -NDBObject.AFDEntry
          Set NDBObject = Nothing
        End If
      Case OT_VOR
        If mAFDIndex = 0 Then
          Set VORObject = Scenery(I)
          mAFDIndex = -((VORObject.Flags And 128) > 0)
          Set VORObject = Nothing
        End If
      Case OT_Exclusion
        ExcludeCount = ExcludeCount + 1
    End Select
  Next I
  
  FileTitle = GetFileTitle(Scenery.File)
  FileBase = Left$(FileTitle, InStrRev(FileTitle, ".") - 1)
  
  If mFileCount + 3 + ExcludeCount > UBound(mFiles) Then
    ReDim Preserve mFiles(mFileCount * 2 + ExcludeCount + 3)
  End If

  mFiles(mFileCount) = "9" & FileBase & ".bgl"
  mFileCount = mFileCount + 1
  If mAFDIndex > 0 Then
    mFiles(mFileCount) = "8" & FileBase & "_AFD.bgl"
    If Not FileExists(AddDir(Txts(0).Text, "Scenery\" & FileBase & "_AFD.bgl")) Then Scenery.AFDRefresh = True
    mAFDIndex = mFileCount
    mFileCount = mFileCount + 1
  End If
  If DoArea16N Then
    mFiles(mFileCount) = "9" & FileBase & "_A16N.bgl"
    mFileCount = mFileCount + 1
  End If
  For I = 1 To ExcludeCount
    mFiles(mFileCount) = "9" & FileBase & "_exc" & I & ".bgl"
    mFileCount = mFileCount + 1
  Next I
   
  ReDim mSelected(UBound(mFiles))
  
  With lstFiles
    For I = 0 To mFileCount - 1
      PathIndex = CLng(Left$(mFiles(I), 1))
      Filename = Mid$(mFiles(I), 2)
      FileTitle = GetFileTitle(Filename)
      .AddItem FileTitle
      .ItemData(I) = PathIndex
      If PathIndex = 9 Or PathIndex = 8 Then
        mSelected(I) = True
        .Selected(I) = True
      ElseIf PathIndex = 1 Or PathIndex = 2 Then
        If FileExists(Filename) Then
          mSelected(I) = True
          .Selected(I) = True
        End If
      End If
    Next I
    
    If mAFDIndex > 0 Then
      .Selected(mAFDIndex) = Scenery.AFDRefresh
    Else
      Cmbs(0).Locked = True
      Cmbs(0).BackColor = vbButtonFace
    End If

    .ListIndex = 0
    lstFiles_Click
  End With

  Cmbs(0).ListIndex = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then Cancel = True: Hide
End Sub

Private Sub lstFiles_Click()
  Dim Value As Integer
  Value = lstFiles.ItemData(lstFiles.ListIndex)
  If Value = 9 Then
    Value = 1
  ElseIf Value = 8 Then
    mDontUpdate = True
    If lstFiles.Selected(lstFiles.ListIndex) Then
      Cmbs(0).ListIndex = 2
    Else
      Cmbs(0).ListIndex = 1
    End If
    mDontUpdate = False
    Value = 1
  End If
  Cmbs(1).ListIndex = Value
End Sub

Private Sub Txts_Change(Index As Integer)
  If Visible Then
    Scenery.AFDRefresh = True
    Cmbs_Click 0
  End If
End Sub

Private Sub Txts_GotFocus(Index As Integer)
  SelectText Txts(Index)
End Sub
