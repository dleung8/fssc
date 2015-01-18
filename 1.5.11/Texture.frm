VERSION 5.00
Begin VB.Form frmTexture 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8655
   Icon            =   "Texture.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CheckBox chkExpand 
      Height          =   195
      Left            =   5160
      TabIndex        =   8
      Tag             =   "2205"
      Top             =   4560
      WhatsThisHelpID =   2205
      Width           =   2895
   End
   Begin VB.CheckBox chkSpecial 
      Height          =   195
      Index           =   4
      Left            =   5160
      TabIndex        =   14
      Tag             =   "2211"
      Top             =   5640
      WhatsThisHelpID =   2206
      Width           =   1800
   End
   Begin VB.CheckBox chkSpecial 
      Height          =   195
      Index           =   3
      Left            =   5160
      TabIndex        =   13
      Tag             =   "2210"
      Top             =   5400
      WhatsThisHelpID =   2206
      Width           =   1800
   End
   Begin VB.CheckBox chkSpecial 
      Height          =   195
      Index           =   2
      Left            =   3120
      TabIndex        =   12
      Tag             =   "2209"
      Top             =   5880
      WhatsThisHelpID =   2206
      Width           =   1800
   End
   Begin VB.CheckBox chkSpecial 
      Height          =   195
      Index           =   1
      Left            =   3120
      TabIndex        =   11
      Tag             =   "2208"
      Top             =   5640
      WhatsThisHelpID =   2206
      Width           =   1800
   End
   Begin VB.CheckBox chkSpecial 
      Height          =   195
      Index           =   0
      Left            =   3120
      TabIndex        =   10
      Tag             =   "2207"
      Top             =   5400
      WhatsThisHelpID =   2206
      Width           =   1800
   End
   Begin VB.PictureBox picBuffer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   7200
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   77
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.FileListBox lstFiles 
      Height          =   2235
      Left            =   225
      TabIndex        =   2
      Top             =   840
      WhatsThisHelpID =   2202
      Width           =   2685
   End
   Begin VB.CommandButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   7200
      TabIndex        =   15
      Tag             =   "1030"
      Top             =   240
      WhatsThisHelpID =   1030
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   7200
      TabIndex        =   16
      Tag             =   "1031"
      Top             =   720
      WhatsThisHelpID =   1031
      Width           =   1215
   End
   Begin VB.CheckBox chkFlip 
      Height          =   195
      Left            =   3120
      TabIndex        =   7
      Tag             =   "2204"
      Top             =   4560
      WhatsThisHelpID =   2204
      Width           =   1815
   End
   Begin VB.PictureBox picTexture 
      ClipControls    =   0   'False
      FillColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3900
      Left            =   3120
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   3900
   End
   Begin VB.CommandButton cmdSave 
      Enabled         =   0   'False
      Height          =   375
      Left            =   7200
      TabIndex        =   17
      Tag             =   "2230"
      Top             =   1320
      WhatsThisHelpID =   2230
      Width           =   1215
   End
   Begin VB.CommandButton cmdBrowse 
      Height          =   375
      Left            =   7200
      TabIndex        =   18
      Tag             =   "1032"
      Top             =   1800
      WhatsThisHelpID =   1032
      Width           =   1215
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   480
      WhatsThisHelpID =   2202
      Width           =   2655
   End
   Begin VB.DirListBox dirDirs 
      Height          =   2565
      Left            =   240
      TabIndex        =   3
      Top             =   3150
      WhatsThisHelpID =   2202
      Width           =   2655
   End
   Begin VB.DriveListBox drvDrives 
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Top             =   5805
      WhatsThisHelpID =   2202
      Width           =   2655
   End
   Begin VB.Label lbls 
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Tag             =   "2202"
      Top             =   240
      WhatsThisHelpID =   2202
      Width           =   2175
   End
   Begin VB.Label lbls 
      Height          =   195
      Index           =   1
      Left            =   3120
      TabIndex        =   5
      Tag             =   "2203"
      Top             =   240
      Width           =   2055
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   3120
      X2              =   6960
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   3120
      X2              =   6960
      Y1              =   4935
      Y2              =   4935
   End
   Begin VB.Label lbls 
      Height          =   195
      Index           =   2
      Left            =   3120
      TabIndex        =   9
      Tag             =   "2206"
      Top             =   5100
      WhatsThisHelpID =   2206
      Width           =   3855
   End
End
Attribute VB_Name = "frmTexture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private mChanged As Boolean, DrawErr As Boolean

Public Enum TexturesEnum
  Tex_File
  TEX_Background
End Enum

Public OldPath As String

' Get a texture file name
Public Function EditData(ByRef Data As String, ByVal TType As TexturesEnum) As Boolean
  Dim Flags As String, I As Integer
  Dim OldCurDir As String
  
  OldCurDir = CurDir$
  
  Load frmTexture

  If TType = Tex_File Then
    Tag = RES_TEX_Texture
    lstFiles.Pattern = Options.TextureFilter
  Else
    Tag = RES_TEX_Background
    lstFiles.Pattern = "*.bmp;*.ico;*.rle;*.wmf;*.emf;*.gif;*.jpg"
    cmdSave.Visible = False
    chkFlip.Visible = False
  End If
  
  If TType = TEX_Background Or Options.FSVersion < Version_FS2K Then
    Line1.Visible = False
    Line2.Visible = False
    lbls(2).Visible = False
    For I = 0 To 4
      chkSpecial(I).Visible = False
    Next I
    lstFiles.Height = 1845
    dirDirs.Top = 2760
    dirDirs.Height = 1665
    drvDrives.Top = 4560
    Height = 5460
  End If
  Lang.PrepareForm Me

  CenterForm Me

  lstFiles.ListIndex = -1
  
  If OldPath = "" Then OldPath = Options.TexturePath
  
  On Error Resume Next
  drvDrives.Drive = Left$(OldPath, 1)
  dirDirs.Path = OldPath
  On Error GoTo 0
  
  If Data = "" Or Data = "z" Then ' dummy value for indeterminate selection
    For I = 0 To 4
      chkSpecial(I).Enabled = False
    Next I
  ElseIf TType = TEX_Background Then
    txtFile.Text = Data
  Else
    Flags = ReadNext(Data, " ")
    chkSpecial(0).Value = -(InStr(Flags, "L") > 0)
    chkSpecial(1).Value = -(InStr(Flags, "S") > 0)
    chkSpecial(2).Value = -(InStr(Flags, "F") > 0)
    chkSpecial(3).Value = -(InStr(Flags, "W") > 0)
    chkSpecial(4).Value = -(InStr(Flags, "H") > 0)
    
    txtFile.Text = Data
  End If
  
  mChanged = False
  Show vbModal, Screen.ActiveForm
  If mChanged Then
    If txtFile.Text = "" Then
      Data = ""
    ElseIf TType = TEX_Background Then
      Data = AddDir(dirDirs.Path, LCase$(txtFile.Text))
    Else
      Flags = "N" & IIf(chkSpecial(0).Value = vbChecked, "L", "") & IIf(chkSpecial(1).Value = vbChecked, "S", "") & IIf(chkSpecial(2).Value = vbChecked, "F", "") & IIf(chkSpecial(3).Value = vbChecked, "W", "") & IIf(chkSpecial(4).Value = vbChecked, "H", "") & " "
      Data = Flags & AddDir(dirDirs.Path, LCase$(txtFile.Text))
    End If

    EditData = True
    OldPath = Options.TexturePath
  End If
  Unload frmTexture
  
  ChangeDir OldCurDir
End Function

' Load the file into the picture buffer
Private Sub LoadFile(ByVal File As String)
  On Error Resume Next
  Dim Base As String, Ext As String, _
    Result As String, I As Integer
  
  DrawErr = False
  
  Set picBuffer.Picture = LoadPicture()
  If File <> "" Then
    LoadTexture File, Result
    If Result = "" Then
      Set picBuffer.Picture = LoadPicture()
      DrawErr = True
    Else
      Set picBuffer.Picture = LoadPicture(Result)
      Kill Result
    End If
    DrawErr = DrawErr Or Err > 0
    
    If Tag = RES_TEX_Texture Then
      ' Detect special textures
      Ext = Mid$(File, InStrRev(File, "."))
      Base = Left$(File, Len(File) - Len(Ext))
      chkSpecial(0).Enabled = FileExists(Base & "_LM" & Ext)
      chkSpecial(1).Enabled = FileExists(Base & "_SP" & Ext)
      chkSpecial(2).Enabled = FileExists(Base & "_FA" & Ext)
      chkSpecial(3).Enabled = FileExists(Base & "_WI" & Ext)
      chkSpecial(4).Enabled = FileExists(Base & "_HW" & Ext)
    
      If Visible Then
        For I = 0 To 4
          chkSpecial(I).Value = -chkSpecial(I).Enabled
        Next I
      End If
    End If
  End If
  cmdSave.Enabled = File <> "" And Not DrawErr And cmdSave.Visible
  picTexture_Paint
End Sub

Private Sub chkExpand_Click()
  picTexture_Paint
End Sub

Private Sub chkFlip_Click()
  picTexture_Paint
End Sub

Private Sub cmdBrowse_Click()
  Dim X As String
  With cDialog
    .FilterIndex = 1
    .DefExt = ""
    If Tag = RES_TEX_Texture Then
      .Filter = Lang.ResolveString(RES_TEX_TextureFilter, Options.TextureFilter)
    Else
      .Filter = Lang.GetString(RES_TEX_PictureFilter)
    End If
    X = .OpenDialog(Lang.GetString(RES_TEX_Browse), IIf(txtFile.Text <> "", AddDir(dirDirs.Path, txtFile.Text), ""))
  End With
  If X <> "" Then txtFile.Text = X
End Sub

Private Sub cmdCancel_Click()
  Hide
End Sub

Private Sub cmdOK_Click()
  Dim File As String
  On Error Resume Next
  File = txtFile.Text
  If File <> "" And Not FileExists(AddDir(dirDirs.Path, File)) Then
    If (GetAttr(File) And vbDirectory) = 0 Then
      MsgBoxEx Me, Lang.GetString(RES_ERR_FileExists), vbExclamation, RES_ERR_FileExists
    Else
      drvDrives.Drive = UCase$(Left$(File, 1))
      dirDirs.Path = GetRealName(File)
    End If
  Else
    mChanged = True
    Hide
  End If
End Sub

Private Sub cmdSave_Click()
  Dim X As String
  ' Ask for filename and save picture
  With cDialog
    .FilterIndex = 1
    .DefExt = "bmp"
    .Filter = Lang.GetString(RES_TEX_BitmapFilter)
    X = .SaveDialog(MakeFileNameNeat(txtFile.Text), Lang.GetString(RES_TEX_SaveAsBitmap))
    If X <> "" Then LoadTexture AddDir(dirDirs.Path, txtFile.Text), X
  End With
End Sub

Private Sub dirDirs_Change()
  lstFiles.Path = dirDirs.Path
  txtFile.Text = ""
  LoadFile ""
End Sub

Private Sub drvDrives_Change()
  Dim TempPath As String
  
  On Error Resume Next
  TempPath = dirDirs.Path
  Do
    Err.Number = 0
    SetScreenMousePointer vbHourglass
    ChDrive drvDrives.Drive
    SetScreenMousePointer vbDefault
    If Err.Number > 0 Then
      If MsgBoxEx(Me, Lang.ResolveString(RES_ERR_DriveExists, UCase$(Left$(drvDrives.Drive, 1)) & "\"), vbCritical Or vbRetryCancel, RES_ERR_DriveExists) = vbCancel Then
        drvDrives.Drive = dirDirs.Path
        dirDirs.Path = TempPath
        Exit Do
      End If
    Else
      dirDirs.Path = UCase$(drvDrives.Drive)
      Exit Do
    End If
  Loop
End Sub

Private Sub Form_Load()
  DialogMenus Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then Cancel = True: Hide
End Sub

Private Sub lstFiles_Click()
  txtFile.Text = lstFiles.List(lstFiles.ListIndex)
End Sub

Private Sub lstFiles_DblClick()
  cmdOK_Click
End Sub

Private Sub picTexture_Paint()
  Dim myWidth As Integer, myHeight As Integer, _
    NewWidth As Integer, NewHeight As Integer
  If DrawErr Then
    picTexture.Cls
    CenterText picTexture, Lang.GetString(RES_ERR_TextureError), 32, (picTexture.ScaleHeight - picTexture.TextHeight("H")) / 2
    DrawExclaimIcon picTexture.hdc, 20, picTexture.ScaleHeight / 2 - 16
  ElseIf picBuffer.Picture = 0 Then
    picTexture.Cls
  Else
    myWidth = picBuffer.ScaleWidth
    myHeight = picBuffer.ScaleHeight
    If chkExpand.Value = vbChecked Then
      NewWidth = 256
      NewHeight = 256
    Else
      NewHeight = myHeight
      NewWidth = myWidth
      If NewWidth > 256 Then
        NewHeight = NewHeight * (256 / NewWidth)
        NewWidth = 256
      End If
      If NewHeight > 256 Then
        NewWidth = NewWidth * (256 / NewHeight)
        NewHeight = 256
      End If
    End If
    StretchBlt picTexture.hdc, 0, IIf(-chkFlip.Value, NewHeight - 1, 0), NewWidth, IIf(-chkFlip.Value, -NewHeight, NewHeight), picBuffer.hdc, 0, 0, myWidth, myHeight, SRCCOPY
    ' Erase old places if necessary (prevents flickering which cls does)
    picTexture.Line (NewWidth, 0)-(256, 256), vbButtonFace, BF
    picTexture.Line (0, NewHeight)-(NewWidth, 256), vbButtonFace, BF
  End If
End Sub

Private Sub txtFile_Change()
  On Error GoTo ErrorLoadFile:
  Dim File As String, Path As String
  Dim Result As Integer
  Const LB_FINDSTRINGEXACT = &H1A2
  Const LB_FINDSTRING = &H18F
  
DrawErr = False
  File = GetRealName(AddDir(dirDirs.Path, txtFile.Text))
  If FileExists(File) Then
    If (GetAttr(File) And vbDirectory) = 0 Then
      LoadFile File
      lstFiles.ListIndex = SendMessageStr(lstFiles.hwnd, LB_FINDSTRINGEXACT, -1, CStr(txtFile.Text))
    Else
      Set picBuffer.Picture = LoadPicture()
    End If
  Else
    File = GetRealName(txtFile.Text)
    If FileExists(File) Then
      If (GetAttr(File) And vbDirectory) = 0 Then
        Path = GetDir(File)
        drvDrives.Drive = Path
        dirDirs.Path = Path
        ' This will trigger txtFile_Change with case 1
        txtFile.Text = GetFileTitle(File)
      Else
        Set picBuffer.Picture = LoadPicture()
      End If
    Else
      Set picBuffer.Picture = LoadPicture()
      Result = SendMessageStr(lstFiles.hwnd, LB_FINDSTRING, -1, CStr(txtFile.Text))
      If Result > -1 Then
        lstFiles.TopIndex = Result
      End If
    End If
  End If

  picTexture_Paint
  Exit Sub
ErrorLoadFile:
  Set picBuffer.Picture = LoadPicture()
  picTexture_Paint
  Exit Sub
End Sub

Private Sub txtFile_GotFocus()
  SelectText txtFile
End Sub
