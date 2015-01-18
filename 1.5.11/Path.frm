VERSION 5.00
Begin VB.Form frmPath 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4365
   ClientLeft      =   150
   ClientTop       =   1530
   ClientWidth     =   4575
   ClipControls    =   0   'False
   Icon            =   "Path.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   Tag             =   "2120"
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Tag             =   "1031"
      Top             =   2040
      WhatsThisHelpID =   1031
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Tag             =   "1030"
      Top             =   1560
      WhatsThisHelpID =   1030
      Width           =   1215
   End
   Begin VB.DriveListBox drvDrives 
      Height          =   315
      Left            =   240
      TabIndex        =   6
      Top             =   3840
      WhatsThisHelpID =   2122
      Width           =   2670
   End
   Begin VB.DirListBox dirDirs 
      Height          =   1890
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      WhatsThisHelpID =   2122
      Width           =   2670
   End
   Begin VB.TextBox txtPath 
      Height          =   288
      Left            =   240
      MaxLength       =   260
      TabIndex        =   2
      Top             =   840
      WhatsThisHelpID =   2121
      Width           =   4095
   End
   Begin VB.Label lblDrives 
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Tag             =   "2123"
      Top             =   3600
      WhatsThisHelpID =   2122
      Width           =   2535
   End
   Begin VB.Label lblDirs 
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Tag             =   "2122"
      Top             =   1320
      WhatsThisHelpID =   2122
      Width           =   2595
   End
   Begin VB.Label lblPath 
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Tag             =   "2121"
      Top             =   600
      WhatsThisHelpID =   2121
      Width           =   2595
   End
   Begin VB.Label lblPrompt 
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4005
   End
End
Attribute VB_Name = "frmPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' [frmPath]
' Lets the user select/specify a folder

Option Explicit

Private mChanged As Boolean, _
        mDir As String, _
        mMustExist As Boolean

' Prompt the user for a folder
' Data - the original folder. This is set to the
'           result when the function ends
' PathMustExist - If true, checks that the selection
'                    is valid
' Returns true if the user pressed OK
Public Function EditData(ByRef Data As String, ByVal Prompt As String, Optional ByVal PathMustExist As Boolean = True) As Boolean
  Dim Y As Long, TempPath As String
  
  On Error Resume Next
  
  mMustExist = PathMustExist
  
  Load frmPath
  mDir = Data
  
  If mDir = "" Then
    TempPath = dirDirs.Path
  ElseIf Dir$(mDir, vbDirectory) <> "" Then
    dirDirs.Path = mDir
    TempPath = mDir
  Else
    TempPath = mDir
    Do
      Y = InStrRev(mDir, "\", Y - 1)
      mDir = Left$(mDir, Y - 1)
      If Dir$(mDir, vbDirectory) <> "" Then
        dirDirs.Path = mDir
        Exit Do
      End If
    Loop Until Y = 0
  End If

  drvDrives.Drive = Left$(TempPath, 2)
  txtPath.Text = TempPath

  lblPrompt.Caption = Prompt
  
  Lang.PrepareForm Me
  mChanged = False
  Show vbModal, Screen.ActiveForm
  
  If mChanged Then
    Data = GetRealName(mDir)
    EditData = True
  End If

  Unload frmPath
End Function

Private Sub cmdCancel_Click()
  Hide
End Sub

Private Sub cmdOK_Click()
  mDir = txtPath.Text

  If mDir <> "" Then
    If Dir$(mDir, vbDirectory) = "" Then
      ' selection does not exist
      If Not mMustExist Then
        If MsgBoxEx(Me, Lang.ResolveString(RES_ERR_DirCreate, mDir), vbQuestion Or vbYesNo, RES_ERR_DirCreate) = vbYes Then
          MakePath mDir
        Else
          txtPath.SetFocus
          Exit Sub
        End If
      Else
        MsgBoxEx Me, Lang.ResolveString(RES_ERR_DirExists, mDir), vbCritical, RES_ERR_DirExists
        txtPath.SetFocus
        Exit Sub
      End If
    End If
  End If
  mChanged = True
  Hide
End Sub

Private Sub dirDirs_Change()
  txtPath.Text = dirDirs.Path
End Sub

Private Sub dirDirs_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeySpace Then
    ' Enter
    dirDirs.Path = dirDirs.List(dirDirs.ListIndex)
  ElseIf KeyAscii = vbKeyBack Then
    ' Ancestor
    dirDirs.Path = dirDirs.List(-2)
  End If
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
  CenterForm Me
  DialogMenus Me

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    Cancel = True
    Hide
  End If
End Sub

Private Sub txtPath_GotFocus()
  SelectText txtPath
End Sub
