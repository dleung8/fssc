VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4950
   ClientLeft      =   1140
   ClientTop       =   1545
   ClientWidth     =   5940
   ClipControls    =   0   'False
   Icon            =   "About.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   Tag             =   "2100"
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdSysInfo 
      Height          =   360
      Left            =   4200
      TabIndex        =   14
      Tag             =   "2110"
      Top             =   4320
      WhatsThisHelpID =   2110
      Width           =   1365
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   360
      Left            =   2640
      TabIndex        =   13
      Tag             =   "1033"
      Top             =   4320
      WhatsThisHelpID =   1033
      Width           =   1380
   End
   Begin VB.Image Image1 
      Height          =   1245
      Left            =   360
      Top             =   360
      Width           =   5175
   End
   Begin VB.Label lblInternet 
      BackStyle       =   0  'Transparent
      Height          =   225
      Index           =   10
      Left            =   360
      TabIndex        =   12
      Tag             =   "2106"
      Top             =   3840
      Width           =   4095
   End
   Begin VB.Label lblInternet 
      BackStyle       =   0  'Transparent
      Caption         =   "106200333"
      Height          =   225
      Index           =   9
      Left            =   2280
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   3600
      Width           =   810
   End
   Begin VB.Label lblInternet 
      BackStyle       =   0  'Transparent
      Caption         =   "ICQ:"
      Height          =   225
      Index           =   8
      Left            =   360
      TabIndex        =   10
      Top             =   3600
      Width           =   1755
   End
   Begin VB.Label lblInternet 
      BackStyle       =   0  'Transparent
      Caption         =   "dleung8@hotmail.com"
      Height          =   225
      Index           =   7
      Left            =   2280
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label lblInternet 
      BackStyle       =   0  'Transparent
      Caption         =   "MSN Messenger:"
      Height          =   225
      Index           =   6
      Left            =   360
      TabIndex        =   8
      Top             =   3360
      Width           =   1755
   End
   Begin VB.Label lblInternet 
      BackStyle       =   0  'Transparent
      Caption         =   "FSSC Guru"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   5
      Left            =   2280
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   3120
      Width           =   795
   End
   Begin VB.Label lblInternet 
      BackStyle       =   0  'Transparent
      Caption         =   "AOL Instant Messenger:"
      Height          =   225
      Index           =   4
      Left            =   360
      TabIndex        =   6
      Top             =   3120
      Width           =   1755
   End
   Begin VB.Label lblInternet 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   3
      Left            =   2280
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   2760
      Width           =   3045
   End
   Begin VB.Label lblInternet 
      BackStyle       =   0  'Transparent
      Height          =   225
      Index           =   2
      Left            =   360
      TabIndex        =   4
      Tag             =   "2105"
      Top             =   2760
      Width           =   1755
   End
   Begin VB.Label lblInternet 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   1
      Left            =   2280
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2520
      Width           =   3045
   End
   Begin VB.Label lblInternet 
      BackStyle       =   0  'Transparent
      Height          =   225
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Tag             =   "2104"
      Top             =   2520
      Width           =   1755
   End
   Begin VB.Label lblCopyright 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   465
      Left            =   375
      TabIndex        =   1
      Top             =   2040
      Width           =   5205
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   368
      TabIndex        =   0
      Top             =   1800
      Width           =   5205
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' [frmAbout]
' Shows credits for FSSC and contact information

Option Explicit

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdSysInfo_Click()
  Dim strSysInfoPath As String
  
  ' Get System Information Path and run it
  strSysInfoPath = RegGetKey("PATH", "", "SOFTWARE\Microsoft\Shared Tools\MSINFO", HKEY_LOCAL_MACHINE)
  If strSysInfoPath = "" Then strSysInfoPath = RegGetKey("MSINFO", "None", "SOFTWARE\Microsoft\Shared Tools Location", HKEY_LOCAL_MACHINE)
  
  If FileExists(strSysInfoPath & "\MSINFO32.EXE") Then _
    strSysInfoPath = strSysInfoPath & "\MSINFO32.EXE"
  
  If FileExists(strSysInfoPath) Then Shell strSysInfoPath, vbNormalFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF11 And Shift = vbAltMask Then
    MsgBoxEx Me, "Congratulations! You found the one and only Easter Egg in Flight Simulator Scenery Creator. E-mail Derek about your achievement!", vbExclamation, 0
  End If
End Sub

Private Sub Form_Load()
  Dim picMouseIcon As IPictureDisp
  CenterForm Me
  DialogMenus Me
  
  Lang.PrepareForm Me
  lblVersion.Caption = Lang.ResolveString(RES_About_Version, App.Major, App.Minor, App.Revision)
  lblCopyright.Caption = Lang.ResolveString(RES_About_Copyright, CopyrightYear)
  lblInternet(1).Caption = Email
  lblInternet(3).Caption = Webpage
  
  Set picMouseIcon = LoadResPicture(1, vbResCursor)
  lblInternet(1).MouseIcon = picMouseIcon
  lblInternet(3).MouseIcon = picMouseIcon
  
  If RegGetKey("", "", "aim", HKEY_CLASSES_ROOT) <> "" Then
    lblInternet(5).MouseIcon = picMouseIcon
  Else
    Set lblInternet(5).Font = lblInternet(7).Font
    lblInternet(5).ForeColor = lblInternet(7).ForeColor
  End If
  Image1 = LoadResPicture(1000, vbResBitmap)
End Sub

Private Sub lblInternet_Click(Index As Integer)
  Dim strRun As String
  ' Execute an internet link
  Select Case Index
    Case 1
      strRun = "mailto:" & Email & "?subject=FS%20Scenery%20Creator%20Comments&body=Hello,"
    Case 3
      strRun = Webpage
    Case 5
      If lblInternet(5).FontUnderline Then
        strRun = "aim:goim?screenname=FSSC+Guru&message=Hello"
      End If
  End Select
  If strRun <> "" Then
    ShellExecute hwnd, "Open", strRun, "", "", vbNormalFocus
  End If
End Sub
