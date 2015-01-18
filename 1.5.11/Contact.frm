VERSION 5.00
Begin VB.Form frmContact 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
   Icon            =   "Contact.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   Tag             =   "2240"
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Tag             =   "1031"
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdContinue 
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Tag             =   "2250"
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txtComment 
      Height          =   1815
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1680
      Width           =   5055
   End
   Begin VB.ComboBox Cmbs 
      Height          =   315
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label lbls 
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   5055
   End
   Begin VB.Label lbls 
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Tag             =   "2242"
      Top             =   870
      Width           =   2055
   End
   Begin VB.Label lbls 
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Tag             =   "2241"
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "frmContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Constants for determining OSPrivate Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

' UDT for determining OS
Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type

Private Const WINDOWS_95 = 0
Private Const WINDOWS_98 = 1
Private Const WINDOWS_NT = 2
Private Const WINDOWS_ME = 3
Private Const WINDOWS_2000 = 4
Private Const WINDOWS_XP = 5

' Call to determine OS version
Private Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer

Public Function GetWindowsVersion() As String
  Dim osinfo As OSVERSIONINFO
  Dim retvalue As Integer

  osinfo.dwOSVersionInfoSize = 148
  osinfo.szCSDVersion = String$(128, 0)
  retvalue = GetVersionExA(osinfo)

  With osinfo
    Select Case .dwPlatformId
      Case VER_PLATFORM_WIN32_WINDOWS
        If .dwMinorVersion = 0 Then
          GetWindowsVersion = "95"
        ElseIf .dwMinorVersion = 10 Then
          GetWindowsVersion = "98"
        ElseIf .dwMinorVersion = 90 Then
          GetWindowsVersion = "ME"
        End If
      Case VER_PLATFORM_WIN32_NT
        If .dwMajorVersion <= 4 Then
          GetWindowsVersion = "NT"
        ElseIf .dwMajorVersion = 5 And .dwMinorVersion = 0 Then
          GetWindowsVersion = "2000"
        ElseIf .dwMajorVersion = 5 And .dwMinorVersion = 1 Then
          GetWindowsVersion = "XP"
        End If
    End Select
  End With
End Function

Private Sub Cmbs_Click()
  lbls(2).Caption = Lang.GetString(RES_Mail_Labels + Cmbs.ListIndex)
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdContinue_Click()
  Dim strRun As String
  strRun = "mailto:" & Email & "?subject=FS%20Scenery%20Creator%20Comments" & "&body=Topic:%20" & Replace(Cmbs.Text, " ", "%20") & "%0AWindows%20Version:%20" & GetWindowsVersion() & "%0ADecimal%20System:%20(" & LocaleInfo(LOCALE_SDECIMAL) & ")%0A%0AComment:%0A" & Replace(Replace(txtComment.Text, " ", "%20"), vbCrLf, "%0A")
  ShellExecute hWnd, "Open", strRun, "", "", vbNormalFocus
End Sub

Private Sub Form_Load()
  CenterForm Me
  DialogMenus Me
  
  Lang.PrepareForm Me
  Lang.AddItems Cmbs, RES_Mail_Reason, 4
  Cmbs.ListIndex = 0
End Sub

