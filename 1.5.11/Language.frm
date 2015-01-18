VERSION 5.00
Begin VB.Form frmLanguage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Language"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   ClipControls    =   0   'False
   Icon            =   "Language.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.ListBox lstLanguage 
      Height          =   2010
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label lblInfo 
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "frmLanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' [frmLanguage]
' Allows the user to select a language for FSSC
Option Explicit

' Open webpage
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private mChanged As Boolean

' Prompts the user for a language
' NoCancel - If true, the user cannot select the
'               cancel button
Public Function EditData(Optional ByVal NoCancel As Boolean = False) As Boolean
  Load frmLanguage
  DialogMenus Me, NoCancel
  lblInfo.Caption = LoadResString(IIf(NoCancel, RES_Lang_DialogBox2, RES_Lang_DialogBox1))
  cmdCancel.Enabled = Not NoCancel
  mChanged = False
  Show vbModal, Screen.ActiveForm
  If mChanged Then
    Lang.Name = lstLanguage.Text
    EditData = True
  End If
  Unload frmLanguage
End Function

Private Sub cmdCancel_Click()
  Hide
End Sub

Private Sub cmdOK_Click()
  mChanged = True
  Hide
End Sub

Private Sub Form_Load()
  CenterForm Me
  Lang.FillList lstLanguage
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    Cancel = True
    Hide
  End If
End Sub

Private Sub lstLanguage_Click()
  If lstLanguage.ItemData(lstLanguage.ListIndex) = 0 Then
    ' "Search for other languages"
    If MsgBoxEx(Me, LoadResString(RES_Lang_Browser), vbOKCancel Or vbExclamation, 0) = vbOK Then
      ShellExecute hwnd, "Open", Webpage & "autolang.html", "", "", vbNormalFocus
    End If
    cmdOK.Enabled = False
  Else
    cmdOK.Enabled = True
  End If
End Sub

Private Sub lstLanguage_DblClick()
  cmdOK_Click
End Sub
