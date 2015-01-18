VERSION 5.00
Begin VB.Form frmAutosave 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6300
   Icon            =   "Autosave.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   Tag             =   "2500"
   Begin VB.CommandButton cmds 
      Cancel          =   -1  'True
      Height          =   375
      Index           =   2
      Left            =   4440
      TabIndex        =   4
      Tag             =   "1033"
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton cmds 
      Height          =   375
      Index           =   1
      Left            =   2640
      TabIndex        =   3
      Tag             =   "2511"
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton cmds 
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   2
      Tag             =   "2510"
      Top             =   2040
      Width           =   1575
   End
   Begin VB.ListBox lstFiles 
      Height          =   1035
      Left            =   840
      TabIndex        =   1
      Top             =   840
      Width           =   5175
   End
   Begin VB.Label lbls 
      Height          =   495
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Tag             =   "2501"
      Top             =   240
      Width           =   5175
   End
End
Attribute VB_Name = "frmAutosave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mResult As Integer

Private Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long

' Shows the Conversion prompt, and returns 1 if a conversion is desired
Public Function DoDialog() As String
  Dim OldValue As MousePointerConstants, File As String, I As Integer
  
  File = Dir$(AddDir(GetTempPathName(), "FSSC*.scn"))
  
  If File <> "" Then
    OldValue = Screen.MousePointer
    Screen.MousePointer = vbDefault
    
    Load frmAutosave
    MessageBeep vbQuestion
    mResult = 2
    
    Do Until File = ""
      lstFiles.AddItem File
      File = Dir$()
    Loop
    lstFiles.ListIndex = 0
    
    Show vbModal, Screen.ActiveForm
    Select Case mResult
      Case 0
        DoDialog = AddDir(GetTempPathName(), lstFiles.Text)
      Case 1
        On Error Resume Next
        Kill AddDir(GetTempPathName(), "FSSC*.scn")
      Case 2
    End Select
    Unload frmAutosave
    Screen.MousePointer = OldValue
  End If
End Function

Private Sub cmds_Click(Index As Integer)
  mResult = Index
  Hide
End Sub

Private Sub Form_Load()
  CenterForm Me
  DialogMenus Me
  Lang.PrepareForm Me
End Sub

Private Sub Form_Paint()
  DrawExclaimIcon hdc, 10, lbls(0).Top / Screen.TwipsPerPixelY
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then Cancel = True: Hide
End Sub

