VERSION 5.00
Begin VB.Form frmSymbols 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   ClipControls    =   0   'False
   Icon            =   "Symbols.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   Tag             =   "2140"
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   3120
      TabIndex        =   4
      Tag             =   "1033"
      Top             =   2640
      WhatsThisHelpID =   1033
      Width           =   1215
   End
   Begin VB.CommandButton cmdInsert 
      Default         =   -1  'True
      Height          =   345
      Left            =   1800
      TabIndex        =   3
      Tag             =   "2150"
      Top             =   2640
      WhatsThisHelpID =   2150
      Width           =   1215
   End
   Begin VB.ListBox lstSymbol 
      Columns         =   12
      Height          =   1620
      Left            =   240
      TabIndex        =   1
      Top             =   480
      WhatsThisHelpID =   2141
      Width           =   4095
   End
   Begin VB.Label lbls 
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      UseMnemonic     =   0   'False
      Width           =   3255
   End
   Begin VB.Label lbls 
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Tag             =   "2141"
      Top             =   240
      WhatsThisHelpID =   2141
      Width           =   3525
   End
End
Attribute VB_Name = "frmSymbols"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSymbol As Integer

' Find a symbol
Public Function GetSymbol() As Integer
  Dim I As Integer
  Const Spacing = "  "

  Load frmSymbols
  lstSymbol.AddItem Spacing & Chr$(128)
  For I = 161 To 255
    lstSymbol.AddItem Spacing & Chr$(I)
  Next I
  Show vbModal, Screen.ActiveForm
  GetSymbol = mSymbol
  Unload frmSymbols
End Function

Private Sub cmdClose_Click()
  mSymbol = 0
  Hide
End Sub

Private Sub cmdInsert_Click()
  mSymbol = lstSymbol.ListIndex + 160
  If mSymbol = 160 Then mSymbol = 128
  Hide
End Sub

Private Sub Form_Load()
  CenterForm Me
  DialogMenus Me
  Lang.PrepareForm Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then Cancel = True: Hide
End Sub

Private Sub lstSymbol_Click()
  Dim Shortcut As String
  On Error Resume Next
  Shortcut = LoadResString(lstSymbol.ListIndex + 160)
  If Shortcut = "" Then
    lbls(1).Caption = ""
  Else
    lbls(1).Caption = Lang.ResolveString(RES_Sym_ShortcutKey, Shortcut)
  End If
End Sub

Private Sub lstSymbol_DblClick()
  cmdInsert_Click
End Sub
