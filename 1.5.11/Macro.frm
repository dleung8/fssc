VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmMacro 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   ClipControls    =   0   'False
   Icon            =   "Macro.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   600
      Top             =   6360
   End
   Begin VB.Frame TabFrame 
      BorderStyle     =   0  'None
      Height          =   5535
      Index           =   3
      Left            =   360
      TabIndex        =   39
      Top             =   600
      Width           =   6000
      Begin VB.CommandButton cmdDefaults 
         Height          =   375
         Left            =   3720
         TabIndex        =   63
         Tag             =   "1474"
         Top             =   5040
         WhatsThisHelpID =   1474
         Width           =   2175
      End
      Begin VB.Frame fraSelector 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   4320
         TabIndex        =   60
         Top             =   120
         Width           =   1575
         Begin VB.CommandButton cmdTexture 
            Height          =   375
            Left            =   360
            TabIndex        =   62
            TabStop         =   0   'False
            Tag             =   "1473"
            Top             =   480
            WhatsThisHelpID =   1473
            Width           =   1215
         End
         Begin VB.CommandButton cmdColor 
            Height          =   375
            Left            =   360
            TabIndex        =   61
            TabStop         =   0   'False
            Tag             =   "1472"
            Top             =   0
            WhatsThisHelpID =   1472
            Width           =   1215
         End
         Begin VB.Image Image1 
            Height          =   240
            Left            =   0
            Picture         =   "Macro.frx":000C
            Stretch         =   -1  'True
            Top             =   30
            Width           =   240
         End
      End
      Begin VB.TextBox PTxts 
         Height          =   285
         Index           =   9
         Left            =   2400
         MaxLength       =   32
         TabIndex        =   59
         Tag             =   "1048"
         Top             =   3630
         WhatsThisHelpID =   1457
         Width           =   1700
      End
      Begin VB.TextBox PTxts 
         Height          =   285
         Index           =   8
         Left            =   2400
         MaxLength       =   32
         TabIndex        =   57
         Tag             =   "1048"
         Top             =   3240
         WhatsThisHelpID =   1457
         Width           =   1700
      End
      Begin VB.TextBox PTxts 
         Height          =   285
         Index           =   7
         Left            =   2400
         MaxLength       =   32
         TabIndex        =   55
         Tag             =   "1048"
         Top             =   2850
         WhatsThisHelpID =   1457
         Width           =   1700
      End
      Begin VB.TextBox PTxts 
         Height          =   285
         Index           =   6
         Left            =   2400
         MaxLength       =   32
         TabIndex        =   53
         Tag             =   "1048"
         Top             =   2460
         WhatsThisHelpID =   1457
         Width           =   1700
      End
      Begin VB.TextBox PTxts 
         Height          =   285
         Index           =   5
         Left            =   2400
         MaxLength       =   32
         TabIndex        =   51
         Tag             =   "1048"
         Top             =   2070
         WhatsThisHelpID =   1457
         Width           =   1700
      End
      Begin VB.TextBox PTxts 
         Height          =   285
         Index           =   4
         Left            =   2400
         MaxLength       =   32
         TabIndex        =   49
         Tag             =   "1048"
         Top             =   1680
         WhatsThisHelpID =   1457
         Width           =   1700
      End
      Begin VB.TextBox PTxts 
         Height          =   285
         Index           =   3
         Left            =   2400
         MaxLength       =   32
         TabIndex        =   47
         Tag             =   "1048"
         Top             =   1290
         WhatsThisHelpID =   1457
         Width           =   1700
      End
      Begin VB.TextBox PTxts 
         Height          =   285
         Index           =   2
         Left            =   2400
         MaxLength       =   32
         TabIndex        =   45
         Tag             =   "1048"
         Top             =   900
         WhatsThisHelpID =   1457
         Width           =   1700
      End
      Begin VB.TextBox PTxts 
         Height          =   285
         Index           =   1
         Left            =   2400
         MaxLength       =   32
         TabIndex        =   43
         Tag             =   "1048"
         Top             =   510
         WhatsThisHelpID =   1457
         Width           =   1700
      End
      Begin VB.TextBox PTxts 
         Height          =   285
         Index           =   0
         Left            =   2400
         MaxLength       =   32
         TabIndex        =   41
         Tag             =   "1048"
         Top             =   120
         WhatsThisHelpID =   1457
         Width           =   1700
      End
      Begin VB.Label Plbls 
         Height          =   390
         Index           =   9
         Left            =   0
         TabIndex        =   58
         Top             =   3660
         WhatsThisHelpID =   1457
         Width           =   2265
      End
      Begin VB.Label Plbls 
         Height          =   390
         Index           =   8
         Left            =   0
         TabIndex        =   56
         Top             =   3270
         WhatsThisHelpID =   1457
         Width           =   2265
      End
      Begin VB.Label Plbls 
         Height          =   390
         Index           =   7
         Left            =   0
         TabIndex        =   54
         Top             =   2880
         WhatsThisHelpID =   1457
         Width           =   2265
      End
      Begin VB.Label Plbls 
         Height          =   390
         Index           =   6
         Left            =   0
         TabIndex        =   52
         Top             =   2490
         WhatsThisHelpID =   1457
         Width           =   2265
      End
      Begin VB.Label Plbls 
         Height          =   390
         Index           =   5
         Left            =   0
         TabIndex        =   50
         Top             =   2100
         WhatsThisHelpID =   1457
         Width           =   2265
      End
      Begin VB.Label Plbls 
         Height          =   390
         Index           =   4
         Left            =   0
         TabIndex        =   48
         Top             =   1710
         WhatsThisHelpID =   1457
         Width           =   2265
      End
      Begin VB.Label Plbls 
         Height          =   390
         Index           =   3
         Left            =   0
         TabIndex        =   46
         Top             =   1320
         WhatsThisHelpID =   1457
         Width           =   2265
      End
      Begin VB.Label Plbls 
         Height          =   390
         Index           =   2
         Left            =   0
         TabIndex        =   44
         Top             =   930
         WhatsThisHelpID =   1457
         Width           =   2265
      End
      Begin VB.Label Plbls 
         Height          =   390
         Index           =   1
         Left            =   0
         TabIndex        =   42
         Top             =   540
         WhatsThisHelpID =   1457
         Width           =   2265
      End
      Begin VB.Label Plbls 
         Height          =   390
         Index           =   0
         Left            =   0
         TabIndex        =   40
         Top             =   150
         WhatsThisHelpID =   1457
         Width           =   2265
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   6360
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      CausesValidation=   0   'False
      Height          =   375
      Left            =   5280
      TabIndex        =   65
      Tag             =   "1031"
      Top             =   6360
      WhatsThisHelpID =   1031
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   64
      Tag             =   "1030"
      Top             =   6360
      WhatsThisHelpID =   1030
      Width           =   1215
   End
   Begin VB.Frame TabFrame 
      BorderStyle     =   0  'None
      Height          =   5535
      Index           =   2
      Left            =   360
      TabIndex        =   29
      Top             =   600
      Width           =   6000
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   2775
         Left            =   3960
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   38
         Top             =   2640
         Width           =   2055
      End
      Begin VB.PictureBox picMacro 
         AutoRedraw      =   -1  'True
         Height          =   1980
         Left            =   3960
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   128
         TabIndex        =   37
         Top             =   405
         Width           =   1980
      End
      Begin VB.CommandButton cmdRefresh 
         Height          =   375
         Left            =   2640
         TabIndex        =   35
         Tag             =   "1471"
         Top             =   5040
         WhatsThisHelpID =   1471
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Height          =   375
         Left            =   1320
         TabIndex        =   34
         Tag             =   "1470"
         Top             =   5040
         WhatsThisHelpID =   1470
         Width           =   1095
      End
      Begin VB.CommandButton cmdBrowse 
         Height          =   375
         Left            =   0
         TabIndex        =   33
         Tag             =   "1032"
         Top             =   5040
         WhatsThisHelpID =   1032
         Width           =   1095
      End
      Begin ComctlLib.TreeView lstMacros 
         Height          =   4215
         Left            =   0
         TabIndex        =   32
         Top             =   720
         WhatsThisHelpID =   1450
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   7435
         _Version        =   327682
         HideSelection   =   0   'False
         Indentation     =   0
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
      End
      Begin VB.TextBox txtMacro 
         Height          =   285
         Left            =   0
         TabIndex        =   31
         Top             =   405
         WhatsThisHelpID =   1450
         Width           =   3735
      End
      Begin VB.Label lbls 
         Height          =   255
         Index           =   16
         Left            =   3960
         TabIndex        =   36
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lbls 
         Height          =   255
         Index           =   15
         Left            =   0
         TabIndex        =   30
         Tag             =   "1450"
         Top             =   120
         WhatsThisHelpID =   1450
         Width           =   1695
      End
   End
   Begin VB.Frame TabFrame 
      BorderStyle     =   0  'None
      Height          =   5535
      Index           =   1
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   6000
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   10
         Left            =   1320
         TabIndex        =   26
         Tag             =   "1455"
         Top             =   4650
         WhatsThisHelpID =   1455
         Width           =   1095
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   9
         Left            =   1320
         TabIndex        =   24
         Tag             =   "1454"
         Top             =   4260
         WhatsThisHelpID =   1454
         Width           =   1095
      End
      Begin VB.ComboBox Cmbs 
         Height          =   315
         Index           =   0
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   5025
         WhatsThisHelpID =   1049
         Width           =   2415
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   8
         Left            =   1320
         TabIndex        =   22
         Tag             =   "1453"
         Top             =   3870
         WhatsThisHelpID =   1453
         Width           =   1095
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   7
         Left            =   1320
         TabIndex        =   19
         Tag             =   "1452"
         Top             =   3480
         WhatsThisHelpID =   1452
         Width           =   1095
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   6
         Left            =   1320
         TabIndex        =   17
         Tag             =   "1451"
         Top             =   3090
         WhatsThisHelpID =   1451
         Width           =   1095
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   5
         Left            =   1320
         TabIndex        =   15
         Tag             =   "1048"
         Top             =   2700
         WhatsThisHelpID =   1048
         Width           =   1095
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   4
         Left            =   1800
         TabIndex        =   13
         Top             =   2310
         WhatsThisHelpID =   1045
         Width           =   1935
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   3
         Left            =   1800
         TabIndex        =   11
         Top             =   1920
         WhatsThisHelpID =   1045
         Width           =   1935
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   2
         Left            =   3120
         TabIndex        =   8
         Tag             =   "1044"
         Top             =   1200
         WhatsThisHelpID =   1042
         Width           =   1095
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   6
         Tag             =   "1043"
         Top             =   1200
         WhatsThisHelpID =   1042
         Width           =   1095
      End
      Begin VB.CheckBox chkLocked 
         Height          =   195
         Left            =   0
         TabIndex        =   3
         Tag             =   "1041"
         Top             =   600
         WhatsThisHelpID =   1041
         Width           =   2655
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   2
         Top             =   120
         WhatsThisHelpID =   1040
         Width           =   3135
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   10
         Left            =   2520
         TabIndex        =   20
         Tag             =   "1091"
         Top             =   3510
         Width           =   825
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   14
         Left            =   0
         TabIndex        =   27
         Tag             =   "1049"
         Top             =   5070
         WhatsThisHelpID =   1049
         Width           =   1185
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   13
         Left            =   0
         TabIndex        =   25
         Tag             =   "1455"
         Top             =   4680
         WhatsThisHelpID =   1455
         Width           =   1185
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   12
         Left            =   0
         TabIndex        =   23
         Tag             =   "1454"
         Top             =   4290
         WhatsThisHelpID =   1454
         Width           =   1185
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   11
         Left            =   0
         TabIndex        =   21
         Tag             =   "1453"
         Top             =   3900
         WhatsThisHelpID =   1453
         Width           =   1185
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   9
         Left            =   0
         TabIndex        =   18
         Tag             =   "1452"
         Top             =   3510
         WhatsThisHelpID =   1452
         Width           =   1185
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   8
         Left            =   0
         TabIndex        =   16
         Tag             =   "1451"
         Top             =   3120
         WhatsThisHelpID =   1451
         Width           =   1185
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   7
         Left            =   0
         TabIndex        =   14
         Tag             =   "1048"
         Top             =   2730
         WhatsThisHelpID =   1048
         Width           =   1185
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   6
         Left            =   600
         TabIndex        =   12
         Tag             =   "1047"
         Top             =   2340
         WhatsThisHelpID =   1045
         Width           =   1065
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   5
         Left            =   600
         TabIndex        =   10
         Tag             =   "1046"
         Top             =   1950
         WhatsThisHelpID =   1045
         Width           =   1065
      End
      Begin VB.Label lbls 
         Height          =   285
         Index           =   4
         Left            =   360
         TabIndex        =   9
         Tag             =   "1045"
         Top             =   1635
         WhatsThisHelpID =   1045
         Width           =   2175
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   3
         Left            =   2640
         TabIndex        =   7
         Tag             =   "1044"
         Top             =   1230
         WhatsThisHelpID =   1042
         Width           =   345
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   2
         Left            =   600
         TabIndex        =   5
         Tag             =   "1043"
         Top             =   1230
         WhatsThisHelpID =   1042
         Width           =   345
      End
      Begin VB.Label lbls 
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Tag             =   "1042"
         Top             =   915
         WhatsThisHelpID =   1042
         Width           =   2175
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Tag             =   "1040"
         Top             =   150
         WhatsThisHelpID =   1040
         Width           =   945
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   6135
      Left            =   120
      TabIndex        =   66
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   10821
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "A"
            Object.Tag             =   "1460"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "B"
            Object.Tag             =   "1461"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "C"
            Object.Tag             =   "1462"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMacro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

' TreeView bug
' NodeClick event fires after a doubleclick in the open
' file dialog box
Private BugFirstTime As Boolean, BugTimer As Long

' TreeView bug
' Clicking Tooltips does not select the item below.
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, bEnable As Long) As Long

' treeview messages
Private Const TV_FIRST = &H1100
Private Const TVM_GETTOOLTIPS = (TV_FIRST + 25)

Private mValueCache(10) As Single

Private mTxtChanged As Boolean
Private mChanged As Boolean

Private NewTop As Single

Private MacroFile As String

Private SetDefaults As Boolean

Private RGBNoExtend As Integer

' 0 = Range
' 1 = Scale
' 2 = V1
' 3 = Altitude
' 4 = Density
' 5 = V2
Private DefaultParams(5) As String
Private Defaults(9) As String

Private Sub AddMacroList()
  Dim I As Integer, J As Integer, Hierarchy() As String, _
    Star As String, PathName As String

  Dim mItem As Node, mItem2 As Node
  
  On Error Resume Next
  
  SetScreenMousePointer vbHourglass
  Set lstMacros.ImageList = frmMain.imgTreeView
  
  lstMacros.Nodes.Clear
  
  For I = 1 To UBound(MacroLst)
    Hierarchy = Split(MacroLst(I).DirName, "\>")
    Set mItem = lstMacros.Nodes("\" & Hierarchy(0))
    If Err.Number > 0 Then
      If Mid$(Hierarchy(0), 2, 1) = ":" Then
        Set mItem = lstMacros.Nodes.Add(, , "\" & Hierarchy(0), Hierarchy(0), 5)
      Else
        Set mItem = lstMacros.Nodes.Add(, , "\" & Hierarchy(0), Hierarchy(0), 1, 2)
      End If
      Err.Number = 0
    End If
    PathName = "\" & Hierarchy(0)
    For J = 1 To UBound(Hierarchy)
      PathName = PathName & "\" & Hierarchy(J)
      Set mItem = lstMacros.Nodes(PathName)
      If Err.Number > 0 Then
        Set mItem = lstMacros.Nodes.Add(mItem, tvwChild, PathName, Hierarchy(J), 1, 2)
        Err.Number = 0
      End If
    Next J
    For J = 1 To UBound(MacroLst(I).Data)
      'Star = "*"
      'Do
'        Err.Number = 0
        lstMacros.Nodes.Add(mItem, tvwChild, MacroLst(I).Data(J).File & Star, MacroLst(I).Data(J).Name, 3).Tag = CStr(I) & " " & CStr(J)
      '  If Err.Number > 0 Then
      '    Star = Star & "*"
      '  End If
      'Loop Until Err.Number = 0
    Next J
    Do
      Set mItem2 = mItem
      Set mItem = mItem.Parent
      If mItem2.Children = 0 Then
        lstMacros.Nodes.Remove mItem2.Key
      Else
        Exit Do
      End If
    Loop Until mItem Is Nothing
    Set mItem = Nothing
    Set mItem2 = Nothing
  Next I
  
  SetScreenMousePointer vbDefault
End Sub

Private Function ConvertToRegion(ByVal Data As String) As String
  Dim Temp As String
  If InStr(Data, ".") = 0 Then
    ConvertToRegion = Data
  ElseIf StrComp(UCase$(Data), LCase$(Data), vbBinaryCompare) = 0 Then
    Temp = Data
    Mid$(Temp, InStrRev(Data, "."), 1) = LocaleInfo(LOCALE_SDECIMAL)
    ConvertToRegion = Temp
  Else
    ConvertToRegion = Data
  End If
End Function

Private Function ConvertToStandard(ByVal Data As String) As String
  Dim DecimalStr As String, Temp As String
  DecimalStr = LocaleInfo(LOCALE_SDECIMAL)
  
  If InStr(Data, DecimalStr) = 0 Then
    ConvertToStandard = Data
  ElseIf StrComp(UCase$(Data), LCase$(Data), vbBinaryCompare) = 0 Then
    Temp = Data
    Mid$(Temp, InStrRev(Data, DecimalStr), 1) = "."
    ConvertToStandard = Temp
  Else
    ConvertToStandard = Data
  End If
End Function

Public Function EditData(Data As clsMacro) As Boolean
  Dim I As Integer
  
  Load frmMacro

  With Data
  
    ResetAll

    On Error Resume Next
    Set lstMacros.SelectedItem = lstMacros.Nodes(.File)

    If Err Then
      txtMacro.Text = .File
      txtMacro_Change
    Else
      lstMacros.SelectedItem.EnsureVisible
      lstMacros_NodeClick lstMacros.SelectedItem
    End If
    
    On Error GoTo 0

    Txts(0).Tag = .Caption(True)
    Txts(0).Text = .Name
    Txts_Change 0 ' (Changes the caption)
    Txts(1).Text = MeterToUser(.X)
    Txts(2).Text = MeterToUser(.Y)
    Txts_Validate 1, False ' (Updates Latitude, Longitude)
    chkLocked.Value = -.Locked
    Txts(5).Text = GeographicToUser(.Rotation)
    Txts(6).Text = NauticalToUser(.Range)
    Txts(7).Text = MeterToUser(.MScale, "##0.0#########")
    Txts(8).Text = MeterToUser(.Altitude)
    Txts(9).Text = MeterToUser(.V1)
    Txts(10).Text = MeterToUser(.V2)
    
    For I = 0 To 9
      PTxts(I).Text = ConvertToRegion(.Param(I))
    Next I
    
    Cmbs(0).ListIndex = .Complexity
    
    SetDefaults = .SceneryIndex = 0

    mChanged = False
    If TabValue > 2 Then TabValue = 0
    Set TabStrip1.SelectedItem = TabStrip1.Tabs(TabValue + 1)
    Show vbModal, Screen.ActiveForm
    If mChanged Then
      .Name = Txts(0).Text
      .X = mValueCache(1)
      .Y = mValueCache(2)
      .Locked = -chkLocked.Value
      .Rotation = mValueCache(5)
      .Range = mValueCache(6)
      .MScale = mValueCache(7)
      .Altitude = mValueCache(8)
      .V1 = mValueCache(9)
      .V2 = mValueCache(10)
      .Complexity = Cmbs(0).ListIndex
      
      If FileExists(txtMacro.Text) Then
        .File = txtMacro.Text
      ElseIf Not lstMacros.SelectedItem Is Nothing Then
        .File = lstMacros.SelectedItem.Key
      Else
        .File = "Invalid: " & txtMacro.Text ' Temporary solution
      End If
      .FileExistsCheck = IIf(FileExists(.File), MFILEEXIST, MFILENOTEXIST)
      
      For I = 0 To 9
        .Param(I) = ConvertToStandard(PTxts(I).Text)
      Next I
      EditData = True
    End If
  End With

  Unload frmMacro
End Function

Public Function EditDataM(Multi() As clsObject) As Boolean
  Dim I As Integer, J As Integer, _
    Data As clsMacro, Temp As clsMacro
  Dim IgnoreValue(22) As Boolean

  ' Keeps track of which values need to be stored, and
  ' which are skipped

  ' When loading to the form, a false value means
  '   fill the control, true means make the control
  '   indeterminate (blank)
  ' When recording changes, a false value means store
  '   the value, true means do not store the value

  Set Data = Multi(1)

  For I = 2 To UBound(Multi)
    Set Temp = Multi(I)
    With Temp
      IgnoreValue(0) = IgnoreValue(0) Or Data.Name <> .Name
      IgnoreValue(1) = IgnoreValue(1) Or Data.X <> .X
      IgnoreValue(2) = IgnoreValue(2) Or Data.Y <> .Y
      IgnoreValue(3) = IgnoreValue(3) Or Data.Locked <> .Locked
      IgnoreValue(5) = IgnoreValue(5) Or Data.Rotation <> .Rotation
      IgnoreValue(6) = IgnoreValue(6) Or Data.Range <> .Range
      IgnoreValue(7) = IgnoreValue(7) Or Data.MScale <> .MScale
      IgnoreValue(8) = IgnoreValue(8) Or Data.Altitude <> .Altitude
      IgnoreValue(9) = IgnoreValue(9) Or Data.V1 <> .V1
      IgnoreValue(10) = IgnoreValue(10) Or Data.V2 <> .V2
      For J = 11 To 20
        IgnoreValue(J) = IgnoreValue(J) Or Data.Param(J - 11) <> .Param(J - 11)
      Next J
      IgnoreValue(21) = IgnoreValue(21) Or Data.Complexity <> .Complexity
      IgnoreValue(22) = IgnoreValue(22) Or Data.File <> .File
    End With
    Set Temp = Nothing
  Next I

  Load frmMacro
  MultiSelection = True
  

  ' Fill Data
  With Data
  
    If Not IgnoreValue(22) Then
      On Error Resume Next
      Set lstMacros.SelectedItem = lstMacros.Nodes(.File)
  
      If Err Then
        txtMacro.Text = .File
        txtMacro_Change
      Else
        lstMacros.SelectedItem.EnsureVisible
        lstMacros_NodeClick lstMacros.SelectedItem
      End If
      
      On Error GoTo 0
    End If
    
    ResetAll

    If Not IgnoreValue(0) Then Txts(0).Text = .Name
    Txts(0).Tag = Lang.GetString(RES_Obj_Macro)
    Txts_Change 0
    If Not IgnoreValue(1) Then Txts(1).Text = MeterToUser(.X)
    If Not IgnoreValue(2) Then Txts(2).Text = MeterToUser(.Y)
    Txts_Validate 1, False
    If Not IgnoreValue(3) Then
      chkLocked.Value = -.Locked
    Else
      chkLocked.Value = vbGrayed
    End If
    If Not IgnoreValue(5) Then Txts(5).Text = GeographicToUser(.Rotation)
    If Not IgnoreValue(6) Then Txts(6).Text = NauticalToUser(.Range)
    If Not IgnoreValue(7) Then Txts(7).Text = MeterToUser(.MScale, "##0.0#########")
    If Not IgnoreValue(8) Then Txts(8).Text = MeterToUser(.Altitude)
    If Not IgnoreValue(9) Then Txts(9).Text = MeterToUser(.V1)
    If Not IgnoreValue(10) Then Txts(10).Text = MeterToUser(.V2)
    
    For J = 11 To 20
      If Not IgnoreValue(J) Then
        PTxts(J - 11).Text = ConvertToRegion(.Param(J - 11))
      End If
    Next J
    
    If Not IgnoreValue(21) Then Cmbs(0).ListIndex = .Complexity
  End With

  mChanged = False
  If TabValue > 2 Then TabValue = 0
  Set TabStrip1.SelectedItem = TabStrip1.Tabs(TabValue + 1)
  Show vbModal, Screen.ActiveForm
  If mChanged Then
    IgnoreValue(0) = Txts(0).Text = ""
    IgnoreValue(1) = Txts(1).Text = ""
    IgnoreValue(2) = Txts(2).Text = ""
    IgnoreValue(3) = chkLocked.Value = vbGrayed
    IgnoreValue(5) = Txts(5).Text = ""
    IgnoreValue(6) = Txts(6).Text = ""
    IgnoreValue(7) = Txts(7).Text = ""
    IgnoreValue(8) = Txts(8).Text = ""
    IgnoreValue(9) = Txts(9).Text = ""
    IgnoreValue(10) = Txts(10).Text = ""
    For J = 11 To 20
      IgnoreValue(J) = PTxts(J - 11).Text = ""
    Next J
    IgnoreValue(21) = Cmbs(0).ListIndex = -1
    IgnoreValue(22) = txtMacro.Text = ""

    For I = 1 To UBound(Multi)
      Set Temp = Multi(I)
      With Temp
        If Not IgnoreValue(0) Then .Name = Txts(0).Text
        If Not IgnoreValue(1) Then .X = mValueCache(1)
        If Not IgnoreValue(2) Then .Y = mValueCache(2)
        If Not IgnoreValue(3) Then .Locked = -chkLocked.Value
        If Not IgnoreValue(5) Then .Rotation = mValueCache(5)
        If Not IgnoreValue(6) Then .Range = mValueCache(6)
        If Not IgnoreValue(7) Then .MScale = mValueCache(7)
        If Not IgnoreValue(8) Then .Altitude = mValueCache(8)
        If Not IgnoreValue(9) Then .V1 = mValueCache(9)
        If Not IgnoreValue(10) Then .V2 = mValueCache(10)
        For J = 11 To 20
          If Not IgnoreValue(J) Then
            .Param(J - 11) = ConvertToStandard(PTxts(J - 11).Text)
          End If
        Next J
        If Not IgnoreValue(21) Then .Complexity = Cmbs(0).ListIndex
        If Not IgnoreValue(22) Then
          If FileExists(txtMacro.Text) Then
            .File = txtMacro.Text
            .FileExistsCheck = MFILEEXIST
          Else
            .File = lstMacros.SelectedItem.Key
            .FileExistsCheck = MFILENOTEXIST
          End If
        End If
      End With
    Next I
    EditDataM = True
  End If

  Unload frmMacro
  MultiSelection = False
End Function

Private Sub ParseASDMacro(File As String, ByRef bmpFile As String)
  Dim X As Integer, Temp As String, Token As String, _
    Key As String, Data As String, I As Long, _
    J As Long, Slash As Integer
  X = FreeFile
    
  For I = 0 To 9
    Plbls(I).Caption = Lang.ResolveString(RES_Macro_Params, I + 4)
  Next I
  
  J = -999
  Open File For Input As #X
  LineInputEx X, Temp
  If Trim$(Temp) = ";ASDesign Compatible Macro" Then
    Do Until EOF(X) Or (I > 25)
      I = I + 1
      LineInputEx X, Temp
      Temp = Trim$(Temp)
      If Right$(Temp, 1) = "\" Then
        Temp = Trim$(Left$(Temp, Len(Temp) - 1))
        Slash = 1
      ElseIf Slash = 1 Then
        Slash = 2
      ElseIf Slash = 2 Then
        Exit Do
      End If
      If Left$(Temp, 1) = ";" Then
        Temp = Mid$(Temp, 2)
        Do Until Temp = ""
          Token = Trim$(ReadNext(Temp, ","))
          Key = Trim$(ReadNext(Token, "="))
          If Key <> "" Then J = J + 1
          Data = Token
          Select Case Key
            Case "Name"
              txtInfo.Text = Data
            Case "Bitmap"
              bmpFile = Data
            Case "Latitude"
              J = 1
            Case "Longitude"
            Case "Rotation"
            Case "Range"
              Txts(6).Text = Data
            Case "Scale"
              DefaultParams(1) = ConvertToRegion(Data)
            Case "Elevation"
              DefaultParams(2) = ConvertToRegion(Data)
            Case "Visibility"
              DefaultParams(3) = Data
            Case "Density"
              DefaultParams(4) = Data
            Case Else
              If Between(J, 4, 13) Then
                Plbls(J - 4).Caption = Key & ":"
                Defaults(J - 4) = Data
              End If
          End Select
        Loop
      End If
    Loop
  End If
End Sub

Private Sub ParseMacroDefaults(MetaType As MacroMetaType)
  Dim Meta As String, Param As String, I As Long
  
  Meta = MetaType.Defaults
  If Meta <> "" Then
    For I = 0 To 9
      Param = ReadNext(Meta, ";")
      If Param <> "" Then Defaults(I) = Trim$(ConvertToRegion(Param))
    Next I
  Else
    Meta = MetaType.DefaultParams
    If Meta <> "" Then
      For I = 0 To 9
        Param = ReadNext(Meta, ",")
        If Param <> "" Then Defaults(I) = Trim$(ConvertToRegion(Param))
      Next I
    End If
  End If

  Select Case MetaType.DefaultDensity
    Case "very dense": DefaultParams(4) = CStr(4)
    Case "dense":      DefaultParams(4) = CStr(3)
    Case "normal":     DefaultParams(4) = CStr(2)
    Case "sparse":     DefaultParams(4) = CStr(1)
    Case Else:         DefaultParams(4) = CStr(0)
  End Select
  
  If MetaType.DefaultRange <> "" Then DefaultParams(3) = MetaType.DefaultRange
End Sub

Private Sub ParseMacroDesc(MetaType As MacroMetaType)
  Dim Desc As String, Pos As Integer, Scl As Single
  
  Desc = MetaType.MacroDesc
  
  DefaultParams(1) = ""
  
  Pos = InStr(Desc, "scale=")
  If Pos > 0 Then
    DefaultParams(1) = LTrim$(Mid$(Desc, Pos + 6))
  Else
    Pos = InStr(Desc, "scale =")
    If Pos > 0 Then
      DefaultParams(1) = LTrim$(Mid$(Desc, Pos + 7))
    Else
      Pos = InStr(Desc, "using scale ")
      If Pos > 0 Then
        DefaultParams(1) = LTrim$(Mid$(Desc, Pos + 11))
      End If
    End If
  End If
  If DefaultParams(1) = "" Then DefaultParams(1) = MetaType.MScale
  If DefaultParams(1) = "" Then DefaultParams(1) = MetaType.DefaultScale
  If DefaultParams(1) = "" Then DefaultParams(1) = MetaType.AirportScale
  If DefaultParams(1) = "" Then DefaultParams(1) = CStr(1)
  txtInfo.Text = Replace(Desc, "\n", vbCrLf)
End Sub

Private Sub ParseMacroParam(MetaType As MacroMetaType)
  Dim Meta As String, Param As String, I As Integer, TempStr
  Meta = MetaType.ParamDescr
  If Meta <> "" Then
    For I = 0 To 9
      Param = ReadNext(Meta, ",")
      If Param <> "" Then
        Plbls(I).Caption = Param & ":"
      Else
        SetEnabled PTxts(I), False
      End If
    Next I
  Else
    Meta = MetaType.VODData
    
    If Meta <> "" Then
      ' Meta data exists for macro, read properties
      ReadNext Meta, " "
      ReadNext Meta, " "
      SetEnabled Txts(6), (ReadNext(Meta, " ") <> "not_used")
      SetEnabled Txts(7), (ReadNext(Meta, " ") <> "not_used")
      SetEnabled Txts(5), (ReadNext(Meta, " ") = "rotation")
      If Txts(5).Locked Then Txts(5).Text = GeographicToUser(0)
        
      For I = 0 To 3
        Param = ReadNext(Meta, " ")
        If Param = "not_used" Then
          SetEnabled PTxts(I), False
        ElseIf Param = "" Then
        Else
          Plbls(I).Caption = Replace(Param, "_", " ") & ":"
        End If
      Next I
      SetEnabled Txts(9), (ReadNext(Meta, " ") <> "not_used")
      
      TempStr = ReadNext(Meta, " ")
      SetEnabled Txts(8), (TempStr = "altitude" Or TempStr = "elevation")
      If InStr(ReadNext(Meta, " "), "complexity") = 0 Then
        Cmbs(0).BackColor = vbButtonFace
        Cmbs(0).Locked = True
        Cmbs(0).ListIndex = 0
      End If
      ReadNext Meta, " "
      SetEnabled Txts(10), (ReadNext(Meta, " ") <> "not_used")
      For I = 4 To 9
        Param = ReadNext(Meta, " ")
        If Param = "not_used" Then
          SetEnabled PTxts(I), False
        ElseIf Param = "" Then
        Else
          Plbls(I).Caption = Replace(Param, "_", " ") & ":"
        End If
      Next I
    End If
  End If

  If MetaType.RGBEnabled = "True" Or MetaType.RGBEnabled = "Yes" Or MetaType.RGBEnabled = "1" Then RGBNoExtend = 1
End Sub

' Clear all the text boxes and option buttons
Private Sub ResetAll()
  Dim I As Integer
  Cmbs(0).BackColor = vbWindowBackground
  Cmbs(0).Locked = False
  For I = 5 To 10
    SetEnabled Txts(I), True
  Next I
  For I = 0 To 9
    SetEnabled PTxts(I), True
    PTxts(I).Text = ""
  Next I
  For I = 0 To 3
    Plbls(I).Caption = Lang.ResolveString(RES_Macro_Params, I + 6)
    Defaults(I) = ""
  Next I
  For I = 4 To 9
    Plbls(I).Caption = Lang.ResolveString(RES_Macro_Params, I + 11)
    Defaults(I) = ""
  Next I
  DefaultParams(0) = "20"
  DefaultParams(1) = "1"
  DefaultParams(2) = "0"
  DefaultParams(3) = "10000"
  DefaultParams(4) = "0"
  DefaultParams(5) = "100"
  RGBNoExtend = 2
End Sub

Private Sub chkLocked_Click()
  Dim Value As Boolean
  Value = Not -chkLocked.Value
  SetEnabled Txts(1), Value
  SetEnabled Txts(2), Value
  SetEnabled Txts(3), Value
  SetEnabled Txts(4), Value
End Sub

Private Sub cmdBrowse_Click()
  Dim X As String, OldDir As String
  Static Dirs As String

  ' Ask the user for the location of the executable file
  With cDialog
    .Filter = Lang.GetString(RES_MacroFilter)
    .FilterIndex = 1
    .DefExt = "api"
    OldDir = CurDir$
    If Dirs = "" Then Dirs = Left$(Options.MacroPath, InStr(Options.MacroPath, ";") - 1)
    ChangeDir Dirs
    X = .OpenDialog(Lang.GetString(RES_Macro_SelectCaption))
    Dirs = CurDir$
    ChangeDir OldDir
    DoEvents
    If X <> "" Then
      Set lstMacros.SelectedItem = Nothing
      txtMacro.Text = GetRealName(X)
      BugFirstTime = True
      BugTimer = Timer
    End If
  End With
End Sub

Private Sub cmdCancel_Click()
  Hide
End Sub

Private Sub cmdColor_Click()
  Dim X As Long, Index As Integer, Res As String
  Index = Val(fraSelector.Tag)
  
  X = ValEx(PTxts(Index).Text)
  If X = 0 Then
    X = 0
  ElseIf X < 256 Then
    X = X Or PalMask
  Else
    X = (X - 256) Or RGBMask
  End If

  If frmColor.EditData(X, RGBNoExtend) Then
    If X And PalMask Then
      ' Palette
      Res = X And PalColormask
    ElseIf X = 0 Then
      ' None
      Res = ""
    Else
      ' RGB
      Res = (X And (RGBColorMask Or TransparentMask)) + 256
    End If
    PTxts(Index).Text = Res
  End If
  PTxts(Index).SetFocus
End Sub

Private Sub cmdDefaults_Click()
  Dim I As Integer
  
  Txts(6).Text = NauticalToUser(ValEx(DefaultParams(0)))
  Txts(7).Text = MeterToUser(ValEx(DefaultParams(1)), "##0.0#########")
  Txts(8).Text = MeterToUser(ValEx(DefaultParams(2)))
  Txts(9).Text = MeterToUser(ValEx(DefaultParams(3)))
  Cmbs(0).ListIndex = Val(DefaultParams(4))
  Txts(10).Text = MeterToUser(ValEx(DefaultParams(5)))
  
  For I = 0 To 9
    PTxts(I).Text = ConvertToRegion(Defaults(I))
  Next I
End Sub

Private Sub cmdEdit_Click()
  If FileExists(MacroFile) Then
    On Error Resume Next
    Shell Options.TextEditor & " " & QuoteString(MacroFile), vbNormalFocus
    If Err.Number > 0 Then MsgBoxEx Me, Lang.ResolveString(RES_ERR_MacroOpen, Error$), vbCritical, RES_ERR_MacroOpen
  End If
End Sub

Private Sub cmdOK_Click()
  Dim I As Integer, Msg As String, Cancel As Boolean
  
  ' Validate event not fired when Enter key pressed
  ' bug workaround
  If TypeOf ActiveControl Is TextBox Then
    If ActiveControl.Name <> "txtMacro" Then
      Txts_Validate ActiveControl.Index, Cancel
      If Cancel Then Exit Sub
    End If
  End If
  
  For I = 1 To 10
    If Not Validate(Txts(I), Msg, mValueCache(I)) Then _
      GoTo ValidationError:
  Next I
  mChanged = True
  Hide
  Exit Sub
ValidationError:
  Set TabStrip1.SelectedItem = TabStrip1.Tabs(Chr$(Txts(I).Container.Index + 64))
  MsgBoxEx Me, Msg, vbCritical, RES_ERR_Bound
  Txts(I).SetFocus
  Exit Sub
End Sub

Private Sub cmdRefresh_Click()
  Dim X As String
  On Error Resume Next
  If FileExists(txtMacro.Text) Then
    X = txtMacro.Text
  ElseIf Not lstMacros.SelectedItem Is Nothing Then
    X = lstMacros.SelectedItem.Key
  Else
    X = txtMacro.Text
  End If
  
  LoadMacros
  AddMacroList
    
  Set lstMacros.SelectedItem = lstMacros.Nodes(X)

  If Err Then
    txtMacro.Text = X
    txtMacro_Change
  Else
    lstMacros.SelectedItem.EnsureVisible
    lstMacros_NodeClick lstMacros.SelectedItem
  End If
End Sub

Private Sub cmdTexture_Click()
  Dim Ans As String, Index As Integer
  Index = Val(fraSelector.Tag)
  
  Ans = PTxts(Index).Text
  If Ans <> "" Then Ans = "N " & Ans
  If frmTexture.EditData(Ans, Tex_File) Then
    PTxts(Index).Text = GetFileTitle(Ans)
  End If
  PTxts(Index).SetFocus
End Sub

Private Sub Form_Activate()
  Dim hwndTT As Long
  
  ' Treeview Bug workaround
  ' Get the handle of the TreeView's tooltip window
  hwndTT = SendMessageLong(lstMacros.hwnd, TVM_GETTOOLTIPS, 0, 0)
  If hwndTT Then EnableWindow hwndTT, 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Handle Ctrl+ [Shift + ] {TAB}s
  If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
    Set TabStrip1.SelectedItem = TabStrip1.Tabs(TabStrip1.SelectedItem.Index Mod TabStrip1.Tabs.Count + 1)
  ElseIf Shift = (vbCtrlMask Or vbShiftMask) And KeyCode = vbKeyTab Then
    Set TabStrip1.SelectedItem = TabStrip1.Tabs((TabStrip1.SelectedItem.Index + TabStrip1.Tabs.Count - 2) Mod TabStrip1.Tabs.Count + 1)
  End If
End Sub

Private Sub Form_Load()
  Dim I As Integer
  
  CenterForm Me
  DialogMenus Me

  For I = 1 To TabFrame.Count
    With TabFrame(I)
      .Move TabFrame(1).Left, TabFrame(1).Top, TabFrame(1).Width, TabFrame(1).Height
      .Enabled = False
      .Visible = False
    End With
  Next I

  Lang.PrepareForm Me
  
  Lang.AddItems Cmbs(0), RES_Complexity1, IIf(Options.FSVersion < Version_FS2K, 5, 6)
  
  AddMacroList
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then Cancel = True: Hide
End Sub

Private Sub lstMacros_NodeClick(ByVal Node As ComctlLib.Node)
  If BugFirstTime And Timer - BugTimer < 1 Then
    Set lstMacros.SelectedItem = Nothing
  Else
    txtMacro.Text = Node.Text
  End If
End Sub

Private Sub PTxts_GotFocus(Index As Integer)
  NewTop = PTxts(Index).Top
  fraSelector.Tag = Index
  Timer2.Enabled = True
  SelectText PTxts(Index)
End Sub

Private Sub TabStrip1_BeforeClick(Cancel As Integer)
  Timer1.Enabled = True
End Sub

Private Sub TabStrip1_Click()
  ' Handle Tab clicks
  Dim SelItem As Integer
  SelItem = Asc(TabStrip1.SelectedItem.Key) - 64
  
  If TabStrip1.Tag <> "" Then
    If SelItem <> TabStrip1.Tag Then
      With TabFrame(TabStrip1.Tag)
        .Visible = False
        .Enabled = False
      End With
    End If
  End If
  
  With TabFrame(SelItem)
    .Enabled = True
    .Visible = True
  End With
  
  If TabStrip1.Tag <> "" And TabStrip1.Visible Then
    If Not ActiveControl Is TabStrip1 Then
      If TabStrip1.Tag <> SelItem Then
        Select Case SelItem
          Case 1: Txts(0).SetFocus
          Case 2
            ' Treeview Bug workaround
            Dim hwndTT As Long
            
            ' Get the handle of the TreeView's tooltip window
            hwndTT = SendMessage(lstMacros.hwnd, TVM_GETTOOLTIPS, 0, 0)
            If hwndTT Then EnableWindow hwndTT, 1
          
            txtMacro.SetFocus
          Case 3: PTxts(0).SetFocus
        End Select
      End If
    End If
  End If
  
  TabStrip1.Tag = SelItem
End Sub

Private Sub Timer1_Timer()
  TabStrip1_Click
  Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
  Dim OldTop As Single
  OldTop = fraSelector.Top
  If Abs(OldTop - NewTop) < 120 Then
    fraSelector.Top = NewTop
  ElseIf OldTop - NewTop > 0 Then
    fraSelector.Top = OldTop - 120
  Else
    fraSelector.Top = OldTop + 120
  End If
End Sub

Private Sub txtMacro_Change()
  Dim I As Long, J As Long, _
    Meta As MacroMetaType, TempStr As String, _
    bmpFile As String, bmpTitle As String
  
  If FileExists(txtMacro.Text) Then
    MacroFile = txtMacro.Text
    cmdOK.Enabled = True
  ElseIf Not lstMacros.SelectedItem Is Nothing Then
    If Left$(lstMacros.SelectedItem.Key, 1) = "\" Then
      cmdOK.Enabled = False
      MacroFile = ""
    Else
      MacroFile = lstMacros.SelectedItem.Key
      cmdOK.Enabled = True
    End If
  Else
    cmdOK.Enabled = False
    MacroFile = ""
  End If
  
  ResetAll
  
  I = InStr(MacroFile, "*")
  If I > 0 Then MacroFile = Left$(MacroFile, I - 1)
  
  If Left$(MacroFile, 8) = "LibObj: " Then
    TempStr = lstMacros.SelectedItem.Tag
    I = Val(ReadNext(TempStr, " "))
    J = Val(TempStr)
    If MacroLst(I).Data(J).Bitmap <> "" Then
      bmpFile = MultiDir(MacroLst(I).Data(J).Bitmap, Options.MacroPicPath)
    End If
    DefaultParams(1) = MacroLst(I).Data(J).MScale
    DefaultParams(3) = MacroLst(I).Data(J).V1
    DefaultParams(5) = MacroLst(I).Data(J).V2
    If TabStrip1.Tabs.Count = 3 Then TabStrip1.Tabs.Remove 3
    Set TabStrip1.SelectedItem = TabStrip1.Tabs(2)
    cmdEdit.Enabled = False
    GoTo LoadPic:
  End If

  If TabStrip1.Tabs.Count = 2 Then
    TabStrip1.Tabs.Add 3, "C", Lang.GetString(RES_Macro_Parameters)
    Set TabStrip1.SelectedItem = TabStrip1.Tabs(2)
  End If

  If MacroFile <> "" And FileExists(MacroFile) Then
    cmdEdit.Enabled = True
  Else
    cmdEdit.Enabled = False
    picMacro.Picture = LoadPicture()
    txtInfo.Text = ""
    Exit Sub
  End If

  If Right$(MacroFile, 4) = ".scm" Then
    ParseASDMacro MacroFile, bmpFile
    bmpTitle = MakeFileNameNeat(MacroFile)
    If bmpFile = "" Then
      bmpFile = AddDir(GetDir(MacroFile), bmpTitle) & ".bmp"
    Else
      bmpFile = AddDir(GetDir(MacroFile), bmpFile)
    End If
  Else
    Meta = SearchMeta(MacroFile)
    
    ParseMacroParam Meta
    ParseMacroDesc Meta
    ParseMacroDefaults Meta

    bmpTitle = MakeFileNameNeat(MacroFile)
    bmpFile = AddDir(GetDir(MacroFile), bmpTitle) & ".bmp"
  End If

  If Not FileExists(bmpFile) Then _
    bmpFile = MultiDir(bmpTitle & ".bmp", Options.MacroPicPath)
  If Not FileExists(bmpFile) Then _
    bmpFile = AddDir(GetDir(MacroFile), bmpTitle) & ".gif"
  If Not FileExists(bmpFile) Then _
    bmpFile = MultiDir(bmpTitle & ".gif", Options.MacroPicPath)

LoadPic:
  If FileExists(bmpFile) Then
    picMacro.Picture = LoadPicture(bmpFile)
  Else
    picMacro.Picture = LoadPicture()
    picMacro.Cls
    CenterText picMacro, Lang.GetString(RES_ERR_MacroPic1), 36, picMacro.ScaleHeight / 2 - picMacro.TextHeight("H")
    CenterText picMacro, Lang.GetString(RES_ERR_MacroPic2), 36, picMacro.ScaleHeight / 2
    DrawExclaimIcon picMacro.hdc, 2, picMacro.ScaleHeight / 2 - 16
  End If
  If SetDefaults Or Options.UseMacroDefaults Then cmdDefaults_Click
End Sub

Private Sub txtMacro_GotFocus()
  SelectText txtMacro
End Sub

Private Sub txtMacro_KeyPress(KeyAscii As Integer)
  Set lstMacros.SelectedItem = Nothing
End Sub

Private Sub Txts_Change(Index As Integer)
  Dim TempStr As String
  If Index = 0 Then
    TempStr = Txts(0).Text
    If TempStr = "" Then
      Caption = Txts(0).Tag
    Else
      Caption = TempStr
    End If
  End If
  mTxtChanged = True
End Sub

Private Sub Txts_GotFocus(Index As Integer)
  If Index = 1 Or Index = 2 Or Index >= 5 Then
    SmartSelectText Txts(Index)
  Else
    If Index = 0 Then ReturnSymbol 0, 0, 2
    SelectText Txts(Index)
  End If
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
    Msg As String, _
    Value As Single, TempLatLon As clsLatLon

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
  
  If Cancel Then
    mTxtChanged = True
    SmartSelectText Txts(Index)
  End If
End Sub
