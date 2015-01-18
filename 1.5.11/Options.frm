VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   Icon            =   "Options.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   Tag             =   "2300"
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   6240
   End
   Begin VB.CommandButton cmdDefaults 
      Height          =   375
      Left            =   5040
      TabIndex        =   47
      Tag             =   "2480"
      Top             =   6240
      WhatsThisHelpID =   2480
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   46
      Tag             =   "1031"
      Top             =   6240
      WhatsThisHelpID =   1031
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   45
      Tag             =   "1030"
      Top             =   6240
      WhatsThisHelpID =   1030
      Width           =   1215
   End
   Begin VB.Frame TabFrame 
      BorderStyle     =   0  'None
      Height          =   5415
      Index           =   6
      Left            =   360
      TabIndex        =   37
      Top             =   600
      Width           =   5895
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   1
         Left            =   2520
         TabIndex        =   42
         Tag             =   "2442"
         Top             =   630
         WhatsThisHelpID =   2442
         Width           =   1095
      End
      Begin VB.CheckBox chks 
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   41
         Tag             =   "2442"
         Top             =   660
         WhatsThisHelpID =   2442
         Width           =   2535
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   2
         Left            =   2520
         TabIndex        =   44
         Top             =   1020
         WhatsThisHelpID =   2443
         Width           =   3015
      End
      Begin VB.CheckBox chks 
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   38
         Tag             =   "2440"
         Top             =   270
         WhatsThisHelpID =   2440
         Width           =   2450
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   0
         Left            =   2520
         TabIndex        =   39
         Tag             =   "2440"
         Top             =   240
         WhatsThisHelpID =   2440
         Width           =   1095
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   14
         Left            =   0
         TabIndex        =   43
         Tag             =   "2443"
         Top             =   1050
         WhatsThisHelpID =   2443
         Width           =   2445
      End
      Begin VB.Label lbls 
         Height          =   195
         Index           =   13
         Left            =   3720
         TabIndex        =   40
         Tag             =   "2441"
         Top             =   270
         WhatsThisHelpID =   2440
         Width           =   1065
      End
   End
   Begin VB.Frame TabFrame 
      BorderStyle     =   0  'None
      Caption         =   "Folders"
      Height          =   5415
      Index           =   5
      Left            =   360
      TabIndex        =   27
      Top             =   600
      Visible         =   0   'False
      Width           =   5895
      Begin VB.TextBox txtsTool 
         Height          =   285
         Index           =   1
         Left            =   2880
         TabIndex        =   36
         Top             =   1320
         WhatsThisHelpID =   2438
         Width           =   2895
      End
      Begin VB.CommandButton cmdTBrowse 
         Height          =   375
         Left            =   4680
         TabIndex        =   34
         Tag             =   "1032"
         Top             =   600
         WhatsThisHelpID =   1032
         Width           =   1095
      End
      Begin VB.TextBox txtsTool 
         Height          =   285
         Index           =   0
         Left            =   2880
         TabIndex        =   33
         Top             =   240
         WhatsThisHelpID =   2437
         Width           =   2895
      End
      Begin VB.CommandButton cmdTDelete 
         Height          =   375
         Left            =   1440
         TabIndex        =   31
         Tag             =   "2483"
         Top             =   4920
         WhatsThisHelpID =   2483
         Width           =   975
      End
      Begin VB.CommandButton cmdTNew 
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Tag             =   "2482"
         Top             =   4920
         WhatsThisHelpID =   2482
         Width           =   975
      End
      Begin VB.ListBox lstTools 
         Height          =   4545
         Left            =   120
         TabIndex        =   29
         Top             =   240
         WhatsThisHelpID =   2436
         Width           =   2295
      End
      Begin VB.Label lbls 
         Height          =   195
         Index           =   12
         Left            =   2760
         TabIndex        =   35
         Tag             =   "2438"
         Top             =   1080
         WhatsThisHelpID =   2438
         Width           =   1335
      End
      Begin VB.Label lbls 
         Height          =   195
         Index           =   11
         Left            =   2760
         TabIndex        =   32
         Tag             =   "2437"
         Top             =   0
         WhatsThisHelpID =   2437
         Width           =   1335
      End
      Begin VB.Label lbls 
         Height          =   195
         Index           =   10
         Left            =   0
         TabIndex        =   28
         Tag             =   "2436"
         Top             =   0
         WhatsThisHelpID =   2436
         Width           =   2415
      End
   End
   Begin VB.Frame TabFrame 
      BorderStyle     =   0  'None
      Caption         =   "Folders"
      Height          =   5415
      Index           =   4
      Left            =   360
      TabIndex        =   15
      Top             =   600
      Visible         =   0   'False
      Width           =   5895
      Begin VB.ComboBox cmbCategory 
         Height          =   315
         Left            =   2880
         TabIndex        =   26
         Top             =   2160
         WhatsThisHelpID =   2439
         Width           =   2895
      End
      Begin VB.TextBox txtsMacro 
         Height          =   285
         Index           =   1
         Left            =   2880
         TabIndex        =   24
         Top             =   1320
         WhatsThisHelpID =   2438
         Width           =   2895
      End
      Begin VB.CommandButton cmdMBrowse 
         Height          =   375
         Left            =   4680
         TabIndex        =   22
         Tag             =   "1032"
         Top             =   600
         WhatsThisHelpID =   1032
         Width           =   1095
      End
      Begin VB.TextBox txtsMacro 
         Height          =   285
         Index           =   0
         Left            =   2880
         TabIndex        =   21
         Top             =   240
         WhatsThisHelpID =   2437
         Width           =   2895
      End
      Begin VB.CommandButton cmdMDelete 
         Height          =   375
         Left            =   1440
         TabIndex        =   19
         Tag             =   "2483"
         Top             =   4920
         WhatsThisHelpID =   2483
         Width           =   975
      End
      Begin VB.CommandButton cmdMNew 
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Tag             =   "2482"
         Top             =   4920
         WhatsThisHelpID =   2482
         Width           =   975
      End
      Begin VB.ListBox lstMacros 
         Height          =   4545
         Left            =   120
         TabIndex        =   17
         Top             =   240
         WhatsThisHelpID =   2435
         Width           =   2295
      End
      Begin VB.Label lbls 
         Height          =   195
         Index           =   9
         Left            =   2760
         TabIndex        =   25
         Tag             =   "2439"
         Top             =   1920
         WhatsThisHelpID =   2439
         Width           =   1335
      End
      Begin VB.Label lbls 
         Height          =   195
         Index           =   8
         Left            =   2760
         TabIndex        =   23
         Tag             =   "2438"
         Top             =   1080
         WhatsThisHelpID =   2438
         Width           =   1335
      End
      Begin VB.Label lbls 
         Height          =   195
         Index           =   7
         Left            =   2760
         TabIndex        =   20
         Tag             =   "2437"
         Top             =   0
         WhatsThisHelpID =   2437
         Width           =   1335
      End
      Begin VB.Label lbls 
         Height          =   195
         Index           =   6
         Left            =   0
         TabIndex        =   16
         Tag             =   "2435"
         Top             =   0
         WhatsThisHelpID =   2435
         Width           =   2415
      End
   End
   Begin VB.Frame TabFrame 
      BorderStyle     =   0  'None
      Caption         =   "Folders"
      Height          =   5415
      Index           =   3
      Left            =   360
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   5895
      Begin VB.CommandButton cmdChange 
         Height          =   375
         Left            =   4440
         TabIndex        =   14
         Tag             =   "2481"
         Top             =   1500
         WhatsThisHelpID =   2481
         Width           =   1215
      End
      Begin VB.ComboBox cmbScheme 
         Height          =   315
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   240
         WhatsThisHelpID =   2433
         Width           =   2415
      End
      Begin VB.ListBox lstObjs 
         Height          =   4935
         Left            =   0
         TabIndex        =   9
         Top             =   240
         WhatsThisHelpID =   2432
         Width           =   2895
      End
      Begin VB.Label lbls 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   4
         Left            =   3960
         TabIndex        =   13
         Top             =   1560
         WhatsThisHelpID =   2434
         Width           =   300
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   3
         Left            =   3120
         TabIndex        =   12
         Tag             =   "2434"
         Top             =   1560
         WhatsThisHelpID =   2434
         Width           =   735
      End
      Begin VB.Label lbls 
         Height          =   195
         Index           =   5
         Left            =   3120
         TabIndex        =   10
         Tag             =   "2433"
         Top             =   0
         WhatsThisHelpID =   2433
         Width           =   1335
      End
      Begin VB.Label lbls 
         Height          =   195
         Index           =   2
         Left            =   0
         TabIndex        =   8
         Tag             =   "2432"
         Top             =   0
         WhatsThisHelpID =   2432
         Width           =   2415
      End
   End
   Begin VB.Frame TabFrame 
      BorderStyle     =   0  'None
      Caption         =   "Folders"
      Height          =   5415
      Index           =   2
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   5895
      Begin VB.CommandButton cmdBrowse 
         Height          =   375
         Left            =   4200
         TabIndex        =   6
         Tag             =   "2481"
         Top             =   4920
         WhatsThisHelpID =   2481
         Width           =   1575
      End
      Begin ComctlLib.ListView lstFolders 
         Height          =   4575
         Left            =   0
         TabIndex        =   5
         Top             =   240
         WhatsThisHelpID =   2431
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   8070
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   "2420"
            Text            =   ""
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   "2421"
            Text            =   ""
            Object.Width           =   6200
         EndProperty
      End
      Begin VB.Label lbls 
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   4
         Tag             =   "2431"
         Top             =   0
         WhatsThisHelpID =   2431
         Width           =   3015
      End
   End
   Begin VB.Frame TabFrame 
      BorderStyle     =   0  'None
      Caption         =   "General"
      Height          =   5415
      Index           =   1
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   5895
      Begin ComctlLib.TreeView tvwSettings 
         Height          =   5055
         Left            =   0
         TabIndex        =   2
         Top             =   240
         WhatsThisHelpID =   2430
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   8916
         _Version        =   327682
         Indentation     =   441
         LabelEdit       =   1
         Style           =   1
         ImageList       =   "ImageList1"
         Appearance      =   1
      End
      Begin VB.Label lbls 
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Tag             =   "2430"
         Top             =   0
         WhatsThisHelpID =   2430
         Width           =   3015
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   6015
      Left            =   120
      TabIndex        =   48
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   10610
      MultiRow        =   -1  'True
      TabFixedWidth   =   2117
      TabFixedHeight  =   482
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   6
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "A"
            Object.Tag             =   "2301"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "B"
            Object.Tag             =   "2302"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "C"
            Object.Tag             =   "2303"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "D"
            Object.Tag             =   "2304"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "E"
            Object.Tag             =   "2305"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "F"
            Object.Tag             =   "2306"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   600
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Options.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Options.frx":011E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Options.frx":0230
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Options.frx":0342
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Options.frx":0454
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Options.frx":0566
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Options.frx":0678
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Type TypeFavorites
  File As String
  Name As String
  Category As String
End Type

Private mChanged As Boolean
Private mMouseDown As Boolean

Private Const CB_SETDROPPEDWIDTH = &H160

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Private mColors() As Long
Private mMacros() As TypeFavorites
Private mTools() As TypeFavorites

Private mValueCache(1) As Single

Private TabStripIndex As Integer

Private Const IconClose = 1
Private Const IconOpen = 2
Private Const IconCheckOn = 3
Private Const IconCheckOff = 4
Private Const IconCheckGray = 5
Private Const IconRadioOn = 6
Private Const IconRadioOff = 7

Private Function AddBaseToList(ByVal ID As Integer) As Node
  Set AddBaseToList = tvwSettings.Nodes.Add(, , "ID" & ID, Lang.GetString(ID))
End Function

Private Function AddToList(Parent As Node, ByVal PictureID As Integer, ByVal Low As Integer, ByVal HighOrCount As Integer) As Node
  Dim I As Integer
  Parent.Expanded = True
  Parent.Image = IconClose
  Parent.ExpandedImage = IconOpen
  If HighOrCount < Low Then
    For I = Low To Low + HighOrCount - 1
      Set AddToList = tvwSettings.Nodes.Add(Parent, tvwChild, "ID" & I, Lang.GetString(I), PictureID)
    Next I
  Else
    For I = Low To HighOrCount
      Set AddToList = tvwSettings.Nodes.Add(Parent, tvwChild, "ID" & I, Lang.GetString(I), PictureID)
    Next I
  End If
End Function

Private Function GetCheck(ByVal ID As Integer) As Boolean
  GetCheck = tvwSettings.Nodes("ID" & ID).Image = IconCheckOn
End Function

Private Function GetItem(ByVal ID As Integer) As String
  GetItem = lstFolders.ListItems("ID" & ID).Tag
End Function

Private Function GetRadio(ByVal ID As Integer) As Integer
  Dim I As Integer
  I = ID
  For I = ID To ID + 100
    If tvwSettings.Nodes("ID" & I).Image = IconRadioOn Then
      GetRadio = I - ID
      Exit For
    End If
  Next I
End Function

Private Sub LoadCustomMacroFile()
  Dim FileCounter As Integer, _
    X As Integer, I As Integer, _
    File As String, myStr As String, myFile As String
  
  ReDim mMacros(10)
  
  File = AddDir(App.Path, "Macros\custom.lst")

  If FileExists(File) Then
    X = FreeFile
    Open File For Input As #X
    LineInputEx X, myStr
    If myStr = "FS Scenery Creator Macro Hierarchy" Then
      Do Until EOF(X)
        LineInputEx X, myStr
        If StrComp(Left$(myStr, 9), "Versions=", vbTextCompare) = 0 Then
        
        ElseIf StrComp(Left$(myStr, 5), "Main=", vbTextCompare) = 0 Then
        
        ElseIf myStr <> "" Then
          FileCounter = FileCounter + 1
          If FileCounter > UBound(mMacros) Then ReDim Preserve mMacros(FileCounter * 2)

          With mMacros(FileCounter)
            myFile = ReadNext(myStr, ",")
            .Name = ReadLast(myFile, "\")
            .Category = myFile
            .File = ReadNext(myStr, ",")
          End With
        End If
      Loop
    End If
    Close #X
  End If
  ReDim Preserve mMacros(FileCounter)
  
  For I = 1 To FileCounter
    lstMacros.AddItem mMacros(I).Name
  Next I
End Sub

Private Sub SaveCustomMacroFile()
  Dim X As Integer, I As Integer, _
    File As String
  
  File = AddDir(App.Path, "Macros\custom.lst")

  X = FreeFile
  On Error Resume Next
  Open File For Output As #X
  If Err > 0 Then Exit Sub ' Probably no Macros directory
  On Error GoTo 0
  Print #X, "FS Scenery Creator Macro Hierarchy"
  For I = 1 To UBound(mMacros)
    With mMacros(I)
      Print #X, AddDir(.Category, .Name) & ", " & .File
    End With
  Next I
  Close #X
End Sub

Private Sub SetCheck(ByVal ID As Integer, ByVal Checked As Boolean)
  tvwSettings.Nodes("ID" & ID).Image = IIf(Checked, IconCheckOn, IconCheckOff)
End Sub

Private Sub SetItem(ByVal ID As Integer, File As String)
  With lstFolders.ListItems("ID" & ID)
    .Tag = File
    .SubItems(1) = CompactedPath(File, lstFolders.ColumnHeaders(2).Width / Screen.TwipsPerPixelX, hdc)
  End With
End Sub

Private Sub SetRadio(ByVal ID As Integer, ByVal Value As Integer, Optional ByVal ClearOld As Boolean = False)
  Dim Value2 As Integer
  If ClearOld Then
    Value2 = GetRadio(ID)
    tvwSettings.Nodes("ID" & (ID + Value2)).Image = IconRadioOff
  End If
  tvwSettings.Nodes("ID" & (ID + Value)).Image = IconRadioOn
End Sub

' Edit Options data
Public Function EditData(Data As clsOptions) As Boolean
  Dim myPath As String, I As Integer

  Load frmOptions
  With Data
    SetCheck RES_OPT_RememberWindowState, .RememberWindowState
    SetCheck RES_OPT_NeatRecentFiles, .NeatRecentFiles
    SetCheck RES_OPT_ShowHeaderProperties, .ShowHeaderProperties
    SetCheck RES_OPT_Remember, .Remember
    SetCheck RES_OPT_ShowFractionalMinutes, .ShowFractionalMinutes
    SetRadio RES_OPT_UnitOfMeasure1, -(Not .Metric)
    SetRadio RES_OPT_Orientation1, -.Magnetic
    SetCheck RES_OPT_OldStyleMenus, .OldStyleMenus
    SetCheck RES_OPT_SaveCompressed, .SaveCompressed
    SetCheck RES_OPT_ShowExportWizard, .ShowExportWizard
    SetCheck RES_OPT_UseMacroDefaults, .UseMacroDefaults
    
    SetRadio RES_OPT_FSVersion1, .FSVersion
    
    SetCheck RES_OPT_EditConfig, .EditConfig
    SetCheck RES_OPT_AutoCompress, .AutoCompress
    SetCheck RES_OPT_KeepSourceFile, .KeepSourceFile
    SetCheck RES_OPT_SaveBeforeCompile, .SaveBeforeCompile
    
    SetRadio RES_OPT_Crosshair1, -.CrossHair
    SetCheck RES_OPT_FocusCircles, .FocusCircle
    SetCheck RES_OPT_PointCircles, .PointCircle
    SetCheck RES_OPT_SnapPoints, .SnapPoints
    SetCheck RES_OPT_FillPolygons, .FillPolygons
    SetCheck RES_OPT_ThickLines, .ThickLines
    SetCheck RES_OPT_FillObjects, .FillObjects
    SetCheck RES_OPT_ShowCompass, .ShowCompass
  
    For I = RES_Obj_Runway To RES_Obj_Point
      SetCheck I, .ObjectVisible(I - RES_Obj_Header)
    Next I
    
    SetItem RES_OPT_FSFolder, .FSPath
    SetItem RES_OPT_TexFolder, .TexturePath
    SetItem RES_OPT_Compiler, .Compiler
    SetItem RES_OPT_Compress, .CompressName
    SetItem RES_OPT_TextEditor, .TextEditor
    
    myPath = .MacroPath
    For I = 1 To 5
      SetItem I, ReadNext(myPath, ";")
    Next I

    myPath = .MacroPicPath
    For I = 21 To 25
      SetItem I, ReadNext(myPath, ";")
    Next I
    
    cmbScheme.ListIndex = .ColorScheme
    If .ColorScheme = -1 Then .GetColors mColors
    
    LoadCustomMacroFile
    ReDim mTools(.ToolCount)
    For I = 1 To UBound(mTools)
      mTools(I).File = .ToolExe(I)
      mTools(I).Name = .ToolName(I)
      lstTools.AddItem mTools(I).Name
    Next I
    
    If lstMacros.ListCount > 0 Then lstMacros.ListIndex = 0
    If lstTools.ListCount > 0 Then lstTools.ListIndex = 0
    
    chks(0).Value = -(.AutoSave > 0)
    chks_Click 0
    Txts(0).Text = .AutoSave
    
    chks(1).Value = -(.Grid > 0)
    chks_Click 1
    Txts(1).Text = MeterToUser(.Grid)

    Txts(2).Text = .TextureFilter
        
    If TabStripIndex = 0 Then TabStripIndex = 1
    Set TabStrip1.SelectedItem = TabStrip1.Tabs(TabStripIndex)
    
    mChanged = False
    Show vbModal, frmMain
    
    TabStripIndex = TabStrip1.SelectedItem.Index
    
    If mChanged Then
      .RememberWindowState = GetCheck(RES_OPT_RememberWindowState)
      .NeatRecentFiles = GetCheck(RES_OPT_NeatRecentFiles)
      .ShowHeaderProperties = GetCheck(RES_OPT_ShowHeaderProperties)
      .Remember = GetCheck(RES_OPT_Remember)
      .ShowFractionalMinutes = GetCheck(RES_OPT_ShowFractionalMinutes)
      .Metric = (GetRadio(RES_OPT_UnitOfMeasure1) = 0)
      .Magnetic = (GetRadio(RES_OPT_Orientation1) = 1)
      .OldStyleMenus = GetCheck(RES_OPT_OldStyleMenus)
      .SaveCompressed = GetCheck(RES_OPT_SaveCompressed)
      .ShowExportWizard = GetCheck(RES_OPT_ShowExportWizard)
      .UseMacroDefaults = GetCheck(RES_OPT_UseMacroDefaults)
      
      .FSVersion = GetRadio(RES_OPT_FSVersion1)
      
      .EditConfig = GetCheck(RES_OPT_EditConfig)
      .AutoCompress = GetCheck(RES_OPT_AutoCompress)
      .KeepSourceFile = GetCheck(RES_OPT_KeepSourceFile)
      .SaveBeforeCompile = GetCheck(RES_OPT_SaveBeforeCompile)
      
      .CrossHair = -GetRadio(RES_OPT_Crosshair1)
      .FocusCircle = GetCheck(RES_OPT_FocusCircles)
      .PointCircle = GetCheck(RES_OPT_PointCircles)
      .SnapPoints = GetCheck(RES_OPT_SnapPoints)
      .FillPolygons = GetCheck(RES_OPT_FillPolygons)
      .ThickLines = GetCheck(RES_OPT_ThickLines)
      .FillObjects = GetCheck(RES_OPT_FillObjects)
      .ShowCompass = GetCheck(RES_OPT_ShowCompass)
    
      For I = RES_Obj_Runway To RES_Obj_Point
        .ObjectVisible(I - RES_Obj_Header) = GetCheck(I)
      Next I
      
      .FSPath = GetItem(RES_OPT_FSFolder)
      .TexturePath = GetItem(RES_OPT_TexFolder)
      .Compiler = GetItem(RES_OPT_Compiler)
      .CompressName = GetItem(RES_OPT_Compress)
      .TextEditor = GetItem(RES_OPT_TextEditor)
      
      myPath = ""
      For I = 1 To 5
        myPath = myPath & GetItem(I) & ";"
      Next I
      .MacroPath = myPath
  
      myPath = ""
      For I = 21 To 25
        myPath = myPath & GetItem(I) & ";"
      Next I
      .MacroPicPath = myPath
      
      .ColorScheme = cmbScheme.ListIndex
      
      For I = 0 To MAX_OBJ + 6
        .ObjectColor(I) = mColors(I)
      Next I
      .GetGLColors
      
      SaveCustomMacroFile
      
      .ToolCount = UBound(mTools)
      
      For I = 1 To UBound(mTools)
        .ToolExe(I) = mTools(I).File
        .ToolName(I) = mTools(I).Name
      Next I
      
      .AutoSave = chks(0).Value * mValueCache(0)
      .Grid = chks(1).Value * mValueCache(1)
  
      .TextureFilter = Txts(2).Text

      EditData = True
    End If
  End With
  Unload frmOptions
End Function

Private Sub chks_Click(Index As Integer)
  Select Case Index
    Case 0
      SetEnabled Txts(0), -chks(0).Value
      If -chks(0).Value And Val(Txts(0).Text) = 0 Then Txts(0).Text = "15"
    Case 1
      SetEnabled Txts(1), -chks(1).Value
      If -chks(1).Value And Val(Txts(1).Text) = 0 Then Txts(1).Text = MeterToUser(1000)
  End Select
End Sub

Private Sub cmbCategory_Change()
  Dim X As Integer
  X = lstMacros.ListIndex + 1
  If X = 0 Then Exit Sub
  mMacros(X).Category = cmbCategory.Text
End Sub

Private Sub cmbCategory_Click()
  cmbCategory_Change
End Sub

Private Sub cmbCategory_KeyPress(KeyAscii As Integer)
  If KeyAscii = 44 Then ' Comma
    KeyAscii = 0
  End If
End Sub

Private Sub cmbScheme_Click()
  Select Case cmbScheme.ListIndex
    Case 0
      Options.GetDefaultColors mColors
    Case 1
      Options.GetDefaultAirportColors mColors
  End Select
  lstObjs_Click
End Sub

Private Sub cmdBrowse_Click()
  Dim myPath As String, Key As String, Index As Integer
  Key = lstFolders.SelectedItem.Key
  Index = Val(Mid$(Key, 3))
  myPath = GetItem(Index)
  
  Select Case Index
    Case RES_OPT_FSFolder, RES_OPT_TexFolder
      frmPath.EditData myPath, Lang.GetString(Index + 10), True
      SetItem Index, myPath
    Case RES_OPT_Compiler, RES_OPT_Compress, RES_OPT_TextEditor
      With cDialog
        .Filter = Lang.GetString(RES_OPT_ExecutableFilter)
        .FilterIndex = 1
        .DefExt = "exe"
        myPath = .OpenDialog(Lang.GetString(Index + 10), myPath)
        If myPath <> "" Then SetItem Index, myPath
      End With
    Case 1, 2, 3, 4, 5
      frmPath.EditData myPath, Lang.GetString(RES_OPT_MacroFolder + 10), True
      SetItem Index, GetRealName(myPath)
    Case 21, 22, 23, 24, 25
      frmPath.EditData myPath, Lang.GetString(RES_OPT_MacroPicFolder + 10), True
      SetItem Index, GetRealName(myPath)
  End Select
End Sub

Private Sub cmdCancel_Click()
  Hide
End Sub

Private Sub cmdChange_Click()
  Dim X As Integer, Col As Long
  X = lstObjs.ListIndex + 1
  Col = mColors(lstObjs.ListIndex)
  If cDialog.ColorDialog(Col, True) Then
    mColors(X) = Col
    lbls(4).BackColor = Col
    cmbScheme.ListIndex = -1
  End If
End Sub

Private Sub cmdDefaults_Click()
  Dim I As Integer, TempStr As String, _
    FSPath As String, FSVersion As Integer, _
    Path As String, myPath As String

  Path = GetRealName(App.Path)
  
  FSVersion = GetRadio(RES_OPT_FSVersion1)
  
  FSPath = Options.GetFSPath(FSVersion)
  
  For I = Version_FS2K2 To Version_FS95 Step -1
    If FSPath <> "" Then
      Exit For
    Else
      FSPath = Options.GetFSPath(I)
      FSVersion = I
    End If
  Next I
  
  If FSPath = "" Then
    FSPath = "C:\"
  End If
  
  SetCheck RES_OPT_RememberWindowState, False
  SetCheck RES_OPT_NeatRecentFiles, True
  SetCheck RES_OPT_ShowHeaderProperties, False
  SetCheck RES_OPT_Remember, True
  SetCheck RES_OPT_ShowFractionalMinutes, True
  SetRadio RES_OPT_UnitOfMeasure1, 0, True
  SetRadio RES_OPT_Orientation1, 0, True
  SetCheck RES_OPT_OldStyleMenus, False
  SetCheck RES_OPT_SaveCompressed, True
  SetCheck RES_OPT_ShowExportWizard, False
  SetCheck RES_OPT_UseMacroDefaults, True
  
  SetRadio RES_OPT_FSVersion1, FSVersion, True
  
  SetCheck RES_OPT_EditConfig, True
  SetCheck RES_OPT_KeepSourceFile, False
  SetCheck RES_OPT_SaveBeforeCompile, True
  
  SetRadio RES_OPT_Crosshair1, 0, True
  SetCheck RES_OPT_FocusCircles, True
  SetCheck RES_OPT_PointCircles, True
  SetCheck RES_OPT_SnapPoints, True
  SetCheck RES_OPT_FillPolygons, True
  SetCheck RES_OPT_ThickLines, True
  SetCheck RES_OPT_FillObjects, True
  SetCheck RES_OPT_ShowCompass, False

  For I = RES_Obj_Runway To RES_Obj_Point
    SetCheck I, True
  Next I
    
  SetItem RES_OPT_FSFolder, FSPath
  SetItem RES_OPT_TexFolder, AddDir(FSPath, "Texture")
  SetItem RES_OPT_Compiler, AddDir(Path, "SCASM\SCASM.EXE")
  TempStr = GetRealName(AddDir(Path, "bglzip.exe"))
  SetItem RES_OPT_Compress, TempStr
  SetCheck RES_OPT_AutoCompress, FileExists(TempStr)
  SetItem RES_OPT_TextEditor, GetRealName(AddDir(GetWindowsDir(), "NOTEPAD.EXE"))
    
  For I = 1 To 5
    SetItem I, ""
  Next I

  For I = 21 To 25
    SetItem I, ""
  Next I
    
  SetItem 1, AddDir(Path, "Macros")
  SetItem 21, AddDir(Path, "Macros\Bitmaps")

  TempStr = RegGetKey("PGMDirectory", "None", "Software\Airport 2.xx")
  If TempStr <> "None" Then
    On Error Resume Next
    SetItem 2, GetRealName(AddDir(TempStr, "API"))
    SetItem 22, GetRealName(AddDir(TempStr, "Resource"))
    On Error GoTo 0
  End If

  TempStr = RegGetKey("UserMACRODirectory", "None", "Software\Airport 2.xx")
  If TempStr <> "None" Then
    On Error Resume Next
    SetItem 3, GetRealName(TempStr)
    On Error GoTo 0
  End If

  cmbScheme.ListIndex = 0
  lstObjs.ListIndex = 0
  
  lstTools.Clear
  ReDim mTools(3)

  I = 0
  TempStr = RegGetKey("Program", "", "Software\MW\DXTBmp")
  If TempStr <> "" Then
    I = I + 1
    mTools(I).Name = "DXTBmp"
    mTools(I).File = TempStr
    lstTools.AddItem mTools(I).Name
  End If
  
  TempStr = RegGetKey("AppPath", "", "Software\VB and VBA Program Settings\Brueckner\EOD")
  If TempStr <> "" Then
    I = I + 1
    mTools(I).Name = "Easy Object Designer"
    mTools(I).File = AddDir(TempStr, "eod.exe")
    lstTools.AddItem mTools(I).Name
  End If
  ReDim Preserve mTools(I)

  If lstMacros.ListCount > 0 Then lstMacros.ListIndex = 0
  If lstTools.ListCount > 0 Then lstTools.ListIndex = 0
  
  chks(0).Value = vbChecked
  chks_Click 0
  Txts(0).Text = 15

  chks(1).Value = vbChecked
  chks_Click 1
  Txts(1).Text = MeterToUser(1000)

  Txts(2).Text = "*.bmp;*.r8;*.txr;*.oav;*.?af;*.pat"
End Sub

Private Sub cmdMBrowse_Click()
  Dim myPath As String
  myPath = txtsMacro(0).Text
  With cDialog
    .Filter = Lang.GetString(RES_MacroFilter)
    .FilterIndex = 1
    .DefExt = "api"
    myPath = .OpenDialog(Lang.GetString(RES_OPT_Macro), myPath)
    If myPath <> "" Then
      If InStr(myPath, ",") > 0 Then myPath = GetShortName(myPath)
      txtsMacro(0).Text = myPath
      txtsMacro(1).Text = MakeFileNameNeat(myPath)
    End If
  End With
End Sub

Private Sub cmdMDelete_Click()
  Dim X As Integer, I As Integer
  X = lstMacros.ListIndex
  lstMacros.RemoveItem X
  lstMacros.ListIndex = -1
  
  For I = X + 1 To UBound(mMacros) - 1
    mMacros(I) = mMacros(I + 1)
  Next I
  ReDim Preserve mMacros(UBound(mMacros) - 1)
  lstMacros_Click
End Sub

Private Sub cmdMNew_Click()
  Dim myPath As String, X As Integer
  With cDialog
    .Filter = Lang.GetString(RES_MacroFilter)
    .FilterIndex = 1
    .DefExt = "api"
    myPath = .OpenDialog(Lang.GetString(RES_OPT_Macro))
    If myPath = "" Then Exit Sub
    If InStr(myPath, ",") > 0 Then myPath = GetShortName(myPath)
  End With
  
  X = UBound(mMacros) + 1
  ReDim Preserve mMacros(X)
  With mMacros(X)
    .File = myPath
    .Name = MakeFileNameNeat(myPath)
    .Category = Lang.GetString(RES_OPT_Favorites)
    lstMacros.AddItem .Name
  End With
  lstMacros.ListIndex = X - 1
End Sub

Private Sub cmdOK_Click()
  Dim I As Integer, Msg As String
  
  For I = 0 To 1
    If Not Txts(I).Locked Then _
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

Private Sub cmdTBrowse_Click()
  Dim myPath As String
  myPath = txtsTool(0).Text
  With cDialog
    .Filter = Lang.GetString(RES_OPT_ExecutableFilter)
    .FilterIndex = 1
    .DefExt = "api"
    myPath = .OpenDialog(Lang.GetString(RES_OPT_Tool), myPath)
    If myPath <> "" Then
      txtsTool(0).Text = myPath
      txtsTool(1).Text = MakeFileNameNeat(myPath)
    End If
  End With
End Sub

Private Sub cmdTDelete_Click()
  Dim X As Integer, I As Integer
  X = lstTools.ListIndex
  lstTools.RemoveItem X
  lstTools.ListIndex = -1
  
  For I = X + 1 To UBound(mTools) - 1
    mTools(I) = mTools(I + 1)
  Next I
  ReDim Preserve mTools(UBound(mTools) - 1)
  lstTools_Click
End Sub

Private Sub cmdTNew_Click()
  Dim myPath As String, X As Integer
  With cDialog
    .Filter = Lang.GetString(RES_OPT_ExecutableFilter)
    .FilterIndex = 1
    .DefExt = "exe"
    myPath = .OpenDialog(Lang.GetString(RES_OPT_Tool))
    If myPath = "" Then Exit Sub
  End With
  
  X = UBound(mTools) + 1
  ReDim Preserve mTools(X)
  With mTools(X)
    .File = myPath
    .Name = MakeFileNameNeat(myPath)
    lstTools.AddItem .Name
  End With
  lstTools.ListIndex = X - 1
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
  Dim I As Integer, mItem(1) As Node, _
    TempStr As String, TempStr2 As String, _
    TempWidth As Integer, MaxWidth As Integer, _
    HeightDiff As Single

  DialogMenus Me
  Lang.PrepareForm Me
  
  HeightDiff = TabStrip1.ClientTop - 450
  
  For I = 1 To TabFrame.Count
    With TabFrame(I)
      .Move TabFrame(1).Left, TabStrip1.ClientTop + 150, TabFrame(1).Width, TabFrame(1).Height
      .Enabled = False
      .Visible = False
    End With
  Next I

  cmdOK.Top = cmdOK.Top + HeightDiff
  cmdCancel.Top = cmdCancel.Top + HeightDiff
  cmdDefaults.Top = cmdDefaults.Top + HeightDiff
  TabStrip1.Height = TabStrip1.Height + HeightDiff
  Height = Height + HeightDiff
  
  CenterForm Me

  Set mItem(0) = AddBaseToList(RES_OPT_GeneralMain)
  Set mItem(1) = AddToList(mItem(0), IconCheckOff, RES_OPT_RememberWindowState, RES_OPT_UnitOfMeasureMain)
  AddToList mItem(1), IconRadioOff, RES_OPT_UnitOfMeasure1, 2
  Set mItem(1) = AddToList(mItem(0), IconCheckOff, RES_OPT_OrientationMain, 1)
  AddToList mItem(1), IconRadioOff, RES_OPT_Orientation1, 2
  Set mItem(1) = AddToList(mItem(0), IconCheckOff, RES_OPT_OldStyleMenus, RES_OPT_ShowExportWizard)
  Set mItem(1) = AddToList(mItem(0), IconCheckOff, RES_OPT_UseMacroDefaults, RES_OPT_UseMacroDefaults)
  
  Set mItem(0) = AddBaseToList(RES_OPT_FSVersionMain)
  AddToList mItem(0), IconRadioOff, RES_OPT_FSVersion1, 6
  
  Set mItem(0) = AddBaseToList(RES_OPT_ExportMain)
  AddToList mItem(0), IconCheckOff, RES_OPT_EditConfig, RES_OPT_SaveBeforeCompile
  
  Set mItem(0) = AddBaseToList(RES_OPT_AppearanceMain)
  Set mItem(1) = AddToList(mItem(0), IconCheckOff, RES_OPT_CrosshairMain, 1)
  AddToList mItem(1), IconRadioOff, RES_OPT_Crosshair1, 2
  AddToList mItem(0), IconCheckOff, RES_OPT_FocusCircles, RES_OPT_ShowCompass

  Set mItem(0) = AddBaseToList(RES_OPT_VisibleMain)
  With mItem(0)
    .Expanded = True
    .Image = IconClose
    .ExpandedImage = IconOpen
  End With
  
  For I = RES_Obj_Runway To RES_Obj_Point
    tvwSettings.Nodes.Add mItem(0), tvwChild, "ID" & I, Lang.GetString(I), IconCheckOff
  Next I

  Lang.AddItems lstObjs, RES_Obj_Header + 1, (MAX_OBJ + 6)
  ReDim mColors(MAX_OBJ + 6)
  lstObjs.ListIndex = 0

  Lang.AddItems cmbScheme, RES_OPT_Scheme1, 2
     
  With lstFolders.ListItems
    Set lstFolders.SelectedItem = .Add(, "ID" & RES_OPT_FSFolder, Lang.GetString(RES_OPT_FSFolder))
    .Add , "ID" & RES_OPT_TexFolder, Lang.GetString(RES_OPT_TexFolder)
    .Add , "ID" & RES_OPT_Compiler, Lang.GetString(RES_OPT_Compiler)
    .Add , "ID" & RES_OPT_Compress, Lang.GetString(RES_OPT_Compress)
    .Add , "ID" & RES_OPT_TextEditor, Lang.GetString(RES_OPT_TextEditor)
    For I = 1 To 5
      .Add , "ID" & I, Lang.ResolveString(RES_OPT_MacroFolder, I)
    Next I
    For I = 1 To 5
      .Add , "ID" & I + 20, Lang.ResolveString(RES_OPT_MacroPicFolder, I)
    Next I
  End With
  
  TempStr = Lang.GetString(RES_OPT_Favorites)
  cmbCategory.AddItem TempStr
  For I = 1 To UBound(MacroLst)
    If Mid$(MacroLst(I).DirName, 2, 1) <> ":" Then
      TempStr2 = Replace(MacroLst(I).DirName, "\>", "\")
      If TempStr <> TempStr2 Then
        cmbCategory.AddItem TempStr2
        TempWidth = TextWidth(TempStr2)
        If MaxWidth < TempWidth Then MaxWidth = TempWidth
      End If
    End If
  Next I
  
  MaxWidth = MaxWidth + 700
  
  If MaxWidth > cmbCategory.Width Then
    SendMessageLong cmbCategory.hwnd, CB_SETDROPPEDWIDTH, MaxWidth / Screen.TwipsPerPixelX, 0
  End If
  
  lstMacros_Click
  lstTools_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then Cancel = True: Hide
End Sub

Private Sub lstFolders_DblClick()
  cmdBrowse_Click
End Sub

Private Sub lstMacros_Click()
  Dim Value As Boolean
  Value = (lstMacros.ListIndex > -1) And (lstMacros.ListCount > 0)
  SetEnabled txtsMacro(0), Value
  SetEnabled txtsMacro(1), Value
  cmdMBrowse.Enabled = Value
  cmdMDelete.Enabled = Value
  cmbCategory.BackColor = IIf(Value, vbWindowBackground, vbButtonFace)
  cmbCategory.Locked = Not Value
    
  If Value Then
    With mMacros(lstMacros.ListIndex + 1)
      txtsMacro(0).Text = .File
      txtsMacro(1).Text = .Name
      SetComboText cmbCategory, .Category
    End With
  Else
    txtsMacro(0).Text = ""
    txtsMacro(1).Text = ""
    cmbCategory.Text = ""
  End If
End Sub

Private Sub lstObjs_Click()
  lbls(4).BackColor = mColors(lstObjs.ListIndex + 1)
End Sub

Private Sub lstTools_Click()
  Dim Value As Boolean
  Value = (lstTools.ListIndex > -1) And (lstTools.ListCount > 0)
  SetEnabled txtsTool(0), Value
  SetEnabled txtsTool(1), Value
  cmdTBrowse.Enabled = Value
  cmdTDelete.Enabled = Value
    
  If Value Then
    With mTools(lstTools.ListIndex + 1)
      txtsTool(0).Text = .File
      txtsTool(1).Text = .Name
    End With
  Else
    txtsTool(0).Text = ""
    txtsTool(1).Text = ""
  End If
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
          Case 1: tvwSettings.SetFocus
          Case 2: lstFolders.SetFocus
          Case 3: lstObjs.SetFocus
          Case 4: lstMacros.SetFocus
          Case 5: lstTools.SetFocus
          Case 6: chks(0).SetFocus
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

Private Sub tvwSettings_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    KeyCode = 0
    mMouseDown = True
    tvwSettings_NodeClick tvwSettings.SelectedItem
  Else
    mMouseDown = False
  End If
End Sub

Private Sub tvwSettings_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeySpace Then KeyAscii = 0
End Sub

Private Sub tvwSettings_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mMouseDown = True
End Sub

Private Sub tvwSettings_NodeClick(ByVal Node As ComctlLib.Node)
  Dim mItem As Node, TempStr As String, _
    FSPath As String, DoNewPath As Boolean
  
  If mMouseDown Then
    LockWindowUpdate tvwSettings.hwnd
    Select Case Node.Image
      Case IconCheckOff
        Node.Image = IconCheckOn
      Case IconCheckOn
        Node.Image = IconCheckOff
      Case IconRadioOn
        ' Nothing
      Case IconRadioOff
        Set mItem = Node.FirstSibling
        If mItem.Key = "ID" & RES_OPT_FSVersion1 Then DoNewPath = True
        
        Do Until mItem Is Nothing
          mItem.Image = IconRadioOff
          Set mItem = mItem.Next
        Loop
        Node.Image = IconRadioOn
        
        If DoNewPath Then
          FSPath = Options.GetFSPath(GetRadio(RES_OPT_FSVersion1))
          
          If FSPath <> "" Then
            SetItem RES_OPT_FSFolder, FSPath
            SetItem RES_OPT_TexFolder, AddDir(FSPath, "Texture")
            TempStr = GetRealName(AddDir(FSPath, "bglzip.exe"))
            SetItem RES_OPT_Compress, TempStr
            SetCheck RES_OPT_AutoCompress, FileExists(TempStr)
          End If
        End If
    End Select
    LockWindowUpdate 0
  End If
End Sub

Private Sub Txts_GotFocus(Index As Integer)
  SelectText Txts(Index)
End Sub

Private Sub txtsMacro_Change(Index As Integer)
  Dim X As Integer
  X = lstMacros.ListIndex + 1
  If X = 0 Then Exit Sub
  Select Case Index
    Case 0
      mMacros(X).File = txtsMacro(Index).Text
    Case 1
      mMacros(X).Name = txtsMacro(Index).Text
      lstMacros.List(X - 1) = mMacros(X).Name
  End Select
End Sub

Private Sub txtsMacro_GotFocus(Index As Integer)
  SelectText txtsMacro(Index)
End Sub

Private Sub txtsMacro_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 44 Then ' Comma
    KeyAscii = 0
  End If
  
  If Index = 1 And KeyAscii = 92 Then ' Backslash
    KeyAscii = 0
  End If
End Sub

Private Sub txtsTool_Change(Index As Integer)
  Dim X As Integer
  X = lstTools.ListIndex + 1
  If X = 0 Then Exit Sub
  Select Case Index
    Case 0
      mTools(X).File = txtsTool(Index).Text
    Case 1
      mTools(X).Name = txtsTool(Index).Text
      lstTools.List(X - 1) = mTools(X).Name
  End Select
End Sub

Private Sub txtsTool_GotFocus(Index As Integer)
  SelectText txtsTool(Index)
End Sub
