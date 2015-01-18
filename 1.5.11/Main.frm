VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Flight Simulator Scenery Creator"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   735
   ClientWidth     =   7980
   Icon            =   "Main.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   356
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   532
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   7080
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3960
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Timer TimerMouseMove 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7080
      Top             =   3480
   End
   Begin VB.Timer TimerAutoSave 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   7080
      Top             =   3000
   End
   Begin VB.PictureBox picSymbol 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   1
      Left            =   7440
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picSymbol 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   0
      Left            =   7080
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox Splitter 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   2880
      MousePointer    =   9  'Size W E
      ScaleHeight     =   113
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox picList 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      DragMode        =   1  'Automatic
      Height          =   1695
      Left            =   3480
      ScaleHeight     =   113
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   193
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   2895
      Begin ComctlLib.TreeView lstObjects 
         Height          =   1335
         Left            =   0
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   360
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   2355
         _Version        =   327682
         HideSelection   =   0   'False
         Indentation     =   397
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "imgTreeView"
         Appearance      =   1
      End
      Begin VB.PictureBox picButton 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   8.25
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   2400
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   60
         Width           =   240
      End
      Begin VB.PictureBox picButton 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   8.25
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   75
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   60
         Width           =   240
      End
   End
   Begin VB.PictureBox picEdit 
      BackColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   120
      ScaleHeight     =   109
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   165
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   2535
      Begin VB.HScrollBar HScroll 
         Height          =   255
         LargeChange     =   1000
         Left            =   120
         Max             =   100
         Min             =   -100
         SmallChange     =   100
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1815
      End
      Begin VB.VScrollBar VScroll 
         Height          =   975
         LargeChange     =   1000
         Left            =   2040
         Max             =   100
         Min             =   -100
         SmallChange     =   100
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.PictureBox picProgress 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   7980
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4830
      Visible         =   0   'False
      Width           =   7980
      Begin ComctlLib.ProgressBar barProgress 
         Height          =   195
         Left            =   2640
         TabIndex        =   5
         Top             =   30
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   344
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label lblStatus 
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   30
         Width           =   2490
      End
   End
   Begin ComctlLib.StatusBar Statusbar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   5085
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   7937
            MinWidth        =   7937
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   6615
            MinWidth        =   6615
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   9260
            MinWidth        =   9260
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picToolbar 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   0
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   532
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7980
      Begin ComctlLib.Toolbar Toolbar1 
         Height          =   390
         Left            =   150
         TabIndex        =   1
         Top             =   45
         Width           =   7530
         _ExtentX        =   13282
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         ImageList       =   "imgIcons"
         _Version        =   327682
         BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
            NumButtons      =   26
            BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Visible         =   0   'False
               Object.Tag             =   ""
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               ImageIndex      =   6
            EndProperty
            BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               ImageIndex      =   7
            EndProperty
            BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               ImageIndex      =   8
            EndProperty
            BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               ImageIndex      =   9
            EndProperty
            BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               ImageIndex      =   10
            EndProperty
            BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               ImageIndex      =   11
               Style           =   1
            EndProperty
            BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               ImageIndex      =   12
               Style           =   1
            EndProperty
            BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               ImageIndex      =   13
            EndProperty
            BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               ImageIndex      =   14
            EndProperty
            BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               ImageIndex      =   15
            EndProperty
            BeginProperty Button20 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               ImageIndex      =   16
            EndProperty
            BeginProperty Button21 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button22 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               ImageIndex      =   17
            EndProperty
            BeginProperty Button23 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               ImageIndex      =   18
            EndProperty
            BeginProperty Button24 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               ImageIndex      =   19
            EndProperty
            BeginProperty Button25 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button26 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               ImageIndex      =   20
            EndProperty
         EndProperty
      End
   End
   Begin ComctlLib.ImageList imgListButtons 
      Left            =   7080
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   18
      ImageHeight     =   18
      MaskColor       =   65280
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":099C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imgTreeView 
      Left            =   7080
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":0AAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":0BC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":0CD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":0DE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":0EF6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imgIcons 
      Left            =   7080
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   21
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":1008
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":111A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":122C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":133E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":1450
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":1562
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":1674
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":1786
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":1898
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":19AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":1BCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":1CE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":1DF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":1F04
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":2016
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":2128
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":223A
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":234C
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":245E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":2570
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&File"
      Index           =   0
      Tag             =   "3500"
      Begin VB.Menu mnuNew 
         Caption         =   ""
         Shortcut        =   ^N
         Tag             =   "3501"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   ""
         Shortcut        =   ^O
         Tag             =   "3502"
      End
      Begin VB.Menu mnuSave 
         Caption         =   ""
         Shortcut        =   ^S
         Tag             =   "3503"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   ""
         Tag             =   "3504"
      End
      Begin VB.Menu mnuFileDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImport 
         Caption         =   ""
         Tag             =   "3505"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExport 
         Caption         =   ""
         Tag             =   "3506"
      End
      Begin VB.Menu mnuExportWizard 
         Caption         =   ""
         Tag             =   "3507"
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   ""
         Index           =   1
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   ""
         Index           =   2
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   ""
         Index           =   3
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   ""
         Index           =   4
      End
      Begin VB.Menu mnuFileDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   ""
         Tag             =   "3508"
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Edit"
      Index           =   1
      Tag             =   "3520"
      Begin VB.Menu mnuUndo 
         Caption         =   ""
         Shortcut        =   ^Z
         Tag             =   "3521"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditDash1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCut 
         Caption         =   ""
         Shortcut        =   ^X
         Tag             =   "3522"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   ""
         Shortcut        =   ^C
         Tag             =   "3523"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   ""
         Shortcut        =   ^V
         Tag             =   "3524"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   ""
         Shortcut        =   {DEL}
         Tag             =   "3525"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   ""
         Shortcut        =   ^A
         Tag             =   "3526"
      End
      Begin VB.Menu EditDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSortObjects 
         Caption         =   ""
         Tag             =   "3527"
      End
      Begin VB.Menu mnuTransform 
         Caption         =   ""
         Tag             =   "3528"
      End
      Begin VB.Menu mnuSceneryProperties 
         Caption         =   ""
         Tag             =   "3529"
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Tools"
      Index           =   2
      Tag             =   "3540"
      Begin VB.Menu mnuPredefinedTools 
         Caption         =   ""
         Index           =   0
         Tag             =   "3541"
      End
      Begin VB.Menu mnuPredefinedTools 
         Caption         =   ""
         Index           =   1
         Tag             =   "3542"
      End
      Begin VB.Menu mnuPrograms 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuToolDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPreferences 
         Caption         =   ""
         Tag             =   "3547"
      End
      Begin VB.Menu mnuLanguage 
         Caption         =   ""
         Tag             =   "3548"
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&View"
      Index           =   3
      Tag             =   "3560"
      Begin VB.Menu mnuObjects 
         Caption         =   ""
         Tag             =   "3561"
      End
      Begin VB.Menu mnuViewDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuZoom 
         Caption         =   ""
         Index           =   0
         Tag             =   "3562"
      End
      Begin VB.Menu mnuZoom 
         Caption         =   ""
         Index           =   1
         Tag             =   "3563"
      End
      Begin VB.Menu mnuZoomSpecify 
         Caption         =   ""
         Tag             =   "3564"
      End
      Begin VB.Menu mnuStandard 
         Caption         =   ""
         Tag             =   "3565"
      End
      Begin VB.Menu mnuViewDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOrder 
         Caption         =   ""
         Index           =   0
         Tag             =   "3566"
      End
      Begin VB.Menu mnuOrder 
         Caption         =   ""
         Index           =   1
         Tag             =   "3567"
      End
      Begin VB.Menu mnuViewDash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolbar 
         Caption         =   ""
         Tag             =   "3568"
      End
      Begin VB.Menu mnuStatusbar 
         Caption         =   ""
         Tag             =   "3569"
      End
      Begin VB.Menu mnuScrollbars 
         Caption         =   ""
         Tag             =   "3570"
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Help"
      Index           =   4
      Tag             =   "3580"
      Begin VB.Menu mnuHelp 
         Caption         =   ""
         Index           =   0
         Shortcut        =   {F1}
         Tag             =   "3581"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   ""
         Index           =   1
         Tag             =   "3582"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   ""
         Index           =   2
         Tag             =   "3583"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   ""
         Index           =   3
         Tag             =   "3584"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   ""
         Index           =   4
         Tag             =   "3585"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   ""
         Index           =   5
         Tag             =   "3586"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelpDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContact 
         Caption         =   ""
         Tag             =   "3587"
      End
      Begin VB.Menu mnuWeb 
         Caption         =   ""
         Tag             =   "3588"
      End
      Begin VB.Menu mnuHelpDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   ""
         Tag             =   "3590"
      End
      Begin VB.Menu mnuAcknowledgements 
         Caption         =   ""
         Tag             =   "3591"
      End
      Begin VB.Menu mnuPopup 
         Caption         =   ""
         Begin VB.Menu mnuCenterHere 
            Caption         =   ""
            Tag             =   "3600"
         End
         Begin VB.Menu mnuOrderPopup 
            Caption         =   ""
            Tag             =   "3601"
            Begin VB.Menu mnuPopOrder 
               Caption         =   ""
               Index           =   0
               Tag             =   "3566"
            End
            Begin VB.Menu mnuPopOrder 
               Caption         =   ""
               Index           =   1
               Tag             =   "3567"
            End
         End
         Begin VB.Menu mnuPopDash1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuInsert 
            Caption         =   ""
            Index           =   0
            Tag             =   "3602"
            Begin VB.Menu mnu2D 
               Caption         =   ""
               Index           =   0
               Tag             =   "-3201"
            End
            Begin VB.Menu mnu2D 
               Caption         =   ""
               Index           =   1
               Tag             =   "-3202"
            End
            Begin VB.Menu mnu2D 
               Caption         =   ""
               Index           =   2
               Tag             =   "-3203"
            End
            Begin VB.Menu mnu2D 
               Caption         =   ""
               Index           =   3
               Tag             =   "-3204"
            End
            Begin VB.Menu mnu2D 
               Caption         =   ""
               Index           =   4
               Tag             =   "-3205"
            End
            Begin VB.Menu mnu2D 
               Caption         =   ""
               Index           =   5
               Tag             =   "-3206"
            End
            Begin VB.Menu mnu2D 
               Caption         =   ""
               Index           =   6
               Tag             =   "-3207"
            End
         End
         Begin VB.Menu mnuInsert 
            Caption         =   ""
            Index           =   1
            Tag             =   "3603"
            Begin VB.Menu mnu3D 
               Caption         =   ""
               Index           =   0
               Tag             =   "-3208"
            End
            Begin VB.Menu mnu3D 
               Caption         =   ""
               Index           =   1
               Tag             =   "-3209"
            End
         End
         Begin VB.Menu mnuInsert 
            Caption         =   ""
            Index           =   2
            Tag             =   "3604"
            Begin VB.Menu mnuRadio 
               Caption         =   ""
               Index           =   0
               Tag             =   "-3210"
            End
            Begin VB.Menu mnuRadio 
               Caption         =   ""
               Index           =   1
               Tag             =   "-3211"
            End
            Begin VB.Menu mnuRadio 
               Caption         =   ""
               Index           =   2
               Tag             =   "-3212"
            End
            Begin VB.Menu mnuRadio 
               Caption         =   ""
               Index           =   3
               Tag             =   "-3213"
            End
         End
         Begin VB.Menu mnuInsert 
            Caption         =   ""
            Index           =   3
            Tag             =   "3605"
            Begin VB.Menu mnuMisc 
               Caption         =   ""
               Index           =   0
               Tag             =   "-3214"
            End
            Begin VB.Menu mnuMisc 
               Caption         =   ""
               Index           =   1
               Tag             =   "-3215"
            End
            Begin VB.Menu mnuMisc 
               Caption         =   ""
               Index           =   2
               Tag             =   "-3216"
            End
            Begin VB.Menu mnuMisc 
               Caption         =   ""
               Index           =   3
               Tag             =   "-3217"
            End
            Begin VB.Menu mnuMisc 
               Caption         =   ""
               Index           =   4
               Tag             =   "-3218"
            End
            Begin VB.Menu mnuMisc 
               Caption         =   ""
               Index           =   5
               Tag             =   "-3219"
            End
         End
         Begin VB.Menu mnuAdd 
            Caption         =   ""
            Index           =   0
            Tag             =   "-3210"
         End
         Begin VB.Menu mnuAdd 
            Caption         =   ""
            Index           =   1
            Tag             =   "-3211"
         End
         Begin VB.Menu mnuAdd 
            Caption         =   ""
            Index           =   2
            Tag             =   "-3212"
         End
         Begin VB.Menu mnuAdd 
            Caption         =   ""
            Index           =   3
            Tag             =   "-3213"
         End
         Begin VB.Menu mnuAdd 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnuAdd 
            Caption         =   ""
            Index           =   5
            Tag             =   "-3201"
         End
         Begin VB.Menu mnuAdd 
            Caption         =   ""
            Index           =   6
            Tag             =   "-3214"
         End
         Begin VB.Menu mnuAdd 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu mnuAdd 
            Caption         =   ""
            Index           =   8
            Tag             =   "-3202"
         End
         Begin VB.Menu mnuAdd 
            Caption         =   ""
            Index           =   9
            Tag             =   "-3203"
         End
         Begin VB.Menu mnuAdd 
            Caption         =   ""
            Index           =   10
            Tag             =   "-3204"
         End
         Begin VB.Menu mnuAdd 
            Caption         =   ""
            Index           =   11
            Tag             =   "-3205"
         End
         Begin VB.Menu mnuAdd 
            Caption         =   ""
            Index           =   12
            Tag             =   "-3206"
         End
         Begin VB.Menu mnuAdd 
            Caption         =   ""
            Index           =   13
            Tag             =   "-3207"
         End
         Begin VB.Menu mnuAdd 
            Caption         =   "-"
            Index           =   14
         End
         Begin VB.Menu mnuAdd 
            Caption         =   ""
            Index           =   15
            Tag             =   "-3208"
         End
         Begin VB.Menu mnuAdd 
            Caption         =   ""
            Index           =   16
            Tag             =   "-3209"
         End
         Begin VB.Menu mnuAdd 
            Caption         =   "-"
            Index           =   17
         End
         Begin VB.Menu mnuAdd 
            Caption         =   ""
            Index           =   18
            Tag             =   "-3216"
         End
         Begin VB.Menu mnuAdd 
            Caption         =   ""
            Index           =   19
            Tag             =   "-3217"
         End
         Begin VB.Menu mnuAdd 
            Caption         =   ""
            Index           =   20
            Tag             =   "-3215"
         End
         Begin VB.Menu mnuAdd 
            Caption         =   ""
            Index           =   21
            Tag             =   "-3218"
         End
         Begin VB.Menu mnuAdd 
            Caption         =   ""
            Index           =   22
            Tag             =   "-3219"
         End
         Begin VB.Menu mnuPopDash2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPointAdd 
            Caption         =   ""
            Tag             =   "3607"
         End
         Begin VB.Menu mnuPopDash3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuMainProperties 
            Caption         =   ""
            Shortcut        =   {F4}
         End
         Begin VB.Menu mnuProperties 
            Caption         =   "-"
            Index           =   0
         End
      End
      Begin VB.Menu mnuATISPopup 
         Caption         =   ""
         Begin VB.Menu mnuATIS 
            Caption         =   ""
            Index           =   0
            Tag             =   "1665"
         End
         Begin VB.Menu mnuATIS 
            Caption         =   ""
            Index           =   1
            Tag             =   "1666"
         End
         Begin VB.Menu mnuATIS 
            Caption         =   ""
            Index           =   2
            Tag             =   "1667"
         End
         Begin VB.Menu mnuATIS 
            Caption         =   ""
            Index           =   3
            Tag             =   "1669"
         End
         Begin VB.Menu mnuATIS 
            Caption         =   ""
            Index           =   4
            Tag             =   "1670"
         End
         Begin VB.Menu mnuATIS 
            Caption         =   ""
            Index           =   5
            Tag             =   "1671"
         End
         Begin VB.Menu mnuATIS 
            Caption         =   ""
            Index           =   6
            Tag             =   "1672"
         End
         Begin VB.Menu mnuATIS 
            Caption         =   ""
            Index           =   7
            Tag             =   "1673"
         End
         Begin VB.Menu mnuATIS 
            Caption         =   ""
            Index           =   8
            Tag             =   "1674"
         End
         Begin VB.Menu mnuATIS 
            Caption         =   ""
            Index           =   9
            Tag             =   "1675"
         End
         Begin VB.Menu mnuATIS 
            Caption         =   ""
            Index           =   10
            Tag             =   "1676"
         End
         Begin VB.Menu mnuATIS 
            Caption         =   ""
            Index           =   11
            Tag             =   "1677"
         End
         Begin VB.Menu mnuATIS 
            Caption         =   ""
            Index           =   12
            Tag             =   "1680"
         End
         Begin VB.Menu mnuATIS 
            Caption         =   ""
            Index           =   13
            Tag             =   "1681"
         End
         Begin VB.Menu mnuATIS 
            Caption         =   ""
            Index           =   14
            Tag             =   "1682"
         End
         Begin VB.Menu mnuATIS 
            Caption         =   ""
            Index           =   15
            Tag             =   "1683"
         End
         Begin VB.Menu mnuATIS 
            Caption         =   ""
            Index           =   16
            Tag             =   "1684"
         End
         Begin VB.Menu mnuATIS 
            Caption         =   ""
            Index           =   17
            Tag             =   "1685"
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
#If SUBCLASS Then
Implements ISubclass
#End If

' Popup Menu functions
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hwnd As Long, lprc As Any) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Const TPM_LEFTALIGN = &H0&
Private Const TPM_RIGHTBUTTON = &H2&

' Faster than Move (Hopefully!)
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

' Toolbar
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hwndParent As Long, ByVal hWndChildWindow As Long, ByVal lpClassName As String, ByVal lpsWindowName As String) As Long
Private Const TBSTYLE_FLAT = &H800
Private Const WM_USER = &H400&
Private Const TB_GETSTYLE = (WM_USER + 57)
Private Const TB_SETSTYLE = (WM_USER + 56)

' Return world coordinates of mouse pointer
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

' Process message to get change in decimal operator
Private Const WM_WININICHANGE = &H1A

' Mousewheel event
Private Const WM_MOUSEWHEEL = &H20A
Private Const MK_SHIFT = &H4
Private Const MK_CONTROL = &H8

' Scrolling
Private Const WM_HSCROLL = &H114
Private Const WM_VSCROLL = &H115
Private Const SB_LINELEFT = 0
Private Const SB_LINERIGHT = 1
Private Const SB_LINEDOWN = 1
Private Const SB_LINEUP = 0

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

' Grid properties
Private mMapCenter As PointType
Private mMetersPerPixel As Single

' Used for new object, drag, grid
Private mCursorXY As PointType   ' Previous position
Private mRedrawOverride As Boolean

' Keeps track of which objects need to be redrawn in a drag
Private mFocusIndices() As Integer
Private mRedrawIndices() As Integer
Private mShapeRedraw() As clsShape

' Clipboard Cache
Private mClipboard() As clsObject

' Keeps track of dragging/gridding status
Private Dragging As Boolean, Gridding As Boolean

' So we can catch events from the global analogues
Private WithEvents mScenery As clsScenery, _
        WithEvents mOptions As clsOptions, _
        WithEvents mLang As clsLanguage
Attribute mScenery.VB_VarHelpID = -1
Attribute mOptions.VB_VarHelpID = -1
Attribute mLang.VB_VarHelpID = -1
        
Private mCDialog As clsCDialog

' Flags
Private mDontUpdate As Boolean
Private mSplitterPos As Long
Private mFormActivateCalled As Boolean
Private mInMouseAction As Boolean

' Cache of original editor
Private picDrag As clsOpenGL

' Extra bells and whistles
#If SUBCLASS Then
Private mIconMenus As cIconMenu
#End If
'Private mInfoTip As cTooltip

' Add an object to the scenery
Private Sub AddSceneryObject(ByVal MenuTag As Integer)
  Scenery.Add -MenuTag - 3200, mCursorXY.X, mCursorXY.Y
End Sub

' Function AskIfSave
' Asks the user if he/she wants to save the current file
' Returns: True - User saved it or doesn't want to save
'          False - User aborts action
Private Function AskIfSave() As Boolean
  If Scenery.Changed Then
    Select Case MsgBoxEx(Me, Lang.ResolveString(RES_Main_DoSave, MakeFileNameNeat(Scenery.File)), vbExclamation Or vbYesNoCancel, RES_Main_DoSave)
      Case vbYes
        AskIfSave = SaveAsProc(False)
      Case vbNo
        AskIfSave = True
      Case vbCancel
        AskIfSave = False
    End Select
  Else
    AskIfSave = True
  End If
End Function

Private Sub EraseClipboard()
  Dim I As Integer
  For I = 0 To UBound(mClipboard)
    Set mClipboard(I) = Nothing
  Next I
  ReDim mClipboard(0)
End Sub

' Caches FS data into arrays
Private Sub GetFSData()
  Dim I As Integer, J As Integer, DataStr() As Byte, _
    Temp As String, FileNum As Integer
  
  DataStr = LoadResData("DATA3", 10)
  For I = 0 To 49
    FSColors(I) = RGB(DataStr(I * 3), DataStr(I * 3 + 1), DataStr(I * 3 + 2))
  Next I
  
  FileNum = FreeFile
  Open AddDir(App.Path, "Synth.dat") For Input As #FileNum
  For I = 0 To 26
    LineInputEx FileNum, Temp
    With SynNames(I)
      .ID = ReadNext(Temp, " ")
      .File = Temp
    End With
  Next I
  Close #FileNum

  Open AddDir(App.Path, "Bldg.dat") For Input As #FileNum
  For I = 0 To 7
    LineInputEx FileNum, Temp
    With Building(I)
      .File = "side" & CStr(I + 1) & ".r8"
      .Color = Val(ReadNext(Temp, " "))
      .Windows = Val(ReadNext(Temp, " "))
      .Roof = Val(ReadNext(Temp, " "))
    End With
  Next I
  Building(0).File = "side.r8"
  Close #FileNum

  Open AddDir(App.Path, "AdvBldg1.dat") For Input As #FileNum
  For I = 8 To 85
    LineInputEx FileNum, Building1_3(I)
  Next I
  Close #FileNum
  
  Open AddDir(App.Path, "AdvBldg2.dat") For Input As #FileNum
  For I = 4 To 84
    LineInputEx FileNum, Building2(I)
  Next I
  Close #FileNum

  Open AddDir(App.Path, "AdvBldgR.dat") For Input As #FileNum
  For I = 4 To 33
    LineInputEx FileNum, BuildingR(I)
  Next I
  Close #FileNum
  
  Open AddDir(App.Path, "Regions.dat") For Input As #FileNum
  For I = 0 To 6
    LineInputEx FileNum, Temp
    Regions(I).ID = Mid$(Temp, 2, Len(Temp) - 2)
    For J = 0 To 8
      LineInputEx FileNum, Regions(I).Regions(J)
    Next J
    LineInputEx FileNum, Temp
  Next I
  Close #FileNum
End Sub

' Loads the scenery elements into the list
Private Sub LoadObjects()
  Dim I As Integer, J As Integer, X As Long, _
      Base As Node, mItem(MAX_OBJ - 1) As Node, Obj As Node, _
      ShapeObject As clsShape, PointObject As clsPoint

  SetScreenMousePointer vbHourglass

  ' Prevent flickering
  LockWindowUpdate lstObjects.hwnd
  lstObjects.Nodes.Clear

  Set Base = lstObjects.Nodes.Add(, , "Top", Scenery.Header.Name, 4)
  Base.Expanded = True

  ' Add the top level objects/folders first
  Set mItem(OT_Header) = lstObjects.Nodes.Add(Base, tvwChild, "Obj0", Lang.GetString(RES_Obj_Header), 3)
  For I = 1 To MAX_OBJ - 1
    Set mItem(I) = lstObjects.Nodes.Add(Base, tvwChild, , Lang.GetString(RES_Obj_Header + I), 1, 2)
  Next I

  ' Add the objects
  For I = 1 To Scenery.Count
    If TypeOf Scenery(I) Is clsPoint Then
      If Scenery(I).ObjectIndex = 1 Then
        Set PointObject = Scenery(I)
        Set ShapeObject = PointObject.Parent
        ' For the shapes, we will add another level to include
        ' the points
        Set Obj = lstObjects.Nodes.Add(mItem(ShapeObject.ShapeType), tvwChild, "Nod" & CStr(I), ShapeObject.Caption, 1, 2)
        For J = 0 To ShapeObject.NumPoints
          lstObjects.Nodes.Add Obj, tvwChild, "Obj" & CStr(ShapeObject.Point(J).SceneryIndex), Lang.ResolveString(Res_Obj_Point3, J + 1), 3
        Next J
        Set Obj = Nothing
        Set ShapeObject = Nothing
        Set PointObject = Nothing
      End If
    Else
      Dim ObjectType As Long
      ObjectType = Scenery(I).ObjectType
      ' somehow, using the above line directly below doesn't decrement the reference count, so we store it in a variable before use
      lstObjects.Nodes.Add mItem(ObjectType), tvwChild, "Obj" & CStr(I), Scenery(I).Caption, 3
    End If
  Next I
  Set Base = Nothing
  For I = 0 To MAX_OBJ - 1
    Set mItem(I) = Nothing
  Next I
  LockWindowUpdate 0
  SetScreenMousePointer vbDefault
End Sub

' Go through the specified menu array and resolving
' text strings
Private Sub ProcessMenus(X As Object, ByVal UboundX As Integer)
  Dim I As Integer
  For I = 0 To UboundX
    If Val(X(I).Tag) < 0 Then X(I).Caption = Lang.ResolveString(RES_Menu_ShortcutKey, Lang.GetString(-X(I).Tag))
  Next I
End Sub

' Make zoom buttons enabled or disabled
Private Sub ReAdjustZoomButtons()
  mnuZoom(0).Enabled = (mMetersPerPixel > 0.01)
  mnuZoom(1).Enabled = (mMetersPerPixel < 64)
  Toolbar1.Buttons(22).Enabled = mnuZoom(0).Enabled
  Toolbar1.Buttons(23).Enabled = mnuZoom(1).Enabled
  picEdit_Resize
End Sub

' Does the "Save" or "Save As" routine depending on
' filename
Private Function SaveAsProc(ByVal Prompt As Boolean) As Boolean
  Dim X As String
  If Scenery.File = Lang.GetString(RES_UntitledFile) Or Prompt Or StrComp(Right$(Scenery.File, 4), ".apt", vbTextCompare) = 0 Then
    With cDialog
      .Filter = Lang.GetString(RES_SaveFilter)
      .FilterIndex = 1
      .DefExt = "scn"
      If StrComp(Right$(Scenery.File, 4), ".apt", vbTextCompare) = 0 Then
        X = .SaveDialog(Left$(Scenery.File, Len(Scenery.File) - 3) & "scn", Lang.GetString(RES_Main_SaveCaption))
      Else
        X = .SaveDialog(Scenery.File, Lang.GetString(RES_Main_SaveCaption))
      End If
    End With
    
    If X <> "" Then
      Scenery.SaveFile X
      Options.AddRecentFile Scenery.File
      SaveAsProc = True
    End If
  Else
    Scenery.SaveFile Scenery.File
    Options.AddRecentFile Scenery.File
    SaveAsProc = True
  End If
End Function

' Changes the scaling of picEditor to the current
' magnify factor and adjusts scaling factors
Private Sub SetScale()
  Dim myWidth As Integer, myHeight As Integer, _
    Left As Single, Top As Single, _
    Right As Single, Bottom As Single, _
    H1 As Integer, H2 As Integer, _
    V1 As Integer, V2 As Integer
  
  myWidth = picEdit.ScaleWidth - IIf(VScroll.Visible, VScroll.Width, 0)
  myHeight = picEdit.ScaleHeight - IIf(VScroll.Visible, HScroll.Height, 0)
  
  Left = -myWidth / 2 * mMetersPerPixel + mMapCenter.X
  Right = myWidth / 2 * mMetersPerPixel + mMapCenter.X
  Top = myHeight / 2 * mMetersPerPixel + mMapCenter.Y
  Bottom = -myHeight / 2 * mMetersPerPixel + mMapCenter.Y
  
  picEditor.SetScale Left, Top, Right, Bottom

  ' Change the min, max of scroll bars
  With Scenery.Header
    HScroll.Min = -.Horz / 20: HScroll.Max = .Horz / 20
    VScroll.Min = -.Vert / 20: VScroll.Max = .Vert / 20
  End With
  
  ' Change the LargeChange/SmallChange in the scroll bars
  H1 = CInt((Right - Left) / 50)
  H2 = H1 / 2
  V1 = CInt((Top - Bottom) / 50)
  V2 = V1 / 2
  If H1 < 1 Then H1 = 1
  If H2 < 1 Then H2 = 1
  If V1 < 1 Then V1 = 1
  If V2 < 1 Then V2 = 1
  HScroll.LargeChange = H1
  HScroll.SmallChange = H2
  VScroll.LargeChange = V1
  VScroll.SmallChange = V2
End Sub

' Sets the three "constant" test strings for numericality
Private Sub SetSymbols()
  Dim Signs As String
  Signs = LocaleInfo(LOCALE_SPOSITIVESIGN) & LocaleInfo(LOCALE_SNEGATIVESIGN)
  
  DigitsDecimal = Digits & "." & LocaleInfo(LOCALE_SDECIMAL)
  DigitsSigns = Digits & Signs
  DigitsDecimalSigns = DigitsDecimal & Signs
End Sub

Public Sub UpdateObjectList(Optional Index As Integer = -999)
  Dim X As String, PointObject As clsPoint, ShapeObject As clsShape
  If lstObjects.Visible Then
    If Index <> -999 Then
      If Index = 0 Then
        lstObjects.Nodes("Top").Text = Scenery(Index).Caption
        mScenery_TitleBarChange
      Else
        If TypeOf Scenery(Index) Is clsPoint Then
          Set PointObject = Scenery(Index)
          Set ShapeObject = PointObject.Parent
          ' Change the parent caption
          lstObjects.Nodes("Nod" & CStr(ShapeObject.Point(0).SceneryIndex)).Text = ShapeObject.Caption
          Set ShapeObject = Nothing
          Set PointObject = Nothing
        Else
          lstObjects.Nodes("Obj" & Index).Text = Scenery(Index).Caption
        End If
      End If
    Else
      On Error Resume Next
      X = lstObjects.SelectedItem.Key
      LoadObjects
      Set lstObjects.SelectedItem = lstObjects.Nodes(X)
      lstObjects.SelectedItem.EnsureVisible
    End If
  End If
End Sub

' Given the two diagonal grid corners, set the mapcenter and zoom factor
Private Sub ZoomToFit(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, ByVal PosX As Long, ByVal PosY As Long)
  Dim X As Single, Y As Single, ScaleValue1 As Single, ScaleValue2 As Single
  X = (X1 + X2) / 2
  Y = (Y1 + Y2) / 2
  
  If Not (Abs(X) > Scenery.Header.Horz / 2) And Not (Abs(Y) > Scenery.Header.Vert / 2) Then
    ScaleValue1 = Abs(X2 - X1) / picEditor.myWidth
    ScaleValue2 = Abs(Y2 - Y1) / picEditor.myHeight
    If ScaleValue2 > ScaleValue1 Then ScaleValue1 = ScaleValue2
    If ScaleValue1 < 0.01 Then ScaleValue1 = 0.01
    mMetersPerPixel = ScaleValue1
    mDontUpdate = True
    HScroll.Value = X / 10
    VScroll.Value = -Y / 10
    mMapCenter = MakePoint(X, Y)
    mDontUpdate = False
    picEdit_Resize
    picEdit_MouseMove 0, 0, CSng(PosX), CSng(PosY)
    ReAdjustZoomButtons
  Else
    Beep
  End If
End Sub

Private Sub Form_Activate()
  Dim hwndTT As Long, AutoSaveFile As String
  
  If Not mFormActivateCalled Then
    mFormActivateCalled = True
    
    ' Treeview Bug workaround
    ' Get the handle of the TreeView's tooltip window
    hwndTT = SendMessageLong(lstObjects.hwnd, TVM_GETTOOLTIPS, 0, 0)
    If hwndTT Then EnableWindow hwndTT, 1
  
    Statusbar.Visible = Options.StatusbarVisible
    picEdit.Visible = True
    Form_Resize
    Unload frmSplash

    ' Load a commandline file only after the form is shown
    ' so the statusbar is visible
    
    ' If another instance of FSSC is running, there may be autosave files in use
    If Not App.PrevInstance Then
      AutoSaveFile = frmAutosave.DoDialog()
    End If
    
    If AutoSaveFile <> "" Then
      Scenery.LoadFile GetRealName(AutoSaveFile)
      ' Always prompt to save autosave file
      Scenery.Changed = True
    ElseIf Command$ <> "" Then
      Scenery.LoadFile GetRealName(UnQuoteString(Command$))
      Options.AddRecentFile Scenery.File
    End If

    If Options.ShowTips Then mnuHelp_Click 4
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim I As Integer, _
    deltaX As Single, deltaY As Single
  Dim PointObject As clsPoint
  ' Handle "non-standard" shortcut keys
  Const vbKeyApps = &H5D
  
  If Dragging Or Gridding Then Exit Sub
  
  If (KeyCode = vbKeyApps And Shift = 0) Or (KeyCode = vbKeyF10 And Shift = vbShiftMask) Then
    ' Menu Button
    picEdit_MouseDown vbRightButton, 0, 0, 0
  ElseIf Between(KeyCode, vbKeyF4, vbKeyF8) Then
    If mnuMainProperties.Visible Then
      If Shift = 0 Then
        TabValue = KeyCode - vbKeyF4
        mnuMainProperties_Click
      ElseIf KeyCode = vbKeyF4 And Shift = vbShiftMask Then
        TabValue = 1
        mnuMainProperties_Click
      End If
      TabValue = 0
    End If
  ElseIf KeyCode = vbKeyInsert And Shift = 0 And mnuPointAdd.Visible Then
    mnuPointAdd_Click
  ElseIf KeyCode = vbKeyL And Shift = vbCtrlMask Then
    Toolbar1_ButtonClick Toolbar1.Buttons(16)
  ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
    Toolbar1_ButtonClick Toolbar1.Buttons(17)
  ElseIf (KeyCode = vbKeyK Or KeyCode = vbKeyU) And Shift = vbCtrlMask Then
    ' Lock
    For I = 1 To UBound(mFocusIndices)
      Scenery(mFocusIndices(I)).Locked = (KeyCode = vbKeyK)
    Next I
  ElseIf (Shift = vbCtrlMask) Then
    Select Case KeyCode
      Case vbKeyLeft: deltaX = -1
      Case vbKeyRight: deltaX = 1
      Case vbKeyUp: deltaY = 1
      Case vbKeyDown: deltaY = -1
    End Select
    If deltaX <> 0 Or deltaY <> 0 Then
      deltaX = deltaX * picEditor.ScaleX(1)
      deltaY = deltaY * picEditor.ScaleX(1)
      For I = 1 To UBound(mFocusIndices)
        With Scenery(mFocusIndices(I))
          If Not .Locked Then
            .PositionX = .PositionX + deltaX
            .PositionY = .PositionY + deltaY
            If .ObjectType = OT_Point Then
              Set PointObject = Scenery(mFocusIndices(I))
              PointObject.Parent.BoolTag = True
              Set PointObject = Nothing
            Else
              .UpdateObject
            End If
          End If
        End With
      Next I
      
      For I = 1 To UBound(mFocusIndices)
        If Scenery(mFocusIndices(I)).ObjectType = OT_Point Then
          Set PointObject = Scenery(mFocusIndices(I))
          If PointObject.Parent.BoolTag Then
            PointObject.Parent.Point(0).UpdateObject
            PointObject.Parent.BoolTag = False
          End If
          Set PointObject = Nothing
        End If
      Next I
      
      Scenery.Draw
    End If
  Else
    With Scenery
      If .Count > 0 Then
        If ((Shift = 0) And (KeyCode = vbKeyLeft Or KeyCode = vbKeyUp)) Or ((Shift = vbShiftMask) And (KeyCode = vbKeyTab)) Then
          If .SingleFocus > 1 Then
            .SetSingleFocus .SingleFocus - 1
          Else
            .SetSingleFocus .Count
          End If
        ElseIf (Shift = 0) And (KeyCode = vbKeyRight Or KeyCode = vbKeyDown Or KeyCode = vbKeyTab) Then
          If .SingleFocus < .Count Then
            .SetSingleFocus .SingleFocus + 1
          Else
            .SetSingleFocus 1
          End If
        End If
      End If
    End With
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  Dim lp As POINTAPI, Moved As Boolean, _
    NewX As Single, NewY As Single
  
  Dim DoSomething As Boolean
  
  ' Handle "non-standard" shortcut keys
  Select Case KeyAscii
    Case vbKeyEscape
      If mInMouseAction Then
        ' End drag or grid
        picEdit_MouseMove -2, 0, 0, 0
        picEdit_MouseUp 0, 0, 0, 0
        Scenery.Draw
        DoSomething = True
      End If
    Case 43, 61 'Asc("+"), Asc("=")
      If Gridding Then
        picEdit_MouseMove -2, 0, 0, 0
        GetCursorPos lp
        ScreenToClient picEdit.hwnd, lp
        picEditor.PixelToScale lp.X, lp.Y, NewX, NewY
        ZoomToFit mCursorXY.X, mCursorXY.Y, NewX, NewY, lp.X, lp.Y
        Gridding = False
        DoSomething = True
      Else
        If mnuZoom(0).Enabled Then
          mnuZoom_Click 0
          DoSomething = True
        End If
      End If
    Case 45 'Asc("-")
      If mnuZoom(1).Enabled Then
        If Gridding Then picEdit_MouseMove -2, 0, 0, 0
        mnuZoom_Click 1
        DoSomething = True
      End If
  End Select

  If DoSomething And (Gridding Or Dragging) Then
    If Dragging Then picEditor.CopyTo picDrag.hdc
    GetCursorPos lp
    ScreenToClient picEdit.hwnd, lp
    picEdit_MouseMove vbLeftButton, 0, (lp.X), (lp.Y)
  End If
End Sub

Private Sub Form_Load()
  Dim I As Integer, mhWndToolbarDll As Long
  
  ' Prevents form activate from being called
  ' for any odd reason
  mFormActivateCalled = True
  SetScreenMousePointer vbHourglass

  Set cDialog = New clsCDialog

  Set mOptions = Options
  Set mLang = Lang
  Set mCDialog = cDialog
  frmSplash.LoadPercent 40

  ' Zoom levels
  For I = 0 To 12
    ZoomLevels(I) = Choose(I + 1, 0.01, 0.02, 0.05, 0.1, 0.2, 0.5, 1, 2, 4, 8, 16, 32, 64)
  Next I

  ' CDialog
  cDialog.FilterIndex = 1
  
  ' Icon menus
  #If SUBCLASS Then
  Set mIconMenus = New cIconMenu
  With mIconMenus
    .Attach Me.hwnd
    .ImageList = imgIcons
    .SetIcon 2, 1 ' New
    .SetIcon 3, 2 ' Open
    .SetIcon 4, 3 ' Save
    .SetIcon 8, 4 ' Export
    .SetIcon 18, 5 ' Undo
    .SetIcon 20, 6 ' Cut
    .SetIcon 21, 7 ' Copy
    .SetIcon 22, 8 ' Paste
    .SetIcon 23, 9 ' Delete
    .SetIcon 37, 11 ' List
    .SetIcon 39, 17 ' Zoom in
    .SetIcon 40, 18 ' Zoom out
    .SetIcon 41, 19 ' Zoom
    .SetIcon 44, 15 ' Bring to front
    .SetIcon 45, 16 ' Send to back
    .SetIcon 51, 20 ' Help
    .SetIcon 64, 21 ' Center here
    .SetIcon 66, 15 ' Bring to front
    .SetIcon 67, 16 ' Send to back
  End With
  AttachMessage frmMain, Me.hwnd, WM_WININICHANGE
  AttachMessage frmMain, Me.hwnd, WM_MOUSEWHEEL
  #End If
  
  ' Big tool tip
  'Set mInfoTip = New cTooltip
  'mInfoTip.Create Me
  'mInfoTip.AddTool picEdit

  ' Form initialize
  mSplitterPos = &H7FFFFFFF    ' Splitter pos
  mnuPopup.Visible = False
  mnuATISPopup.Visible = False
  mnuPaste.Enabled = False
  Toolbar1.Buttons(5).Enabled = False ' Undo
  Toolbar1.Buttons(9).Enabled = False ' Paste
  picToolbar.Visible = Options.ToolbarVisible
  HScroll.Visible = Options.ScrollbarsVisible
  VScroll.Visible = Options.ScrollbarsVisible
  
  ' Prevents statusbar from appearing before
  ' everything is resized, prevents the ugly
  ' gray strip in the middle of the screen when
  ' still loading :-)
  Statusbar.Visible = False
  
  ' Make toolbars flat
  mhWndToolbarDll = FindWindowEx(Toolbar1.hwnd, 0, "ToolbarWindow32", "")
  SendMessageLong mhWndToolbarDll, TB_SETSTYLE, 0, SendMessageLong(mhWndToolbarDll, TB_GETSTYLE, 0, 0) Or TBSTYLE_FLAT
  
  ' Cache the button pictures
  DrawSymbolBox picSymbol(0), False
  DrawSymbolBox picSymbol(1), True

  ' Cache data
  GetFSData
  SetSymbols
  TempPathName = GetTempPathName()
  TextureHeader = LoadResData("DATA1", 10)

  frmSplash.LoadPercent 50
   
  ' OpenGL
  Set picEditor = New clsOpenGL
  Set picDrag = New clsOpenGL
  
  frmSplash.LoadPercent 75
  
  ' Scenery
  Set Scenery = New clsScenery
  Set mScenery = Scenery
  
  ReDim mClipboard(0)

  ' Initialization
  Scenery.NewFile

  frmSplash.LoadPercent 90
  
  ' These events were not run: the event objects were
  ' created after the LoadData was called in sub Main
  mOptions_Changed
  mLang_Changed
  
  ' Window position
  mDontUpdate = True
  With Options
    If .RememberWindowState And .WinMainState <> vbMaximized Then
      Move .WinMainLeft, .WinMainTop, .WinMainWidth, .WinMainHeight
    Else
      Move 0, 0, Screen.Width, Screen.Height * 0.85
      WindowState = vbMaximized
    End If
  
    If .ObjectsListWidth > 0 Then picList.Width = .ObjectsListWidth
  End With
  mDontUpdate = False
  mFormActivateCalled = False

  frmSplash.LoadPercent 100
  
  ' Enable the autosave timer
  TimerAutoSave.Enabled = True
  
  SetScreenMousePointer vbDefault
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Cancel = Not AskIfSave()
End Sub

Private Sub Form_Resize()
  Dim myHeight As Single, myPicListWidth As Single
  Const Offset = 4
  
  If Not mDontUpdate Then
    If WindowState <> vbMinimized Then
      On Error Resume Next
      SetScreenMousePointer vbHourglass

      myHeight = ScaleHeight - (picToolbar.Height * -Options.ToolbarVisible) - (Statusbar.Height * -Options.StatusbarVisible)
      mDontUpdate = True
      If picList.Visible Then
        If ScaleWidth - picList.Width < 100 Then
          mDontUpdate = False
          mnuObjects_Click
          SetScreenMousePointer vbDefault
          Exit Sub
        End If
        If picList.Left = 0 Then
          If mSplitterPos < &H7FFFFFFF Then
            myPicListWidth = mSplitterPos
          Else
            myPicListWidth = picList.Width
          End If
          MoveWindow picEdit.hwnd, myPicListWidth + 4, (picToolbar.Height * -Options.ToolbarVisible), ScaleWidth - myPicListWidth - 4, myHeight, 1
          MoveWindow Splitter.hwnd, myPicListWidth, (picToolbar.Height * -Options.ToolbarVisible), 4, myHeight, 1
          MoveWindow picList.hwnd, 0, picEdit.Top, myPicListWidth, myHeight, 1
        Else
          If mSplitterPos < &H7FFFFFFF Then
            myPicListWidth = ScaleWidth - mSplitterPos - Splitter.Width
          Else
            myPicListWidth = picList.Width
          End If
          MoveWindow picEdit.hwnd, 0, (picToolbar.Height * -Options.ToolbarVisible), ScaleWidth - myPicListWidth - 4, myHeight, 1
          MoveWindow Splitter.hwnd, picEdit.Width, (picToolbar.Height * -Options.ToolbarVisible), 4, myHeight, 1
          MoveWindow picList.hwnd, picEdit.Width + 4, picEdit.Top, myPicListWidth, myHeight, 1
        End If
        MoveWindow lstObjects.hwnd, 0, 22, myPicListWidth, picList.Height - 22, 1
      Else
        MoveWindow picEdit.hwnd, 0, (picToolbar.Height * -Options.ToolbarVisible), ScaleWidth, myHeight, 1
      End If
      MoveWindow VScroll.hwnd, picEdit.Width - VScroll.Width - Offset, 0, VScroll.Width, myHeight - HScroll.Height - Offset, 1
      MoveWindow HScroll.hwnd, 0, myHeight - HScroll.Height - Offset, picEdit.Width - VScroll.Width - Offset, HScroll.Height, 1
      mDontUpdate = False
      
      If picEdit.Visible Then picEdit_Resize
      SetScreenMousePointer vbDefault
    End If
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  TimerAutoSave.Enabled = False
  SetScreenMousePointer vbHourglass
  #If SUBCLASS Then
  mIconMenus.Detach
  DetachMessage frmMain, Me.hwnd, WM_WININICHANGE
  DetachMessage frmMain, Me.hwnd, WM_MOUSEWHEEL
  Set mIconMenus = Nothing
  #End If
  
  With Options
    If WindowState <> vbMinimized Then
      .WinMainState = WindowState
    End If
    If WindowState <> vbMaximized Then
      .WinMainLeft = Left
      .WinMainTop = Top
      .WinMainWidth = Width
      .WinMainHeight = Height
    End If
    .ObjectsListWidth = picList.Width
  End With
  
  Options.SaveData
  
  Set mScenery = Nothing
  Set mCDialog = Nothing
  Set Options = Nothing
  Set Scenery = Nothing
  Set cDialog = Nothing
  'Set mInfoTip = Nothing
  
  EraseClipboard
  
  Set picEditor = Nothing
  Set picDrag = Nothing
  
  PurgeAutoSaves
  
  Set mLang = Nothing
  Set Lang = Nothing
  Unload frmSplash   ' Just in case
  SetScreenMousePointer vbDefault
  Closing = True
End Sub

Private Sub HScroll_Change()
  HScroll_Scroll
End Sub

Private Sub HScroll_Scroll()
  mMapCenter.X = CLng(HScroll.Value) * 10
  If Not mDontUpdate Then picEdit_Resize
End Sub

#If SUBCLASS Then
Private Property Get ISubclass_MsgResponse() As EMsgResponse
  ISubclass_MsgResponse = emrPreprocess
End Property
#End If

#If SUBCLASS Then
Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim fwKeys As Long, zDelta As Long, _
    lp As POINTAPI, Moved As Boolean
  
  Select Case iMsg
    Case WM_WININICHANGE
      SetSymbols
'    Case WM_HELP
'      If MsgBoxHelpID > 0 Then
'        ' Help button in Messagebox
'        HtmlHelpLong hwnd, Lang.HelpFile & ">Error", HH_HELP_CONTEXT, MsgBoxHelpID
'      End If
    Case WM_MOUSEWHEEL
      fwKeys = wParam And 65535
      zDelta = wParam / 65536 ', lParam And 65535, lParam / 65536
      
      GetCursorPos lp
      ScreenToClient picEdit.hwnd, lp
      mDontUpdate = True

      If zDelta < 0 Then
        If fwKeys And MK_SHIFT Then
          If HScroll.Value <> HScroll.Max Then
            SendMessageLong HScroll.hwnd, WM_HSCROLL, SB_LINERIGHT, HScroll.hwnd
            Moved = True
          End If
        ElseIf fwKeys And MK_CONTROL Then
          mDontUpdate = False
          Form_KeyPress Asc("-")
        Else
          If VScroll.Value <> VScroll.Max Then
            SendMessageLong VScroll.hwnd, WM_VSCROLL, SB_LINEDOWN, VScroll.hwnd
            Moved = True
          End If
        End If
      ElseIf zDelta > 0 Then
        If fwKeys And MK_SHIFT Then
          If HScroll.Value <> HScroll.Min Then
            SendMessageLong HScroll.hwnd, WM_HSCROLL, SB_LINELEFT, HScroll.hwnd
            Moved = True
          End If
        ElseIf fwKeys And MK_CONTROL Then
          mDontUpdate = False
          Form_KeyPress Asc("+")
        Else
          If VScroll.Value <> VScroll.Min Then
            SendMessageLong VScroll.hwnd, WM_VSCROLL, SB_LINEUP, VScroll.hwnd
            Moved = True
          End If
        End If
      End If

      mDontUpdate = False
    
      If Moved Then
        picEdit_Resize
        If Dragging Then
          ' Copy the new editor to the drag buffer
          picEditor.CopyTo picDrag.hdc
        ElseIf Gridding Then
          ' Erase the old grid
          picEdit_MouseMove -2, 0, 0, 0
          ' Copy the new one to the screen
          picEditor.CopyTo picEdit.hdc
        End If
        ' By scrolling, in effect, the mouse has moved
        ' Simulate the mousemove call so that
        ' dragged items get drawn, and the grid gets drawn
        picEdit_MouseMove vbLeftButton, 0, (lp.X), (lp.Y)
      End If
  End Select
End Function
#End If

Private Sub lstObjects_DblClick()
  ' Call the appropriate EditProperties function
  If Left$(lstObjects.SelectedItem.Key, 3) = "Obj" Then
    TabValue = 0
    Scenery.EditProperties CInt(Mid$(lstObjects.SelectedItem.Key, 4))
  End If
End Sub

Private Sub lstObjects_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then lstObjects_DblClick
End Sub

Private Sub lstObjects_NodeClick(ByVal Node As ComctlLib.Node)
  Dim Index As Integer
  If Left$(Node.Key, 3) = "Obj" Then
    Index = CInt(Mid$(Node.Key, 4))
    If Index > 0 Then
      mCursorXY = MakePoint(Scenery(Index).PositionX, Scenery(Index).PositionY)
      mnuCenterHere_Click
      Scenery.SetSingleFocus Index
    End If
  End If
End Sub

Private Sub mnu2D_Click(Index As Integer)
  AddSceneryObject mnu2D(Index).Tag
End Sub

Private Sub mnu3D_Click(Index As Integer)
  AddSceneryObject mnu3D(Index).Tag
End Sub

Private Sub mnuAbout_Click()
  frmAbout.Show vbModal, Me
End Sub

Private Sub mnuAcknowledgements_Click()
  HtmlHelpString hwnd, Lang.HelpFile, HH_DISPLAY_TOPIC, "Appendix\Thanks.html"
End Sub

Private Sub mnuAdd_Click(Index As Integer)
  AddSceneryObject mnuAdd(Index).Tag
End Sub

Private Sub mnuATIS_Click(Index As Integer)
  frmATIS.Txts(5).SelText = "%" & Chr$(mnuATIS(Index).Tag - 1600)
End Sub

Private Sub mnuCenterHere_Click()
  Dim NewX As Single, NewY As Single
  ' Recenter the map
  With Scenery.Header
    If Not (Abs(mCursorXY.X) > .Horz / 2) And Not (Abs(mCursorXY.Y) > .Vert / 2) Then
      mDontUpdate = True
      HScroll.Value = mCursorXY.X / 10
      VScroll.Value = -mCursorXY.Y / 10
      mMapCenter = MakePoint(mCursorXY.X, mCursorXY.Y)
      mDontUpdate = False
      picEdit_Resize
      picEditor.ScaleToPixel mCursorXY.X, mCursorXY.Y, NewX, NewY
      picEdit_MouseMove 0, 0, NewX, NewY
    End If
  End With
End Sub

Private Sub mnuContact_Click()
  Dim strRun As String
  strRun = "mailto:" & Email & "?subject=FS%20Scenery%20Creator%20Comments&body=Hello,"
  ShellExecute hwnd, "Open", strRun, "", "", vbNormalFocus
End Sub

Private Sub mnuCopy_Click()
  EraseClipboard
  Form_KeyPress vbKeyEscape
  Scenery.DoCopy mClipboard
  
  mnuPaste.Enabled = True
  Toolbar1.Buttons(9).Enabled = True ' Paste
End Sub

Private Sub mnuCut_Click()
  Form_KeyPress vbKeyEscape
  mnuCopy_Click
  Scenery.DoDelete True
End Sub

Private Sub mnuDelete_Click()
  Form_KeyPress vbKeyEscape
  Scenery.DoDelete False
End Sub

Private Sub mnuExit_Click()
  Unload Me
End Sub

Private Sub mnuExport_Click()
  If Scenery.Changed And Options.SaveBeforeCompile Then
    If Not AskIfSave() Then Exit Sub
  End If
    
  If Scenery.ExportPath = "" Or Options.ShowExportWizard Then
    frmExport.DoWizard
  Else
    frmExport.DoCompile
  End If
End Sub

Private Sub mnuExportWizard_Click()
  If Scenery.Changed And Options.SaveBeforeCompile Then
    If Not AskIfSave() Then Exit Sub
  End If

  frmExport.DoWizard
End Sub

Private Sub mnuHelp_Click(Index As Integer)
  Select Case Index
    Case 0 ' Help Topics
      HtmlHelpString hwnd, Lang.HelpFile, HH_DISPLAY_TOPIC, "Index.html"
    Case 1 ' Tutorial
      HtmlHelpString hwnd, Lang.HelpFile & ">Tutorial", HH_DISPLAY_TOPIC, "tutorials\index.html"
    Case 2 ' FAQ
      HtmlHelpString hwnd, Lang.HelpFile, HH_DISPLAY_TOPIC, "FAQ.html"
    Case 3 ' What's New
      HtmlHelpString hwnd, Lang.HelpFile, HH_DISPLAY_TOPIC, "Appendix\History.html"
    Case 4 ' Tip of the day
      frmTip.Show vbModal, Me
    Case 5 ' SCASM reference
      MsgBoxEx Me, "Sorry, this feature has not been implemented yet. Thank you for your patience.", vbInformation, 0
  End Select
End Sub

Private Sub mnuImport_Click()
  MsgBoxEx Me, "Sorry, this feature has not been implemented yet. Thank you for your patience.", vbInformation, 0
End Sub

Private Sub mnuLanguage_Click()
  If frmLanguage.EditData Then
    If mnuObjects.Checked Then
      UpdateObjectList
    End If
  End If
End Sub

Private Sub mnuMainProperties_Click()
  Dim I As Integer, Result As Boolean, _
    Objects() As clsObject, MyObjectType As ObjectTypeEnum, _
    DifferentType As Boolean, PointType As clsPoint

  If mnuMainProperties.Tag > 0 Then
    ' Single
    Scenery.EditProperties mnuMainProperties.Tag
  Else
    ' Multi
    ReDim Objects(UBound(mFocusIndices)) As clsObject
    MyObjectType = Scenery(mFocusIndices(1)).ObjectType
    For I = 1 To UBound(Objects)
      Set Objects(I) = Scenery(mFocusIndices(I))
      DifferentType = DifferentType Or Objects(I).ObjectType <> MyObjectType
    Next I
    If DifferentType Then
      Result = frmProperties.EditMulti(Objects)
    Else
      Select Case MyObjectType
        Case OT_Runway
          Result = frmProperties.EditMulti(Objects)
        Case OT_Building
          Result = frmBuilding.EditDataM(Objects)
        Case OT_Macro
          Result = frmMacro.EditDataM(Objects)
        Case OT_ATIS
          Result = frmProperties.EditMulti(Objects)
        Case OT_VOR
          Result = frmProperties.EditMulti(Objects)
        Case OT_NDB
          Result = frmProperties.EditMulti(Objects)
        Case OT_MenuEntry
          Result = frmProperties.EditMulti(Objects)
        Case OT_Background
          Result = frmProperties.EditMulti(Objects)
        Case OT_FlatArea
          Result = frmProperties.EditMulti(Objects)
        Case OT_Exclusion
          Result = frmProperties.EditMulti(Objects)
        Case OT_Code
          Result = frmProperties.EditMulti(Objects)
        Case OT_Point
          Result = frmProperties.EditDataM(Objects)
          
          ' Special case
          If Result Then
            For I = 1 To UBound(Objects)
              Set PointType = Objects(I)
              PointType.Parent.BoolTag = True
              Set PointType = Nothing
            Next I
            For I = 1 To UBound(Objects)
              Set PointType = Objects(I)
              If PointType.Parent.BoolTag Then
                PointType.Parent.Point(0).UpdateObject
                PointType.Parent.BoolTag = False
              End If
              Set PointType = Nothing
            Next I
            Scenery.Draw
            Scenery.Changed = True
          End If
          Result = False
        Case Else
          ' Error!
      End Select
    End If
    If Result Then
      For I = 1 To UBound(Objects)
        Objects(I).UpdateObject
        Set Objects(I) = Nothing
      Next I
      Scenery.Draw
      Scenery.Changed = True
    End If
  End If
End Sub

Private Sub mnuMisc_Click(Index As Integer)
  AddSceneryObject mnuMisc(Index).Tag
End Sub

Private Sub mnuNew_Click()
  If AskIfSave Then
    Scenery.NewFile
    If Options.ShowHeaderProperties Then Scenery.Header.EditProperties
  End If
End Sub

Private Sub mnuObjects_Click()
  If mnuObjects.Checked Then
    picList.Visible = False
    Splitter.Visible = False
    Toolbar1.Buttons(13).Value = tbrUnpressed
    Form_Resize
    mnuObjects.Checked = False
  Else
    LoadObjects
    Toolbar1.Buttons(13).Value = tbrPressed
    picList.Visible = True
    Splitter.Visible = True
    Form_Resize
    mnuObjects.Checked = True
  End If
End Sub

Private Sub mnuOpen_Click()
  Dim X As String
  If AskIfSave Then
    With cDialog
      .Filter = Lang.GetString(RES_OpenFilter)
      .FilterIndex = 1
      .DefExt = "scn"
      X = .OpenDialog(Lang.GetString(RES_Main_OpenCaption))
    End With

    If X <> "" Then
      Scenery.LoadFile X
      Options.AddRecentFile Scenery.File
    End If
  End If
End Sub

Private Sub mnuOrder_Click(Index As Integer)
  Select Case Index
    Case 0
      Scenery.BringToFront
    Case 1
      Scenery.SendToBack
  End Select
End Sub

Private Sub mnuPaste_Click()
  Form_KeyPress vbKeyEscape
  Scenery.DoPaste mClipboard
End Sub

Private Sub mnuPointAdd_Click()
  Dim PointObject As clsPoint
  Set PointObject = Scenery(Scenery.SingleFocus)
  PointObject.Parent.MidPointInsert PointObject.ObjectIndex - 1
  Set PointObject = Nothing
End Sub

Private Sub mnuPopOrder_Click(Index As Integer)
  mnuOrder_Click Index
End Sub

Private Sub mnuPredefinedTools_Click(Index As Integer)
  Dim X As String, rc As Boolean, _
    OldDir As String, NewDir As String, FileTitle As String, _
    FileList As String, Files() As String, _
    Linker As String, I As Integer
  
  Select Case Index
    Case 0 ' Compile
      With cDialog
        .Filter = Lang.GetString(RES_CompileFilter)
        .FilterIndex = 1
        .DefExt = "sca"
        X = .OpenDialog(Lang.GetString(RES_Main_CompileCaption))
      End With
  
      If X <> "" Then
        If FileExists(Options.Compiler) Then
          OldDir = CurDir$
          NewDir = GetDir(X)
          FileTitle = GetFileTitle(X)
          ChangeDir NewDir
          RunSCASM X, AddDir(NewDir, Left$(FileTitle, InStrRev(FileTitle, ".") - 1) & ".bgl")
          ChangeDir OldDir
        Else
          MsgBoxEx frmMain, Lang.GetString(RES_ERR_CompilerPath), vbCritical, RES_ERR_CompilerPath
        End If
      End If
    Case 1 ' Link
      With cDialog
        .Filter = Lang.GetString(RES_LinkFilter)
        .FilterIndex = 1
        .DefExt = "bgl"
        rc = .SelectMultiDialog(Files(), Lang.GetString(RES_Main_LinkCaption))
        If rc Then X = .SaveDialog("Output.bgl", Lang.GetString(RES_Main_LinkOutputCaption))
      End With
  
      If X <> "" And rc Then
        Linker = AddDir(App.Path, "SCASM\SCLINK.EXE")
        If FileExists(Linker) Then
          For I = 0 To UBound(Files)
            FileList = FileList & " " & GetShortName(Files(I))
          Next I
          RunDosFile Linker, X & FileList, GetDir(X)
        Else
          MsgBoxEx frmMain, Lang.GetString(RES_ERR_CompilerPath), vbCritical, RES_ERR_CompilerPath
        End If
      End If
  End Select
End Sub

Private Sub mnuPreferences_Click()
  Options.EditData
End Sub

Private Sub mnuPrograms_Click(Index As Integer)
  On Error Resume Next
  Shell mnuPrograms(Index).Tag, vbNormalFocus
  If Err > 0 Then
    MsgBoxEx Me, Error$, vbCritical, IDH_ERR_Program
  End If
End Sub

Private Sub mnuProperties_Click(Index As Integer)
  Scenery(mnuProperties(Index).Tag).EditProperties
End Sub

Private Sub mnuRadio_Click(Index As Integer)
  AddSceneryObject mnuRadio(Index).Tag
End Sub

Private Sub mnuRecentFiles_Click(Index As Integer)
  If AskIfSave() Then
    Scenery.LoadFile Options.RecentFiles(Index - 1)
    Options.AddRecentFile Scenery.File
  End If
End Sub

Private Sub mnuSave_Click()
  SaveAsProc False
End Sub

Private Sub mnuSaveAs_Click()
  SaveAsProc True
End Sub

Private Sub mnuSceneryProperties_Click()
  Scenery.EditProperties 0
  picEdit_Resize
End Sub

Private Sub mnuScrollbars_Click()
  Dim Value As Boolean
  Value = Not mnuScrollbars.Checked
  mnuScrollbars.Checked = Value
  Options.ScrollbarsVisible = Value
  VScroll.Visible = Value
  HScroll.Visible = Value
  Form_Resize
End Sub

Private Sub mnuSelectAll_Click()
  Dim I As Integer, PointObject As clsPoint, _
    ShapeObject As clsShape
  Set PointObject = Scenery(Scenery.SingleFocus)
  Set ShapeObject = PointObject.Parent
  For I = 0 To ShapeObject.NumPoints
    Scenery.Focus(ShapeObject.Point(I).SceneryIndex) = True
  Next I
  
  Set ShapeObject = Nothing
  Set PointObject = Nothing
  mScenery_Redraw
End Sub

Private Sub mnuSortObjects_Click()
  Scenery.ResortScenery
  UpdateObjectList
End Sub

Private Sub mnuStandard_Click()
  mMapCenter = MakePoint(0, 0)
  mDontUpdate = True
  HScroll.Value = 0
  VScroll.Value = 0
  mDontUpdate = False
  mMetersPerPixel = 4
  ReAdjustZoomButtons
End Sub

Private Sub mnuStatusbar_Click()
  Dim Value As Boolean
  Value = Not mnuStatusbar.Checked
  mnuStatusbar.Checked = Value
  Options.StatusbarVisible = Value
  Statusbar.Visible = Value
  Form_Resize
End Sub

Private Sub mnuToolbar_Click()
  Dim Value As Boolean
  Value = Not mnuToolbar.Checked
  mnuToolbar.Checked = Value
  Options.ToolbarVisible = Value
  picToolbar.Visible = Value
  Form_Resize
End Sub

Private Sub mnuTransform_Click()
  Dim X As Single, Y As Single, R As Single, A As Boolean
  If frmTransform.EditData(X, Y, R, A) Then
    Scenery.TransformScenery X, Y, R, A, True
  End If
End Sub

Private Sub mnuWeb_Click()
  ShellExecute hwnd, "Open", Webpage, "", "", vbNormalFocus
End Sub

Private Sub mnuZoom_Click(Index As Integer)
  Dim I As Integer
  ' Find the next standard zoom setting
  If Index = 0 Then
    For I = UBound(ZoomLevels) To 0 Step -1
      If ZoomLevels(I) < mMetersPerPixel Then
        mMetersPerPixel = ZoomLevels(I)
        Exit For
      End If
    Next I
  Else
    For I = 0 To UBound(ZoomLevels)
      If ZoomLevels(I) > mMetersPerPixel Then
        mMetersPerPixel = ZoomLevels(I)
        Exit For
      End If
    Next I
  End If
  ReAdjustZoomButtons
End Sub

Private Sub mnuZoomSpecify_Click()
  If frmZoom.EditData(mMetersPerPixel) Then
    ReAdjustZoomButtons
  End If
End Sub

Private Sub mLang_Changed()
  Dim I As Integer
  
  SetScreenMousePointer vbHourglass
  
  ' Resolve language strings
  Lang.PrepareForm Me
  ProcessMenus mnuAdd, mnuAdd.UBound
  ProcessMenus mnu2D, mnu2D.UBound
  ProcessMenus mnu3D, mnu3D.UBound
  ProcessMenus mnuRadio, mnuRadio.UBound
  ProcessMenus mnuMisc, mnuMisc.UBound

  ' Add non-standard shortcut keys
  mnuExit.Caption = mnuExit.Caption & vbTab & "Alt+F4"
  mnuZoom(0).Caption = mnuZoom(0).Caption & vbTab & "+"
  mnuZoom(1).Caption = mnuZoom(1).Caption & vbTab & "-"
  mnuPointAdd.Caption = mnuPointAdd.Caption & vbTab & "Ins"

  ' Add tooltip text for toolbar buttons
  For I = 1 To Toolbar1.Buttons.Count
    Toolbar1.Buttons(I).ToolTipText = Lang.GetString(RES_Toolbar1 + I)
  Next I
  
  ' If scenery name is untitled, change it to the
  ' new localized scenery untitled name
  If Not Scenery Is Nothing Then
    If Scenery.Header.Name = UntitledName Then
      Scenery.Header.Name = Lang.GetString(RES_Untitled)
      Scenery.File = Lang.GetString(RES_UntitledFile)
    End If
  End If
  UntitledName = Lang.GetString(RES_Untitled)
  
  SetScreenMousePointer vbDefault
End Sub

Private Sub mOptions_Changed()
  Dim I As Integer, Value As Boolean
  
  SetScreenMousePointer vbHourglass
  
  With Options
    Value = .OldStyleMenus
    For I = 0 To mnuAdd.UBound
      mnuAdd(I).Visible = Value
    Next I
    Value = Not Value
    For I = 0 To mnuInsert.UBound
      mnuInsert(I).Visible = Value
    Next I
    
    Toolbar1.Buttons(14).Value = IIf(.FillObjects And .FillPolygons, tbrPressed, tbrUnpressed)
    mnuToolbar.Checked = .ToolbarVisible
    mnuStatusbar.Checked = .StatusbarVisible
    mnuScrollbars.Checked = .ScrollbarsVisible
    
    mnuPrograms(0).Visible = (.ToolCount > 0)
    For I = .ToolCount + 1 To mnuPrograms.UBound
      Unload mnuPrograms(I)
    Next I
    For I = mnuPrograms.UBound + 1 To .ToolCount
      Load mnuPrograms(I)
      mnuPrograms(I).Visible = True
    Next I
    For I = 1 To .ToolCount
      mnuPrograms(I).Caption = .ToolName(I)
      mnuPrograms(I).Tag = .ToolExe(I)
    Next I
    
    Scenery.Update .FSVersion
    
    mnu2D(6).Enabled = .FSVersion >= Version_FS2K2
    mnuRadio(0).Enabled = .FSVersion < Version_FS2K
    mnuRadio(3).Enabled = .FSVersion >= Version_FS2K
    mnuMisc(2).Enabled = .FSVersion >= Version_FS2K
    
    mnuAdd(0).Enabled = mnuRadio(0).Enabled
    mnuAdd(3).Enabled = mnuRadio(3).Enabled
    mnuAdd(13).Enabled = mnu2D(6).Enabled
    mnuAdd(18).Enabled = mnuMisc(2).Enabled
    
    picEdit.MousePointer = IIf(.CrossHair, vbCrosshair, vbDefault)
  
    NextAutoSave = (Timer + .AutoSave * 60) Mod 86400
  End With
  
  LoadMacros
  mOptions_RecentFileChanged
  
  Scenery.Draw

  SetScreenMousePointer vbDefault
End Sub

Private Sub mOptions_RecentFileChanged()
  Dim I As Integer, Temp As String
  With Options
    mnuRecentFiles(0).Visible = (Options.RecentFiles(0) <> "")
    For I = 1 To 4
      Temp = .RecentFiles(I - 1)
      mnuRecentFiles(I).Visible = (Temp <> "")
      If .NeatRecentFiles Then
        mnuRecentFiles(I).Caption = "&" & CStr(I) & " " & MakeFileNameNeat(Temp)
      Else
        mnuRecentFiles(I).Caption = "&" & CStr(I) & " " & Temp
      End If
    Next I
  End With
End Sub

Private Sub mScenery_ClearValues()
  mnuStandard_Click
End Sub

Private Sub mScenery_FocusChanged(NewFoci() As Integer)
  ' Make menus visible/invisible
  Dim I As Integer, Value As Boolean, MaxUBound As Integer
  
  MaxUBound = UBound(NewFoci)
  If MaxUBound > 100 Then MaxUBound = 100
  
  For I = MaxUBound + 1 To mnuProperties.UBound
    Unload mnuProperties(I)
  Next I
  
  Select Case UBound(NewFoci)
    Case Is > 1
      For I = mnuProperties.UBound + 1 To MaxUBound
        Load mnuProperties(I)
        mnuProperties(I).Visible = True
      Next I
      mnuPopDash3.Visible = True
      mnuMainProperties.Visible = True
      mnuProperties(0).Visible = True
      For I = 1 To MaxUBound
        mnuProperties(I).Caption = Lang.ResolveString(RES_Menu_Properties2, Scenery(NewFoci(I)).Caption)
        mnuProperties(I).Tag = NewFoci(I)
      Next I
      mnuMainProperties.Caption = Lang.GetString(RES_Menu_Properties)
      mnuMainProperties.Tag = -1
      Toolbar1.Buttons(12).Enabled = True
    Case 1
      If mnuProperties.UBound = 1 Then Unload mnuProperties(1)
      mnuPopDash3.Visible = True
      mnuMainProperties.Visible = True
      mnuProperties(0).Visible = False
      mnuMainProperties.Caption = Lang.ResolveString(RES_Menu_Properties2, Scenery(NewFoci(1)).Caption)
      mnuMainProperties.Tag = NewFoci(1)
      Toolbar1.Buttons(12).Enabled = True
    Case Else
      mnuPopDash3.Visible = False
      mnuMainProperties.Visible = False
      mnuProperties(0).Visible = False
      Toolbar1.Buttons(12).Enabled = False
  End Select
    
  Value = (UBound(NewFoci) > 0)

  mnuCut.Enabled = Value
  mnuCopy.Enabled = Value
  mnuDelete.Enabled = Value
  Toolbar1.Buttons(7).Enabled = Value
  Toolbar1.Buttons(8).Enabled = Value
  Toolbar1.Buttons(10).Enabled = Value
  Toolbar1.Buttons(16).Enabled = Value
  Toolbar1.Buttons(17).Enabled = Value

  If lstObjects.Visible And Value Then
    Set lstObjects.SelectedItem = lstObjects.Nodes("Obj" & CStr(NewFoci(1)))
    lstObjects.SelectedItem.EnsureVisible
  End If

  Value = (UBound(NewFoci) = 1)
  For I = 0 To 1
    mnuOrder(I).Enabled = Value
    mnuPopOrder(I).Enabled = Value
  Next I
  Toolbar1.Buttons(19).Enabled = Value
  Toolbar1.Buttons(20).Enabled = Value
  
  If Value Then Value = Scenery(NewFoci(1)).ObjectType = OT_Point
  mnuSelectAll.Enabled = Value
  mnuPopDash2.Visible = Value
  mnuPointAdd.Visible = Value

  Value = Scenery.Count > 0
  mnuSortObjects.Enabled = Value
'  mnuObjects.Enabled = Value
'  Toolbar1.Buttons(13).Enabled = Value
  mnuTransform.Enabled = Value

  mFocusIndices = NewFoci
End Sub

Private Sub mScenery_Redraw()
  picEditor.CopyTo picEdit.hdc
End Sub

Private Sub mScenery_TitleBarChange()
  Caption = Scenery.Header.Name & " - " & Lang.GetString(RES_LongTitle)
End Sub

Private Sub picButton_Click(Index As Integer)
  If Index = 0 Then
    If picList.Left = 0 Then
      picList.Left = ScaleWidth - picList.Width
    Else
      picList.Left = 0
    End If
    picButton_Paint 0
    Form_Resize
  Else
    mnuObjects_Click
  End If
End Sub

Private Sub picButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  picButton_MouseMove Index, Button, Shift, X, Y
End Sub

Private Sub picButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Static T As Boolean
  Dim Char As String
  
  If Shift = -1 Then
    ' From picButton_Paint
    T = True
    Shift = 0
  End If
  
  If Button = vbLeftButton And Shift = 0 Then
    If Between(X, 0, 16) And Between(Y, 0, 16) Then
      If T = True Then Exit Sub
      Dim R As RECT
      R.Left = 0
      R.Right = 16
      R.Top = 0
      R.Bottom = 16
      DrawEdge picButton(Index).hdc, R, BDR_SUNKENOUTER, BF_RECT
      T = True
    Else
      If T = False Then Exit Sub
      picButton(Index).Line (0, 0)-(15, 15), vbButtonFace, BF
          
      If Index = 1 Then
        Char = "r"
      ElseIf picList.Left = 0 Then
        Char = "8"
      Else
        Char = "w"
      End If
    
      With picButton(Index)
        TextOut .hdc, (.Width / 2 - .TextWidth(Char) / 2) + 1, (.Height / 2 - .TextHeight(Char) / 2), Char, 1
      End With
      T = False
    End If
  End If
End Sub

Private Sub picButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  picButton_MouseMove Index, Button, Shift, -1, -1
End Sub

Private Sub picButton_Paint(Index As Integer)
  picButton_MouseMove Index, vbLeftButton, -1, -1, -1
End Sub

Private Sub picEdit_DblClick()
  Dim lp As POINTAPI, NewX As Single, NewY As Single
  GetCursorPos lp
  ScreenToClient picEdit.hwnd, lp
  ' Recenter the map
  With Scenery.Header
    picEditor.PixelToScale lp.X, lp.Y, NewX, NewY

    If Not (Abs(NewX) > .Horz / 2) And Not (Abs(NewY) > .Vert / 2) Then
      mDontUpdate = True
      HScroll.Value = NewX / 10
      VScroll.Value = -NewY / 10
      mMapCenter = MakePoint(NewX, NewY)
      mDontUpdate = False
      picEdit_Resize
      picEdit_MouseMove 0, 0, CSng(lp.X), CSng(lp.Y)
    End If
  End With
End Sub

Private Sub picEdit_DragDrop(Source As Control, X As Single, Y As Single)
  If Source.Name = "picList" Then picButton_Click 0
End Sub

Private Sub picEdit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim CurPos As POINTAPI, _
    TX As Single, TY As Single, _
    R As Integer, I As Integer, _
    Count As Integer, _
    PointObject As clsPoint
  Static OldShift As Integer
  
  mInMouseAction = True
  
  picEditor.PixelToScale X, Y, TX, TY
  mCursorXY = MakePoint(TX, TY)
  
  If (Shift = 0 Or Shift = vbShiftMask) And Button = vbRightButton Then
    ' Cancel grid
    If Gridding Then picEdit_MouseUp vbLeftButton, 0, 0, 0
    ' Find the nearest item
    R = Scenery.NearestObject(TX, TY, ((Shift And vbShiftMask) = 0))
    
    ' If the current item is not already selected, then
    ' reset all the selections, and select the item,
    ' else, keep the selections as they are
    If Not Scenery.Focus(R) Then
      Scenery.InsideFocusArea TX, TY, ((Shift And vbShiftMask) = 0)
    ElseIf Shift <> OldShift Then
      Scenery.RefreshFocus ((Shift And vbShiftMask) = 0)
    End If
    
    OldShift = Shift
    
    GetCursorPos CurPos
    
    ' Use TrackPopupMenu because you can't nest
    ' "PopupMenu"
    ' i.e. TrackPopupMenu returns as soon as menu is
    ' selected, and PopupMenu returns only after
    ' command is processed
    ' Since ATIS also uses a popupmenu, we need to use
    ' the API here
    mnuRadio(3).Enabled = (Options.FSVersion >= Version_FS2K2) And Scenery.TowerIndex = 0
    mnuAdd(3).Enabled = mnuRadio(3).Enabled
    
    mnuPopup.Visible = True
    TrackPopupMenu GetSubMenu(GetSubMenu(GetMenu(hwnd), 4), 11), TPM_LEFTALIGN Or TPM_RIGHTBUTTON, CurPos.X, CurPos.Y, 0, hwnd, 0&
    mnuPopup.Visible = False
  ElseIf (Shift And vbCtrlMask) And Button = vbLeftButton Then
    ' Toggle the focus state of the object
    R = Scenery.NearestObject(TX, TY, (Shift And vbShiftMask) = 0)
    If R > 0 Then
      Scenery.Focus(R) = Not Scenery.Focus(R)
      Scenery.RefreshFocus ((Shift And vbShiftMask) = 0)
    End If
  ElseIf (Shift = 0 Or Shift = vbShiftMask) And Button = vbLeftButton Then
    ' Find which object is the user actually dragging
    R = Scenery.NearestObject(TX, TY, (Shift And vbShiftMask) = 0)

    If R = 0 Then
      If Scenery.OnPolyLine(TX, TY) Then
        ' If user clicked a line, create point and start drag
        ' (The newly created point is the last item)
        R = Scenery.Count
        mRedrawOverride = True
      End If
    End If

    If R <> 0 Then
      ' Drag
      If Not Scenery.Focus(R) Then Scenery.SetSingleFocus R
      Dragging = True
      ' Keep the cache of the old picture
      With picEditor
        picDrag.PhysicalResize .myWidth, .myHeight
        picDrag.SetScale .myLeft, .myTop, .myRight, .myBottom
        picEditor.CopyTo picDrag.hdc
      End With
      ReDim mRedrawIndices(10)
      For I = 1 To Scenery.Count
        If Scenery.Focus(I) Then
          If Not Scenery(I).Locked Then
            Count = Count + 1
            If Count > UBound(mRedrawIndices) Then _
              ReDim Preserve mRedrawIndices(Count * 2)
            mRedrawIndices(Count) = I
            If TypeOf Scenery(I) Is clsPoint Then
              Set PointObject = Scenery(I)
              PointObject.Parent.BoolTag = True
              Set PointObject = Nothing
            End If
          End If
        End If
      Next I
      ReDim Preserve mRedrawIndices(Count)
      Count = 0
      
      ReDim mShapeRedraw(10)
      For I = 1 To Scenery.Count
        If TypeOf Scenery(I) Is clsPoint Then
          Set PointObject = Scenery(I)
          If PointObject.Parent.BoolTag Then
            Count = Count + 1
            If Count > UBound(mShapeRedraw) Then _
              ReDim Preserve mShapeRedraw(Count * 2)
            Set mShapeRedraw(Count) = PointObject.Parent
            PointObject.Parent.BoolTag = False
          End If
          Set PointObject = Nothing
        End If
      Next I
      ReDim Preserve mShapeRedraw(Count)
    Else
      ' This is the start of a multiple select (Grid)
      Scenery.ClearFocus
      Gridding = True
    End If
  End If
End Sub

Private Sub picEdit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim TX As Single, TY As Single, _
    PX As Single, PY As Single, _
    R As Integer, I As Integer, _
    Offset As PointType, Corner1 As PointType, _
    Corner2 As PointType
  Dim DoSnapToLine As Boolean
  
  Static cYes As Boolean, Coord As PointType
  
  picEditor.PixelToScale X, Y, TX, TY
  
  With Statusbar
    .Style = sbrNormal
    .Panels(1).Text = Lang.ResolveString(RES_Main_CurPos, ReturnPoint(TX, TY).ToString)
    If Not Gridding Then
      If Options.Metric Then
        .Panels(2).Text = Lang.ResolveString(RES_Main_CurPosXY, Round(TX, 2), Round(TY, 2), Lang.GetString(RES_Unit_M))
      Else
        .Panels(2).Text = Lang.ResolveString(RES_Main_CurPosXY, Round(TX * MToFt, 3), Round(TY * MToFt, 3), Lang.GetString(RES_Unit_Ft))
      End If
    End If
  End With

  If (Button = vbLeftButton) And Dragging Then
    Offset = MakePoint(TX - mCursorXY.X, TY - mCursorXY.Y)
    cYes = True
    TimerMouseMove.Enabled = True
    picDrag.CopyTo picEditor.hdc

    ' Draw dragged objects
    
    picEditor.StartDraw
    
    DoSnapToLine = Options.SnapPoints
    For I = 1 To UBound(mRedrawIndices)
      With Scenery(mRedrawIndices(I))
        PX = Offset.X + .PositionX
        PY = Offset.Y + .PositionY
        
        If DoSnapToLine And .ObjectType = OT_Point Then
          ' PX and PY are passed byref .. if any of the functions change PX and PY, it returns true
          If Scenery.NearestPoint(PX, PY) Then
          
          ElseIf Scenery.NearestLine(PX, PY) Then
          
          End If
        End If
        .DrawBottom PX, PY
      End With
    Next I
    
    For I = 1 To UBound(mShapeRedraw)
      mShapeRedraw(I).DrawBottom 0, 0
    Next I
    
    For I = 1 To UBound(mRedrawIndices)
      With Scenery(mRedrawIndices(I))
        .DrawTop Offset.X + .PositionX, Offset.Y + .PositionY
      End With
    Next I
    
    For I = 1 To UBound(mShapeRedraw)
      mShapeRedraw(I).DrawTop 0, 0
    Next I
    
    picEditor.CopyTo picEdit.hdc
  ElseIf (Button = vbLeftButton Or Button < 0) And Gridding Then
    TimerMouseMove.Enabled = True
    ' Multiple select
    ' Erase the previous grid if needed
    
    picEditor.ScaleToPixel mCursorXY.X, mCursorXY.Y, Corner1.X, Corner1.Y
    picEditor.ScaleToPixel Coord.X, Coord.Y, Corner2.X, Corner2.Y
    ' Don't draw the grid if delta X, or Y = 0
    ' (cYes keeps track of whether a grid
    '  was previously drawn)
    
    picEdit.DrawStyle = vbDash
    picEdit.DrawMode = vbNotXorPen
    If Corner1.X <> Corner2.X And Corner1.Y <> Corner2.Y Then
      If cYes Then picEdit.Line (Corner1.X, Corner1.Y)-(Corner2.X, Corner2.Y), , B
    End If
    If Button = vbLeftButton Then
      picEditor.PixelToScale X, Y, Coord.X, Coord.Y
      cYes = True
      ' Draw the grid
      If Corner1.X <> X And Corner1.Y <> Y Then picEdit.Line (Corner1.X, Corner1.Y)-(X, Y), 0, B
    End If
    
    picEdit.DrawStyle = vbSolid
    picEdit.DrawMode = vbCopyPen
    
    ' Distance measurement
    If Options.Metric Then
      Statusbar.Panels(2).Text = Lang.ResolveString(RES_Main_Distance, Round(Distance(mCursorXY.X, mCursorXY.Y, TX, TY), 2), Lang.GetString(RES_Unit_M))
    Else
      Statusbar.Panels(2).Text = Lang.ResolveString(RES_Main_Distance, Round(Distance(mCursorXY.X, mCursorXY.Y, TX, TY) * MToFt, 3), Lang.GetString(RES_Unit_Ft))
    End If
  Else
    R = Scenery.NearestObject(TX, TY, (Shift And vbShiftMask) = 0)
    If R > 0 Then
      With Scenery(R)
        If Options.Metric Then
          picEdit.ToolTipText = Scenery(R).Caption & "   " & Round(Scenery(R).PositionX, 2) & ", " & Round(Scenery(R).PositionY, 2)
        Else
          picEdit.ToolTipText = Scenery(R).Caption & "   " & Round(Scenery(R).PositionX * MToFt, 3) & ", " & Round(Scenery(R).PositionY * MToFt, 3)
        End If
      End With
    Else
      picEdit.ToolTipText = ""
    End If
    If TypeOf Scenery(R) Is clsMacro Then
      Dim MacroObject As clsMacro
      Set MacroObject = Scenery(R)
      If Left$(MacroObject.File, 7) = "LibObj:" Then
        Statusbar.Panels(3).Text = IIf(R = 0, "", Lang.ResolveString(RES_Main_Object, Scenery(R).Caption))
      Else
        Statusbar.Panels(3).Text = IIf(R = 0, "", Lang.ResolveString(RES_Main_Object, Scenery(R).Caption)) & " - " & GetFileTitle(MacroObject.File)
      End If
      Set MacroObject = Nothing
    Else
      Statusbar.Panels(3).Text = IIf(R = 0, "", Lang.ResolveString(RES_Main_Object, Scenery(R).Caption))
    End If
  End If
  If Button < 0 Then cYes = False
End Sub

Private Sub picEdit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim I As Integer, TX As Single, TY As Single, _
    PX As Single, PY As Single, _
    DoSnapToLine As Boolean
  
  picEditor.PixelToScale X, Y, TX, TY
  
  If Button = vbLeftButton And Dragging Then
    ' No need to redraw if nothing was dragged
    ' Note that a left mouse click is also a drag,
    ' so this saves processing
    If mCursorXY.X <> TX Or mCursorXY.Y <> TY Or mRedrawOverride Then
      ' Move each object which was drag into position
      
      DoSnapToLine = Options.SnapPoints
      For I = 1 To UBound(mRedrawIndices)
        With Scenery(mRedrawIndices(I))
          PX = TX + .PositionX - mCursorXY.X
          PY = TY + .PositionY - mCursorXY.Y
          
          If DoSnapToLine And .ObjectType = OT_Point Then
            If Scenery.NearestPoint(PX, PY) Then
              .PositionX = PX
              .PositionY = PY
            ElseIf Scenery.NearestLine(PX, PY) Then
              .PositionX = PX
              .PositionY = PY
            Else
              .PositionX = Round(PX, 2)
              .PositionY = Round(PY, 2)
            End If
          Else
            .PositionX = Round(PX, 2)
            .PositionY = Round(PY, 2)
          End If
        End With
      Next I
      
      Scenery.Draw
      Scenery.Changed = True
      Erase mRedrawIndices
      mRedrawOverride = False
    Else
      ' We still need to redraw the screen since the gray
      ' objects were already drawn over
      picDrag.CopyTo picEditor.hdc
      mScenery_Redraw
    End If
    picEdit_MouseMove -1, 0, X, Y
  ElseIf Button = vbLeftButton And Gridding Then
    ' End of gridding
    picEdit_MouseMove -1, 0, X, Y
    Scenery.SelectRectFocus mCursorXY.X, mCursorXY.Y, TX, TY
  End If
  TimerMouseMove.Enabled = False
  TimerMouseMove.Interval = 1000
  If Dragging Then
    For I = 1 To UBound(mShapeRedraw)
      Set mShapeRedraw(I) = Nothing
    Next I
    ReDim mShapeRedraw(0)
  End If
  Dragging = False
  Gridding = False
  mInMouseAction = False
End Sub

Private Sub picEdit_Paint()
  mScenery_Redraw
  ' Cover up the hole in the lower right corner
  If VScroll.Visible Then
    picEdit.Line (HScroll.Width, VScroll.Height)-(picEdit.ScaleWidth, picEdit.ScaleHeight), vbButtonFace, BF
  End If
End Sub

Private Sub picEdit_Resize()
  If Not mDontUpdate Then
    picEditor.PhysicalResize picEdit.ScaleWidth - IIf(VScroll.Visible, VScroll.Width, 0), picEdit.ScaleHeight - IIf(VScroll.Visible, HScroll.Height, 0)
    SetScale
    Scenery.Draw
    ' Cover up the hole in the lower right corner
    If VScroll.Visible Then
      picEdit.Line (HScroll.Width, VScroll.Height)-(picEdit.ScaleWidth, picEdit.ScaleHeight), vbButtonFace, BF
    End If
  End If
End Sub

Private Sub picList_Paint()
  Dim I As Integer
  picList.Line (0, 0)-(picList.ScaleWidth, 15), picList.BackColor, BF
  picList.Line (0, 0)-(picList.ScaleWidth + 1, 0), vbButtonShadow
  picList.Line (0, 1)-(picList.ScaleWidth + 1, 1), vb3DHighlight

  For I = 9 To 12 Step 3
    picList.Line (25, I)-(picList.ScaleWidth - 24, I), vb3DHighlight
    picList.Line (25, I + 2)-(picList.ScaleWidth - 24, I + 2), vbButtonShadow
    picList.PSet (25, I + 1), vb3DHighlight
    picList.PSet (picList.ScaleWidth - 25, I + 1), vbButtonShadow
  Next I
  picButton_Paint 0
  picButton_Paint 1
End Sub

Private Sub picList_Resize()
  picButton(1).Left = picList.Width - 20
  picList_Paint
End Sub

Private Sub picToolbar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbLeftButton Then MousePointer = vbNoDrop
End Sub

Private Sub picToolbar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  MousePointer = vbDefault
End Sub

Private Sub picToolbar_Paint()
  Dim X As RECT
  ' Draw the edges
  With X
    .Left = 0
    .Right = picToolbar.ScaleWidth
    .Top = 0
    .Bottom = 2
    DrawEdge picToolbar.hdc, X, BDR_SUNKENOUTER, BF_TOP Or BF_BOTTOM

    .Top = 3
    .Bottom = 25

    .Left = Toolbar1.Left - 7
    .Right = .Left + 3
    DrawEdge picToolbar.hdc, X, BDR_RAISEDINNER, BF_RECT

    .Left = Toolbar1.Left - 4
    .Right = .Left + 3
    DrawEdge picToolbar.hdc, X, BDR_RAISEDINNER, BF_RECT
  End With
End Sub

Private Sub Splitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbLeftButton Then
    Splitter.BackColor = RGB(176, 176, 176)
    mSplitterPos = CLng(Splitter.Left + X)
  Else
    If mSplitterPos <> &H7FFFFFFF Then
      Splitter_MouseUp Button, Shift, X, Y
    End If
    mSplitterPos = &H7FFFFFFF
  End If
End Sub

Private Sub Splitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim Temp As Long
  If mSplitterPos <> &H7FFFFFFF Then
    If CLng(Splitter.Left + X) <> mSplitterPos Then
      
      Temp = Splitter.Left
      If picList.Left = 0 Then
        If Splitter.Left + X > (ScaleWidth / 3) Then
          Temp = ScaleWidth / 3 - X
        ElseIf Splitter.Left + X < 150 Then
          Temp = 150 - X
        End If
      Else
        If Splitter.Left + X < (ScaleWidth * 2 / 3) Then
          Temp = ScaleWidth * 2 / 3 - X
        ElseIf Splitter.Left + X > ScaleWidth - 150 Then
          Temp = ScaleWidth - 150 - X
        End If
      End If
      
      Splitter.Move Temp + X
      mSplitterPos = Splitter.Left + CLng(X)
    End If
  End If
End Sub

Private Sub Splitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If mSplitterPos <> &H7FFFFFFF Then
    
    If CLng(X) <> mSplitterPos Then Splitter.Move Splitter.Left + X
    
    Splitter.BackColor = vbButtonFace
    If picList.Left = 0 Then
      If Splitter.Left + X > (ScaleWidth / 3) Then
        Splitter.Left = ScaleWidth / 3 - X
      ElseIf Splitter.Left < 150 Then
        Splitter.Left = 150 - X
      End If
    Else
      If Splitter.Left + X < (ScaleWidth * 2 / 3) Then
        Splitter.Left = ScaleWidth * 2 / 3 - X
      ElseIf Splitter.Left + X > ScaleWidth - 150 Then
        Splitter.Left = ScaleWidth - 150 - X
      End If
    End If
     
    mSplitterPos = Splitter.Left + X
    Form_Resize
    mSplitterPos = &H7FFFFFFF
  End If
End Sub

Private Sub TimerAutoSave_Timer()
  Dim X As Long
  If Options.AutoSave > 0 And Scenery.Changed Then
    X = Timer
    ' Midnight test
    If X - NextAutoSave > 32768 Then Exit Sub
    If X > NextAutoSave Then
      If FileExists(AutoSaveFiles(AutoSavePointer)) Then Kill AutoSaveFiles(AutoSavePointer)
      AutoSaveFiles(AutoSavePointer) = AddDir(GetTempPathName(), "FSSC " & GetFileTitle(Left$(Scenery.File, Len(Scenery.File) - 4)) & " " & Format(Date, "mmddyy") & " " & Format(Time, "HhNnSs") & ".scn")
      Scenery.SaveFile AutoSaveFiles(AutoSavePointer), True
      AutoSavePointer = (AutoSavePointer + 1) Mod 4
    End If
  End If
End Sub

Private Sub TimerMouseMove_Timer()
  Const Margin = 15
  
  Dim lp As POINTAPI, Moved As Boolean
  
  GetCursorPos lp
  ScreenToClient picEdit.hwnd, lp
  mDontUpdate = True
  If lp.X < Margin And HScroll.Value <> HScroll.Min Then
    SendMessageLong HScroll.hwnd, WM_HSCROLL, SB_LINELEFT, HScroll.hwnd
    Moved = True
  End If
  If lp.Y < Margin And VScroll.Value <> VScroll.Min Then
    SendMessageLong VScroll.hwnd, WM_VSCROLL, SB_LINEUP, VScroll.hwnd
    Moved = True
  End If
  If lp.X > picEditor.myWidth - Margin - 2 And HScroll.Value <> HScroll.Max Then
    SendMessageLong HScroll.hwnd, WM_HSCROLL, SB_LINERIGHT, HScroll.hwnd
    Moved = True
  End If
  If lp.Y > picEditor.myHeight - Margin - 2 And VScroll.Value <> VScroll.Max Then
    SendMessageLong VScroll.hwnd, WM_VSCROLL, SB_LINEDOWN, VScroll.hwnd
    Moved = True
  End If
  mDontUpdate = False

  If Moved Then
    picEdit_Resize
    If Dragging Then
      ' Copy the new editor to the drag buffer
      picEditor.CopyTo picDrag.hdc
    ElseIf Gridding Then
      ' Erase the old grid
      picEdit_MouseMove -2, 0, 0, 0
      ' Copy the new one to the screen
      picEditor.CopyTo picEdit.hdc
    End If
    ' By scrolling, in effect, the mouse has moved
    ' Simulate the mousemove call so that
    ' dragged items get drawn, and the grid gets drawn
    picEdit_MouseMove vbLeftButton, 0, (lp.X), (lp.Y)
    TimerMouseMove.Interval = 200
  Else
    TimerMouseMove.Interval = 1000
  End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
  Dim I As Integer
  Select Case Button.Index
    Case 1
      mnuNew_Click
    Case 2
      mnuOpen_Click
    Case 3
      mnuSave_Click
    Case 4
      mnuExport_Click
    Case 6
'      mnuUndo_Click
    Case 7
      mnuCut_Click
    Case 8
      mnuCopy_Click
    Case 9
      mnuPaste_Click
    Case 10
      mnuDelete_Click
    Case 12
      If mnuMainProperties.Visible Then
        mnuMainProperties_Click
      End If
    Case 13
      mnuObjects_Click
    Case 14
      With Options
        .FillPolygons = (Toolbar1.Buttons(14).Value = tbrPressed)
        .FillObjects = .FillPolygons
        .ThickLines = .FillPolygons
      End With
      For I = 1 To Scenery.Count
        Scenery(I).UpdateObject
      Next I
      Scenery.Draw
    Case 16
      ' Rotate anti-clockwise
      For I = 1 To UBound(mFocusIndices)
        With Scenery(mFocusIndices(I))
          .Rotation = EnsureRotation(.Rotation - 1)
          .UpdateObject
        End With
      Next I
      Scenery.Changed = True
      Scenery.Draw
    Case 17
      ' Rotate clockwise
      For I = 1 To UBound(mFocusIndices)
        With Scenery(mFocusIndices(I))
          .Rotation = EnsureRotation(.Rotation + 1)
          .UpdateObject
        End With
      Next I
      Scenery.Changed = True
      Scenery.Draw
    Case 19
      mnuOrder_Click 0
    Case 20
      mnuOrder_Click 1
    Case 22
      mnuZoom_Click 0
    Case 23
      mnuZoom_Click 1
    Case 24
      mnuZoomSpecify_Click
    Case 26
      mnuHelp_Click 0
  End Select
End Sub

Private Sub VScroll_Change()
  VScroll_Scroll
End Sub

Private Sub VScroll_Scroll()
  mMapCenter.Y = -CLng(VScroll.Value) * 10
  If Not mDontUpdate Then picEdit_Resize
End Sub
