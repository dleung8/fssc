VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmRunway 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   ClipControls    =   0   'False
   Icon            =   "Runway.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame TabFrame 
      BorderStyle     =   0  'None
      Height          =   5535
      Index           =   4
      Left            =   360
      TabIndex        =   45
      Top             =   600
      Width           =   4440
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   16
         Left            =   2040
         TabIndex        =   71
         Tag             =   "1184"
         Top             =   5220
         WhatsThisHelpID =   1184
         Width           =   1095
      End
      Begin VB.CheckBox chkDistance 
         Height          =   195
         Left            =   0
         TabIndex        =   70
         Tag             =   "1184"
         Top             =   5250
         WhatsThisHelpID =   1184
         Width           =   2055
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   14
         Left            =   2040
         TabIndex        =   67
         Tag             =   "1182"
         Top             =   4440
         WhatsThisHelpID =   1180
         Width           =   1095
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   13
         Left            =   2040
         TabIndex        =   65
         Tag             =   "1181"
         Top             =   4050
         WhatsThisHelpID =   1180
         Width           =   1095
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   12
         Left            =   2040
         TabIndex        =   63
         Tag             =   "1180"
         Top             =   3660
         WhatsThisHelpID =   1180
         Width           =   1095
      End
      Begin VB.CheckBox chkVASI 
         Height          =   195
         Left            =   240
         TabIndex        =   61
         Tag             =   "1179"
         Top             =   3390
         WhatsThisHelpID =   1179
         Width           =   3735
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   15
         Left            =   2040
         TabIndex        =   69
         Tag             =   "1183"
         Top             =   4830
         WhatsThisHelpID =   1183
         Width           =   1095
      End
      Begin VB.ComboBox Cmbs 
         Height          =   315
         Index           =   8
         ItemData        =   "Runway.frx":000C
         Left            =   2040
         List            =   "Runway.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   2925
         WhatsThisHelpID =   1177
         Width           =   2295
      End
      Begin VB.ComboBox Cmbs 
         Height          =   315
         Index           =   7
         ItemData        =   "Runway.frx":0010
         Left            =   2040
         List            =   "Runway.frx":0012
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   2535
         WhatsThisHelpID =   1177
         Width           =   2295
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   11
         Left            =   2040
         TabIndex        =   56
         Tag             =   "1176"
         Top             =   2160
         WhatsThisHelpID =   1176
         Width           =   1095
      End
      Begin VB.ComboBox Cmbs 
         Height          =   315
         Index           =   6
         ItemData        =   "Runway.frx":0014
         Left            =   2040
         List            =   "Runway.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   1755
         WhatsThisHelpID =   1175
         Width           =   2295
      End
      Begin VB.CheckBox chkThreshold 
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   52
         Tag             =   "1174"
         Top             =   1440
         WhatsThisHelpID =   1174
         Width           =   3975
      End
      Begin VB.CheckBox chkThreshold 
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   51
         Tag             =   "1173"
         Top             =   1200
         WhatsThisHelpID =   1173
         Width           =   3975
      End
      Begin VB.CheckBox chkThreshold 
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   50
         Tag             =   "1172"
         Top             =   930
         WhatsThisHelpID =   1172
         Width           =   4215
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   10
         Left            =   2040
         TabIndex        =   49
         Tag             =   "1171"
         Top             =   510
         WhatsThisHelpID =   1171
         Width           =   1095
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   9
         Left            =   2040
         TabIndex        =   47
         Tag             =   "1170"
         Top             =   120
         WhatsThisHelpID =   1170
         Width           =   1095
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   29
         Left            =   480
         TabIndex        =   66
         Tag             =   "1182"
         Top             =   4470
         WhatsThisHelpID =   1180
         Width           =   1515
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   28
         Left            =   480
         TabIndex        =   64
         Tag             =   "1181"
         Top             =   4080
         WhatsThisHelpID =   1180
         Width           =   1515
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   27
         Left            =   480
         TabIndex        =   62
         Tag             =   "1180"
         Top             =   3690
         WhatsThisHelpID =   1180
         Width           =   1515
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   30
         Left            =   0
         TabIndex        =   68
         Tag             =   "1183"
         Top             =   4860
         WhatsThisHelpID =   1183
         Width           =   1995
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   26
         Left            =   240
         TabIndex        =   59
         Tag             =   "1178"
         Top             =   2970
         WhatsThisHelpID =   1177
         Width           =   1755
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   25
         Left            =   240
         TabIndex        =   57
         Tag             =   "1177"
         Top             =   2580
         WhatsThisHelpID =   1177
         Width           =   1755
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   24
         Left            =   240
         TabIndex        =   55
         Tag             =   "1176"
         Top             =   2190
         WhatsThisHelpID =   1176
         Width           =   1755
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   23
         Left            =   240
         TabIndex        =   53
         Tag             =   "1175"
         Top             =   1800
         WhatsThisHelpID =   1175
         Width           =   1755
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   22
         Left            =   0
         TabIndex        =   48
         Tag             =   "1171"
         Top             =   540
         WhatsThisHelpID =   1171
         Width           =   1995
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   21
         Left            =   0
         TabIndex        =   46
         Tag             =   "1170"
         Top             =   150
         WhatsThisHelpID =   1170
         Width           =   1995
      End
   End
   Begin VB.Frame TabFrame 
      BorderStyle     =   0  'None
      Height          =   5535
      Index           =   1
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   4440
      Begin VB.CheckBox chkVisible 
         Height          =   195
         Left            =   0
         TabIndex        =   25
         Tag             =   "1155"
         Top             =   5160
         WhatsThisHelpID =   1155
         Width           =   4215
      End
      Begin VB.CheckBox chkAFD 
         Height          =   195
         Left            =   0
         TabIndex        =   24
         Tag             =   "1154"
         Top             =   4800
         WhatsThisHelpID =   1154
         Width           =   4215
      End
      Begin VB.ComboBox Cmbs 
         Height          =   315
         Index           =   0
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   4245
         WhatsThisHelpID =   1153
         Width           =   2895
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   8
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   21
         Tag             =   "1152"
         Top             =   3870
         WhatsThisHelpID =   1152
         Width           =   1095
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   7
         Left            =   1320
         TabIndex        =   19
         Tag             =   "1151"
         Top             =   3480
         WhatsThisHelpID =   1150
         Width           =   1095
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   6
         Left            =   1320
         TabIndex        =   17
         Tag             =   "1150"
         Top             =   3090
         WhatsThisHelpID =   1150
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
         Index           =   11
         Left            =   0
         TabIndex        =   22
         Tag             =   "1153"
         Top             =   4290
         WhatsThisHelpID =   1153
         Width           =   1185
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   10
         Left            =   0
         TabIndex        =   20
         Tag             =   "1152"
         Top             =   3900
         WhatsThisHelpID =   1152
         Width           =   1185
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   9
         Left            =   0
         TabIndex        =   18
         Tag             =   "1151"
         Top             =   3510
         WhatsThisHelpID =   1150
         Width           =   1185
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   8
         Left            =   0
         TabIndex        =   16
         Tag             =   "1150"
         Top             =   3120
         WhatsThisHelpID =   1150
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
   Begin VB.Frame TabFrame 
      BorderStyle     =   0  'None
      Height          =   5535
      Index           =   5
      Left            =   360
      TabIndex        =   72
      Top             =   600
      Width           =   4455
      Begin VB.ComboBox Cmbs 
         Height          =   315
         Index           =   9
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   94
         Top             =   4545
         WhatsThisHelpID =   1611
         Width           =   2055
      End
      Begin VB.CheckBox chkILSFlags 
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   92
         Tag             =   "1625"
         Top             =   4260
         WhatsThisHelpID =   1625
         Width           =   3135
      End
      Begin VB.CheckBox chkILSFlags 
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   91
         Tag             =   "1624"
         Top             =   4020
         WhatsThisHelpID =   1624
         Width           =   3135
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   24
         Left            =   1800
         TabIndex        =   90
         Tag             =   "1612"
         Top             =   3600
         WhatsThisHelpID =   1612
         Width           =   1095
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   23
         Left            =   1800
         TabIndex        =   88
         Tag             =   "1606"
         Top             =   3210
         WhatsThisHelpID =   1606
         Width           =   1095
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   22
         Left            =   1800
         TabIndex        =   86
         Tag             =   "1604"
         Top             =   2820
         WhatsThisHelpID =   1603
         Width           =   1095
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   21
         Left            =   1800
         MaxLength       =   5
         TabIndex        =   84
         Tag             =   "1601"
         Top             =   2430
         WhatsThisHelpID =   1601
         Width           =   1095
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   20
         Left            =   1800
         TabIndex        =   82
         Tag             =   "1600"
         Top             =   2040
         WhatsThisHelpID =   1600
         Width           =   2535
      End
      Begin VB.CheckBox chkILS 
         Height          =   195
         Left            =   0
         TabIndex        =   80
         Tag             =   "1194"
         Top             =   1710
         WhatsThisHelpID =   1194
         Width           =   4215
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   19
         Left            =   2280
         TabIndex        =   79
         Tag             =   "1193"
         Top             =   1260
         WhatsThisHelpID =   1190
         Width           =   1095
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   18
         Left            =   2280
         TabIndex        =   77
         Tag             =   "1192"
         Top             =   870
         WhatsThisHelpID =   1190
         Width           =   1095
      End
      Begin VB.TextBox Txts 
         Height          =   285
         Index           =   17
         Left            =   2280
         TabIndex        =   75
         Tag             =   "1191"
         Top             =   480
         WhatsThisHelpID =   1190
         Width           =   1095
      End
      Begin VB.CheckBox chkMarkers 
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   78
         Tag             =   "1193"
         Top             =   1290
         WhatsThisHelpID =   1190
         Width           =   1995
      End
      Begin VB.CheckBox chkMarkers 
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   76
         Tag             =   "1192"
         Top             =   900
         WhatsThisHelpID =   1190
         Width           =   1995
      End
      Begin VB.CheckBox chkMarkers 
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   74
         Tag             =   "1191"
         Top             =   510
         WhatsThisHelpID =   1190
         Width           =   1995
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   40
         Left            =   240
         TabIndex        =   93
         Tag             =   "1611"
         Top             =   4590
         WhatsThisHelpID =   1611
         Width           =   1515
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   39
         Left            =   240
         TabIndex        =   89
         Tag             =   "1612"
         Top             =   3630
         WhatsThisHelpID =   1612
         Width           =   1515
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   38
         Left            =   240
         TabIndex        =   87
         Tag             =   "1606"
         Top             =   3240
         WhatsThisHelpID =   1606
         Width           =   1515
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   37
         Left            =   240
         TabIndex        =   85
         Tag             =   "1604"
         Top             =   2850
         WhatsThisHelpID =   1603
         Width           =   1515
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   36
         Left            =   240
         TabIndex        =   83
         Tag             =   "1601"
         Top             =   2460
         WhatsThisHelpID =   1601
         Width           =   1515
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   35
         Left            =   240
         TabIndex        =   81
         Tag             =   "1600"
         Top             =   2070
         WhatsThisHelpID =   1600
         Width           =   1515
      End
      Begin VB.Label lbls 
         Height          =   255
         Index           =   31
         Left            =   0
         TabIndex        =   73
         Tag             =   "1190"
         Top             =   150
         WhatsThisHelpID =   1190
         Width           =   2955
      End
   End
   Begin VB.Frame TabFrame 
      BorderStyle     =   0  'None
      Height          =   5535
      Index           =   3
      Left            =   360
      TabIndex        =   34
      Top             =   600
      Width           =   4440
      Begin VB.ComboBox Cmbs 
         Height          =   315
         Index           =   5
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   5025
         WhatsThisHelpID =   1165
         Width           =   2175
      End
      Begin VB.ComboBox Cmbs 
         Height          =   315
         Index           =   4
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   2745
         WhatsThisHelpID =   1166
         Width           =   2175
      End
      Begin VB.ComboBox Cmbs 
         Height          =   315
         Index           =   3
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   630
         WhatsThisHelpID =   1165
         Width           =   2175
      End
      Begin VB.Label lblPreview 
         Height          =   855
         Left            =   1680
         TabIndex        =   44
         Tag             =   "3023"
         Top             =   1680
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lbls 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   20
         Left            =   2880
         TabIndex        =   43
         Top             =   4440
         Width           =   735
      End
      Begin VB.Label lbls 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   19
         Left            =   2880
         TabIndex        =   42
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lbls 
         Alignment       =   2  'Center
         Height          =   195
         Index           =   18
         Left            =   2640
         TabIndex        =   41
         Tag             =   "1168"
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label lbls 
         Height          =   195
         Index           =   17
         Left            =   0
         TabIndex        =   39
         Tag             =   "1167"
         Top             =   4800
         WhatsThisHelpID =   1165
         Width           =   2025
      End
      Begin VB.Label lbls 
         Height          =   195
         Index           =   16
         Left            =   0
         TabIndex        =   37
         Tag             =   "1166"
         Top             =   2520
         WhatsThisHelpID =   1166
         Width           =   2025
      End
      Begin VB.Label lbls 
         Height          =   195
         Index           =   15
         Left            =   0
         TabIndex        =   35
         Tag             =   "1165"
         Top             =   405
         WhatsThisHelpID =   1165
         Width           =   2025
      End
      Begin VB.Image imgRunway 
         Height          =   735
         Index           =   5
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   4680
         Width           =   495
      End
      Begin VB.Image imgRunway 
         Height          =   3615
         Index           =   4
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   495
      End
      Begin VB.Image imgRunway 
         Height          =   735
         Index           =   3
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame TabFrame 
      BorderStyle     =   0  'None
      Height          =   5535
      Index           =   2
      Left            =   360
      TabIndex        =   26
      Top             =   600
      Width           =   4440
      Begin VB.CheckBox chkRed 
         Height          =   195
         Left            =   0
         TabIndex        =   33
         Tag             =   "1163"
         Top             =   5160
         WhatsThisHelpID =   1163
         Width           =   3735
      End
      Begin VB.ComboBox Cmbs 
         Height          =   315
         Index           =   2
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   4710
         WhatsThisHelpID =   1161
         Width           =   2895
      End
      Begin VB.ComboBox Cmbs 
         Height          =   315
         Index           =   1
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   4320
         WhatsThisHelpID =   1161
         Width           =   2895
      End
      Begin VB.ListBox lstMarkers 
         Height          =   3660
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   28
         Top             =   360
         WhatsThisHelpID =   1160
         Width           =   4095
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   13
         Left            =   0
         TabIndex        =   31
         Tag             =   "1162"
         Top             =   4755
         WhatsThisHelpID =   1161
         Width           =   1185
      End
      Begin VB.Label lbls 
         Height          =   390
         Index           =   12
         Left            =   0
         TabIndex        =   29
         Tag             =   "1161"
         Top             =   4365
         WhatsThisHelpID =   1161
         Width           =   1185
      End
      Begin VB.Label lbls 
         Height          =   255
         Index           =   14
         Left            =   0
         TabIndex        =   27
         Tag             =   "1160"
         Top             =   120
         WhatsThisHelpID =   1160
         Width           =   2505
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
      Left            =   3720
      TabIndex        =   96
      Tag             =   "1031"
      Top             =   6360
      WhatsThisHelpID =   1031
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   95
      Tag             =   "1030"
      Top             =   6360
      WhatsThisHelpID =   1030
      Width           =   1215
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   6135
      Left            =   120
      TabIndex        =   97
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   10821
      MultiRow        =   -1  'True
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   7
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "A"
            Object.Tag             =   "1280"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "B"
            Object.Tag             =   "1281"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "C"
            Object.Tag             =   "1282"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "D"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "D1"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "E"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "E1"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmRunway"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mValueCache(8) As Single

Private mTxtChanged As Boolean
Private mChanged As Boolean
Private mDontUpdate As Boolean

Private NearOpt As clsRunwayOpt
Private FarOpt As clsRunwayOpt

Private CurOpt As clsRunwayOpt

Private Sub AutomaticCalculate()
  Dim HDistance As Single, HConversion As Single, _
      VDistance As Single, VConversion As Single, _
      TDistance As Single, TConversion As Single
      
  UserToMeter Txts(7).Text, 0, HConversion, 0, 0
  If HConversion > 0 Then
    HDistance = ValEx(Txts(7).Text) * HConversion
    UserToMeter Txts(6).Text, 0, VConversion, 0, 0
    If VConversion > 0 Then
      VDistance = ValEx(Txts(6).Text) * VConversion
      UserToMeter Txts(9).Text, 0, TConversion, 0, 0
      If TConversion > 0 Then
        TDistance = ValEx(Txts(9).Text) * TConversion
      
        Dim Num1 As Integer, Num2 As Integer
        Select Case Cmbs(7).ListIndex
          Case 0: Num1 = 0
          Case 7, 8, 9, 10: Num1 = 1
          Case 1, 3, 5: Num1 = 2
          Case 2, 4, 6: Num1 = 3
        End Select
        
        If Options.FSVersion >= Version_FS2K Then
          Select Case Cmbs(8).ListIndex
            Case 0: Num2 = 0
            Case 7, 8, 9, 10: Num2 = 1
            Case 1, 3, 5: Num2 = 2
            Case 2, 4, 6: Num2 = 3
          End Select
        Else
          Select Case Cmbs(8).ListIndex
            Case 0: Num2 = 0
            Case 1: Num2 = 1
          End Select
        End If
            
        If Num2 > Num1 Then
          Num1 = Num2
        End If
  
        Select Case Num1
          Case 0
            CurOpt.VDistance = 0
            CurOpt.BarSpacing = 0
          Case 1
            CurOpt.VDistance = VDistance / 2 - TDistance - 250
            CurOpt.BarSpacing = 0
          Case 2
            CurOpt.VDistance = VDistance / 2 - TDistance - 250
            CurOpt.BarSpacing = 215
          Case 3
            CurOpt.VDistance = VDistance / 2 - TDistance - 300
            CurOpt.BarSpacing = 130
        End Select
  
        CurOpt.HDistance = HDistance / 2 + 15
        Txts(12).Text = MeterToUser(CurOpt.HDistance)
        Txts(13).Text = MeterToUser(CurOpt.VDistance)
        Txts(14).Text = MeterToUser(CurOpt.BarSpacing)
      End If
    End If
  End If
End Sub

Public Function EditData(Data As clsRunway) As Boolean
  Dim I As Integer
  
  Load frmRunway

  With Data
    Txts(0).Tag = .Caption(True)
    Txts(0).Text = .Name
    Txts_Change 0 ' (Changes the caption)
    Txts(1).Text = MeterToUser(.X)
    Txts(2).Text = MeterToUser(.Y)
    Txts_Validate 1, False ' (Updates Latitude, Longitude)
    chkLocked.Value = -.Locked
    Txts(5).Text = GeographicToUser(.Rotation)
    Txts(6).Text = MeterToUser(.Length)
    Txts(7).Text = MeterToUser(.Width)
    Txts(8).Text = .ID
    Txts_Validate 8, False
    Cmbs(0).ListIndex = -.RunwayPos
    chkAFD.Value = -.AFDEntry
    chkVisible.Value = -.RunwayVisible
    
    For I = 0 To 6
      lstMarkers.Selected(I) = (.Markers And 2 ^ I) > 0
    Next I
    Cmbs(1).ListIndex = .EdgeLights
    Cmbs(2).ListIndex = .CenterLights And 3
    Cmbs(4).ListIndex = RunwayNumToListIndex(.Surface)
    
    If Options.FSVersion >= Version_FS2K Then
      lstMarkers.Selected(7) = (.Markers And 2 ^ 7) > 0
      For I = 11 To 15
        lstMarkers.Selected(I - 3) = (.Markers2 And 2 ^ (I - 8)) > 0
      Next I
      chkRed.Value = -((.CenterLights And 4) > 0)
      Cmbs(3).ListIndex = RunwayNumToListIndex(.Far.ExtSurface)
      Cmbs(5).ListIndex = RunwayNumToListIndex(.Near.ExtSurface)
    Else
      Cmbs(3).ListIndex = 0
      Cmbs(5).ListIndex = 0
    End If
    lstMarkers.ListIndex = 0
    
    .Near.CopyTo NearOpt
    .Far.CopyTo FarOpt

    mChanged = False
    If TabValue > 4 Then TabValue = 0
    Set TabStrip1.SelectedItem = TabStrip1.Tabs(TabValue + 1)
    Show vbModal, Screen.ActiveForm
    If mChanged Then
      Scenery.AFDRefresh = Scenery.AFDRefresh Or _
             .X <> mValueCache(1) Or _
             .Y <> mValueCache(2) Or _
             .Rotation <> mValueCache(5) Or _
             .Length <> mValueCache(6) Or _
             .Width <> mValueCache(7) Or _
             .ID <> Txts(8).Text Or _
             .Surface <> RunwayListIndexToNum(Cmbs(4).ListIndex) Or _
             .Near.Compare(NearOpt) Or _
             .Far.Compare(FarOpt)

      .Name = Txts(0).Text
      .X = mValueCache(1)
      .Y = mValueCache(2)
      .Locked = -chkLocked.Value
      .Rotation = mValueCache(5)
      .Length = mValueCache(6)
      .Width = mValueCache(7)
      .ID = Txts(8).Text
      
      .RunwayPos = -Cmbs(0).ListIndex
      .AFDEntry = -chkAFD.Value
      .RunwayVisible = -chkVisible.Value
      .Markers = 0
      For I = 0 To 6
        .Markers = .Markers + -lstMarkers.Selected(I) * 2 ^ I
      Next I
      
      .EdgeLights = Cmbs(1).ListIndex
      .CenterLights = Cmbs(2).ListIndex + (chkRed.Value * 4)
      .Surface = RunwayListIndexToNum(Cmbs(4).ListIndex)

      NearOpt.CopyTo .Near
      FarOpt.CopyTo .Far

      If Options.FSVersion >= Version_FS2K Then
        .Markers = .Markers + -lstMarkers.Selected(7) * 2 ^ 7
        .Markers2 = 0
        For I = 11 To 15
          .Markers2 = .Markers2 + -lstMarkers.Selected(I - 3) * 2 ^ (I - 8)
        Next I
        .Far.ExtSurface = RunwayListIndexToNum(Cmbs(3).ListIndex)
        .Near.ExtSurface = RunwayListIndexToNum(Cmbs(5).ListIndex)
      End If
          
      EditData = True
    End If
  End With

  Unload frmRunway
End Function

Private Function RunwayListIndexToNum(ByVal Num As Byte) As Long
  Select Case Options.FSVersion
    Case Is >= Version_FS2K2
      If Num = 33 Then
        RunwayListIndexToNum = 99
      ElseIf Num >= 28 Then
        RunwayListIndexToNum = Num + 36
      Else
        RunwayListIndexToNum = Num
      End If
    Case Version_CFS2
      If Num = 10 Then
        RunwayListIndexToNum = 99
      ElseIf Num >= 5 Then
        RunwayListIndexToNum = Num + 59
      Else
        RunwayListIndexToNum = Num
      End If
    Case Version_FS2K
      If Num = 10 Then
        RunwayListIndexToNum = 99
      Else
        RunwayListIndexToNum = Num
      End If
    Case Else
      If Num = 4 Then
        RunwayListIndexToNum = 99
      Else
        RunwayListIndexToNum = Num
      End If
  End Select
End Function

Private Function RunwayNumToListIndex(ByVal Surface As Byte) As Long
  Select Case Options.FSVersion
    Case Is >= Version_FS2K2
      If Surface = 99 Then
        RunwayNumToListIndex = 33
      ElseIf Surface >= 64 Then
        RunwayNumToListIndex = Surface - 36
      Else
        RunwayNumToListIndex = Surface
      End If
    Case Version_CFS2
      If Surface = 99 Then
        RunwayNumToListIndex = 10
      ElseIf Surface >= 64 Then
        RunwayNumToListIndex = Surface - 59
      Else
        RunwayNumToListIndex = Surface
      End If
    Case Version_FS2K
      If Surface = 99 Then
        RunwayNumToListIndex = 10
      Else
        RunwayNumToListIndex = Surface
      End If
    Case Else
      If Surface = 99 Then
        RunwayNumToListIndex = 4
      Else
        RunwayNumToListIndex = Surface
      End If
  End Select
End Function

Private Sub UpdateRunway()
  Dim RunwayStr As String, RunwayRev As String
  RunwayStr = Txts(8).Text
  RunwayRev = ReverseRunway(RunwayStr)
  If Options.FSVersion >= Version_FS2K Then
    With lstMarkers
      .List(8) = Lang.ResolveString(RES_Rwy_OneSided, RunwayRev)
      .List(9) = Lang.ResolveString(RES_Rwy_Closed, RunwayStr)
      .List(10) = Lang.ResolveString(RES_Rwy_Closed, RunwayRev)
      .List(11) = Lang.ResolveString(RES_Rwy_STOL, RunwayStr)
      .List(12) = Lang.ResolveString(RES_Rwy_STOL, RunwayRev)
    End With
  End If
  
  With TabStrip1
    .Tabs(4) = Lang.ResolveString(RES_Rwy_Tab4, RunwayStr)
    .Tabs(5) = Lang.ResolveString(RES_Rwy_Tab4, RunwayRev)
    .Tabs(6) = Lang.ResolveString(RES_Rwy_Tab6, RunwayStr)
    .Tabs(7) = Lang.ResolveString(RES_Rwy_Tab6, RunwayRev)
  End With
  
  lbls(19).Caption = RunwayRev
  lbls(20).Caption = RunwayStr
End Sub

Private Sub chkDistance_Click()
  SetEnabled Txts(16), chkDistance.Value = vbChecked
  If Not chkDistance.Value = vbChecked Then
    CurOpt.SignHSpacing = 0
  End If
End Sub

Private Sub chkILS_Click()
  Dim Value As Boolean
  Value = -chkILS.Value
  SetEnabled Txts(20), Value
  SetEnabled Txts(21), Value
  SetEnabled Txts(22), Value
  SetEnabled Txts(23), Value
  SetEnabled Txts(24), Value
  chkILSFlags(0).Enabled = Value
  chkILSFlags(1).Enabled = Value
  Cmbs(9).BackColor = IIf(Value, vbWindowBackground, vbButtonFace)
  Cmbs(9).Locked = Not Value
  CurOpt.ILSEnabled = Value
End Sub

Private Sub chkILSFlags_Click(Index As Integer)
  If mDontUpdate Then Exit Sub
  CurOpt.ILSFlags = -(Cmbs(9).ListIndex <> 0) * 1 + -(Cmbs(9).ListIndex = 1) * 16 + chkILSFlags(0).Value * 128 + chkILSFlags(1).Value * 64
End Sub

Private Sub chkLocked_Click()
  Dim Value As Boolean
  Value = Not -chkLocked.Value
  SetEnabled Txts(1), Value
  SetEnabled Txts(2), Value
  SetEnabled Txts(3), Value
  SetEnabled Txts(4), Value
End Sub

Private Sub chkMarkers_Click(Index As Integer)
  Dim Value As Boolean, TxtIndex As Integer, Temp As Single
  Value = (chkMarkers(Index).Value = vbChecked)
  TxtIndex = Index + 17
  SetEnabled Txts(Index + 17), Value
  If Value Then
    Select Case Index
      Case 0
        Txts(TxtIndex) = NauticalToUser(0.1, "0.0")
        Temp = 0.1
      Case 1
        Txts(TxtIndex) = NauticalToUser(0.5, "0.0")
        Temp = 0.5
      Case 2
        Txts(TxtIndex) = NauticalToUser(5, "0.0")
        Temp = 5
    End Select
  Else
    Txts(TxtIndex) = NauticalToUser(0, "0.0")
    Temp = 0
  End If
  
  If mDontUpdate Then Exit Sub
  
  Select Case Index
    Case 0: CurOpt.IM = Temp
    Case 1: CurOpt.MM = Temp
    Case 2: CurOpt.OM = Temp
  End Select
End Sub

Private Sub chkThreshold_Click(Index As Integer)
  Dim Value As Boolean, I As Integer
  Value = (chkThreshold(0).Value = vbChecked)
  chkThreshold(1).Enabled = Value
  chkThreshold(2).Enabled = Value
  SetEnabled Txts(11), Value
  For I = 6 To 8
    Cmbs(I).BackColor = IIf(Value, vbWindowBackground, vbButtonFace)
    Cmbs(I).Locked = Not Value
  Next I
  chkVASI.Enabled = Value
  
  Value = Value And Not (chkVASI.Value = vbChecked)
  SetEnabled Txts(12), Value
  SetEnabled Txts(13), Value
  SetEnabled Txts(14), Value
  
  If Not mDontUpdate Then
    CurOpt.ThrLights = chkThreshold(0).Value * 1 + chkThreshold(1).Value * 4 + chkThreshold(2).Value * 64
  End If
End Sub

Private Sub chkVASI_Click()
  Dim Value As Boolean
  Value = Not (chkVASI.Value = vbChecked)
  SetEnabled Txts(12), Value
  SetEnabled Txts(13), Value
  SetEnabled Txts(14), Value
  If Not mDontUpdate Then
    CurOpt.Automatic = (chkVASI.Value = vbChecked)
    If chkVASI.Value = vbChecked Then
      AutomaticCalculate
    Else
    End If
  End If
End Sub

Private Sub Cmbs_Click(Index As Integer)
  Dim Result As String, Surface As Integer, Num As Integer
  If mDontUpdate Then Exit Sub
  Select Case Index
    Case 3, 4, 5
      Surface = Cmbs(Index).ListIndex
      If Index <> 4 Then
        If Surface = Cmbs(Index).ListCount - 1 Or Options.FSVersion < Version_FS2K Then Surface = Cmbs(4).ListIndex
      End If
      
      Num = RunwayListIndexToNum(Surface)
      If Num >= 10 Then
        LoadTexture AddDir(Options.FSPath, "Texture\Runway") & Format$(Num, "00") & ".bmp", Result
      Else
        LoadTexture AddDir(Options.FSPath, "Texture\Runway") & Format$(Num, "00") & ".r8", Result
      End If
      
      lbls(19).Visible = True
      lbls(20).Visible = True
      If FileExists(Result) Then
        imgRunway(Index).Picture = LoadPicture(Result)
        lblPreview.Visible = False
        Kill Result
      Else
        imgRunway(Index).Picture = LoadPicture()
        If Num = 99 Then
          lbls(19).Visible = False
          lbls(20).Visible = False
        Else
          lblPreview.Visible = True
        End If
      End If
      
      If Index = 4 Then
        If Options.FSVersion >= Version_FS2K Then
          If Cmbs(3).ListIndex = Cmbs(3).ListCount - 1 Then Cmbs_Click 3
          If Cmbs(5).ListIndex = Cmbs(5).ListCount - 1 Then Cmbs_Click 5
          If Surface >= 10 Then
            Cmbs(3).Locked = False
            Cmbs(5).Locked = False
            Cmbs(3).BackColor = vbWindowBackground
            Cmbs(5).BackColor = vbWindowBackground
          Else
            Cmbs(3).Locked = True
            Cmbs(5).Locked = True
            Cmbs(3).BackColor = vbButtonFace
            Cmbs(5).BackColor = vbButtonFace
          End If
        Else
          Cmbs_Click 3
          Cmbs_Click 5
        End If
      End If
    Case 6
      CurOpt.ApprLights = Cmbs(6).ListIndex
    Case 7
      CurOpt.VASILeft = Cmbs(7).ListIndex
      If chkVASI.Value = vbChecked Then
        AutomaticCalculate
      End If
    Case 8
      If Options.FSVersion >= Version_FS2K Then
        CurOpt.VASIRight = Cmbs(8).ListIndex
      Else
        CurOpt.PAPI = Cmbs(8).ListIndex
      End If
      If chkVASI.Value = vbChecked Then
        AutomaticCalculate
      End If
    Case 9
      chkILSFlags_Click 0 ' Save settings
  End Select
End Sub

Private Sub cmdCancel_Click()
  Hide
End Sub

Private Sub cmdOK_Click()
  Dim I As Integer, Msg As String, Cancel As Boolean

  ' Validate event not fired when Enter key pressed
  ' bug workaround
  If TypeOf ActiveControl Is TextBox Then
    Txts_Validate ActiveControl.Index, Cancel
    If Cancel Then Exit Sub
  End If

  For I = 1 To 8
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Handle Ctrl+ [Shift + ] {TAB}s
  If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
    Set TabStrip1.SelectedItem = TabStrip1.Tabs(TabStrip1.SelectedItem.Index Mod TabStrip1.Tabs.Count + 1)
  ElseIf Shift = (vbCtrlMask Or vbShiftMask) And KeyCode = vbKeyTab Then
    Set TabStrip1.SelectedItem = TabStrip1.Tabs((TabStrip1.SelectedItem.Index + TabStrip1.Tabs.Count - 2) Mod TabStrip1.Tabs.Count + 1)
  End If
End Sub

Private Sub Form_Load()
  Dim I As Integer, HeightDiff As Single
  
  DialogMenus Me

  Lang.AddItems Cmbs(0), RES_Rwy_Position, 2
  Lang.AddItems Cmbs(1), RES_Rwy_Intensity, 4
  Lang.AddItems Cmbs(2), RES_Rwy_Intensity, 4
  Lang.AddItems Cmbs(6), RES_Rwy_ApprLights, 11
  Lang.AddItems Cmbs(9), RES_Rdo_LocalizerPos, 3
  
  If Options.FSVersion >= Version_FS2K Then
    Lang.AddItems lstMarkers, RES_Rwy_Markers, 9
    For I = 1 To 4
      lstMarkers.AddItem ""
    Next I
    Lang.AddItems Cmbs(7), RES_Rwy_VASI, 11
    Lang.AddItems Cmbs(8), RES_Rwy_VASI, 11
    Select Case Options.FSVersion
      Case Is >= Version_FS2K2
        For I = 3 To 5
          Lang.AddItems Cmbs(I), RES_Rwy_Surface, 28
          Lang.AddItems Cmbs(I), RES_Rwy_Surface + 64, 5
        Next I
      Case Version_FS2K
        For I = 3 To 5
          Lang.AddItems Cmbs(I), RES_Rwy_Surface, 10
        Next I
      Case Version_CFS2
        For I = 3 To 5
          Lang.AddItems Cmbs(I), RES_Rwy_Surface, 5
          Lang.AddItems Cmbs(I), RES_Rwy_Surface + 64, 5
        Next I
    End Select
    Lang.AddItems Cmbs(3), RES_Rwy_Surface + 99, 1
    Lang.AddItems Cmbs(4), RES_Rwy_Surface + 98, 1
    Lang.AddItems Cmbs(5), RES_Rwy_Surface + 99, 1
  Else
    Lang.AddItems lstMarkers, RES_Rwy_Markers, 7
    chkRed.Visible = False
    Cmbs(3).AddItem Lang.GetString(RES_Rwy_Surface + 99)
    Lang.AddItems Cmbs(4), RES_Rwy_Surface, 4
    Cmbs(5).AddItem Lang.GetString(RES_Rwy_Surface + 99)
    Lang.AddItems Cmbs(7), RES_Rwy_VASI, 3
    chkThreshold(1).Top = chkThreshold(1).Top + chkThreshold(1).Height / 2
    chkILSFlags(0).Top = chkILSFlags(0).Top - chkILSFlags(0).Height * 1.5
    chkILSFlags(1).Top = chkILSFlags(1).Top - chkILSFlags(1).Height
    Txts(15).Top = Txts(15).Top + Txts(15).Height / 2
    lbls(30).Top = Txts(15).Top + 30
    chkThreshold(2).Visible = False
    
    lbls(25).Tag = RES_Rwy_VASIlbl
    lbls(26).Tag = RES_Rwy_PAPIlbl
    lbls(26).WhatsThisHelpID = RES_Rwy_PAPIlbl
    Cmbs(8).WhatsThisHelpID = RES_Rwy_PAPIlbl
    Lang.AddItems Cmbs(8), RES_Rwy_PAPI, 2
    
    lbls(39).Visible = False
    Txts(24).Visible = False
    chkDistance.Visible = False
    Txts(16).Visible = False
  End If
  
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
  TabStrip1.Height = TabStrip1.Height + HeightDiff
  Height = Height + HeightDiff
  
  CenterForm Me
  
  Set NearOpt = New clsRunwayOpt
  Set FarOpt = New clsRunwayOpt
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then Cancel = True: Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set NearOpt = Nothing
  Set FarOpt = Nothing
End Sub

Private Sub TabStrip1_BeforeClick(Cancel As Integer)
  Timer1.Enabled = True
End Sub

Private Sub TabStrip1_Click()
  ' Handle Tab clicks
  Dim SelItem As Integer, SelIndex As Integer, OldItem As Integer
  
  Dim TempInt As Integer
  
  SelIndex = TabStrip1.SelectedItem.Index
  SelItem = Asc(TabStrip1.SelectedItem.Key) - 64

  If TabStrip1.Tag <> "" Then
    If SelIndex <> TabStrip1.Tag Then
      OldItem = Asc(TabStrip1.Tabs(CInt(TabStrip1.Tag)).Key) - 64
      If OldItem <> SelItem Then
        With TabFrame(OldItem)
          .Visible = False
          .Enabled = False
        End With
      End If
    End If
  End If

  With TabFrame(SelItem)
    .Enabled = True
    .Visible = True
  End With

  If TabStrip1.Tag <> "" And TabStrip1.Visible Then
    If Not ActiveControl Is TabStrip1 Then
      If TabStrip1.Tag <> SelIndex Then
        Select Case SelItem
          Case 1: Txts(0).SetFocus
          Case 2: lstMarkers.SetFocus
          Case 3: Cmbs(4).SetFocus
          Case 4
            If OldItem = SelItem Then _
              Txts(10).SetFocus
              ' Need to ensure that lost_focus is called
              ' to save last item
            Txts(9).SetFocus
          Case 5
            If OldItem = SelItem Then _
              Txts(18).SetFocus
            chkMarkers(0).SetFocus
        End Select
        DoEvents ' Make LostFocus events execute so that
                 ' the previous value from the current
                 ' textbox is saved
      End If
    End If
  End If
  
  If Not mDontUpdate Then
    Select Case SelIndex
      Case 4, 5
        If SelIndex = 4 Then
          Set CurOpt = NearOpt
        Else
          Set CurOpt = FarOpt
        End If
        With CurOpt
          mDontUpdate = True
          Txts(9).Text = MeterToUser(.ThrLength)
          Txts(10).Text = MeterToUser(.ExtLength)
          chkThreshold(0).Value = -((.ThrLights And 1) > 0)
          chkThreshold(1).Value = -((.ThrLights And 4) > 0)
          Cmbs(6).ListIndex = .ApprLights
          Txts(11).Text = .NumStrobes
          Cmbs(7).ListIndex = .VASILeft
          chkVASI.Value = -.Automatic
          If Not .Automatic Then
            Txts(12).Text = MeterToUser(.HDistance)
            Txts(13).Text = MeterToUser(.VDistance)
            Txts(14).Text = MeterToUser(.BarSpacing)
          Else
            AutomaticCalculate
          End If
          Txts(15).Text = Append(.GlideSlope, RES_Unit_AbbrevDeg, "0.0")

          If Options.FSVersion >= Version_FS2K Then
            chkThreshold(2).Value = -((.ThrLights And 64) > 0)
            Cmbs(8).ListIndex = .VASIRight
            chkDistance.Value = -(.SignHSpacing > 0)
            Txts(16).Text = MeterToUser(.SignHSpacing)
            chkDistance_Click
          Else
            Cmbs(8).ListIndex = .PAPI
          End If

          chkThreshold_Click 0
          chkVASI_Click
          mDontUpdate = False
        End With
        Txts_GotFocus 9
      Case 6, 7
        If SelIndex = 6 Then
          Set CurOpt = NearOpt
        Else
          Set CurOpt = FarOpt
        End If
        With CurOpt
          mDontUpdate = True
          chkMarkers(0).Value = -(.IM > 0)
          chkMarkers(1).Value = -(.MM > 0)
          chkMarkers(2).Value = -(.OM > 0)
          chkMarkers_Click 0
          chkMarkers_Click 1
          chkMarkers_Click 2
          Txts(17).Text = NauticalToUser(.IM)
          Txts(18).Text = NauticalToUser(.MM)
          Txts(19).Text = NauticalToUser(.OM)
          chkILS.Value = -.ILSEnabled
          Txts(20).Text = .ILSName
          Txts(21).Text = .ILSID
          Txts(22).Text = Append(.ILSFrequency, RES_Unit_AbbrevMhz, "000.00")
          Txts(23).Text = NauticalToUser(.ILSRange)
          If Options.FSVersion >= Version_FS2K Then Txts(24).Text = Append(.ILSBeamWidth, RES_Unit_AbbrevDeg, "0.0")
          chkILSFlags(0).Value = -((.ILSFlags And 128) > 0)
          chkILSFlags(1).Value = -((.ILSFlags And 64) > 0)
          
          TempInt = -((.ILSFlags And 1) > 0)
          If TempInt > 0 Then TempInt = TempInt + -((.ILSFlags And 16) = 0)
          Cmbs(9).ListIndex = TempInt
          
          chkILS_Click
          mDontUpdate = False
        End With
        Txts_GotFocus 17
    End Select
  End If

  TabStrip1.Tag = SelIndex
End Sub

Private Sub Timer1_Timer()
  TabStrip1_Click
  Timer1.Enabled = False
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
  Select Case Index
    Case 1, 2, 5, 6, 7, 9, 10, 12, 13, 14, 15, 16, 17, 18, 19, 22, 23, 24
      SmartSelectText Txts(Index)
    Case Else
      If Index = 0 Then ReturnSymbol 0, 0, 2
      SelectText Txts(Index)
  End Select
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
  Select Case Index
    Case 0
      KeyAscii = ReturnSymbol(KeyAscii, 0, 1)
    Case 8, 21
      KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
  End Select
End Sub

Private Sub Txts_Validate(Index As Integer, Cancel As Boolean)
  Dim valX As Single, valY As Single, _
    Distance As Double, Angle As Single, _
    LevelIndex As Integer, Msg As String, _
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
    Case 8
      UpdateRunway
    Case 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24
      If Validate(Txts(Index), Msg, Value) Then
        Select Case Index
          Case 9
            CurOpt.ThrLength = Value
            If chkVASI.Value = vbChecked Then
              AutomaticCalculate
            End If
          Case 10
            CurOpt.ExtLength = Value
          Case 11
            CurOpt.NumStrobes = Value
          Case 12
            CurOpt.HDistance = Value
          Case 13
            CurOpt.VDistance = Value
          Case 14
            CurOpt.BarSpacing = Value
          Case 15
            CurOpt.GlideSlope = Value
          Case 16
            CurOpt.SignHSpacing = Value
          Case 17
            CurOpt.IM = Value
          Case 18
            CurOpt.MM = Value
          Case 19
            CurOpt.OM = Value
          Case 20
            CurOpt.ILSName = Txts(20).Text
          Case 21
            CurOpt.ILSID = Txts(21).Text
          Case 22
            CurOpt.ILSFrequency = Value
          Case 23
            CurOpt.ILSRange = Value
          Case 24
            CurOpt.ILSBeamWidth = Value
        End Select
      Else
        MsgBoxEx Me, Msg, vbCritical, RES_ERR_Bound
        SmartSelectText Txts(Index)
        mTxtChanged = True
        Cancel = True
      End If
  End Select
End Sub
