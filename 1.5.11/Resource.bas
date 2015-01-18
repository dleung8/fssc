Attribute VB_Name = "basResource"
Option Explicit

' Language
Public Const RES_LangName = 1000

' Titles
Public Const RES_Title = 1010
Public Const RES_LongTitle = 1011
Public Const RES_Tutorial = 1012
Public Const RES_Untitled = 1013

' File Names
Public Const RES_LangHelpName = 1020
Public Const RES_UntitledFile = 1021

' Common Labels
Public Const RES_LBL_X = 1043
Public Const RES_LBL_Y = 1044
Public Const RES_LBL_Latitude = 1046
Public Const RES_LBL_Longitude = 1047
Public Const RES_LBL_Rotation = 1048
Public Const RES_LBL_Complexity = 1049
Public Const RES_Tab_Properties = 1059

Public Const RES_Complexity1 = 1060

' Units of measure
Public Const RES_Unit_M = 1070
Public Const RES_Unit_Ft = 1071
Public Const RES_Unit_Nm = 1072
Public Const RES_Unit_Km = 1073
Public Const RES_Unit_Mi = 1074
Public Const RES_Unit_Deg = 1075
Public Const RES_Unit_Min = 1076
Public Const RES_Unit_Sec = 1077
Public Const RES_Unit_Mhz = 1078
Public Const RES_Unit_Khz = 1079
Public Const RES_Unit_AbbrevM = 1080
Public Const RES_Unit_AbbrevFt = 1081
Public Const RES_Unit_AbbrevNm = 1082
Public Const RES_Unit_AbbrevKm = 1083
Public Const RES_Unit_AbbrevMi = 1084
Public Const RES_Unit_AbbrevDeg = 1085
Public Const RES_Unit_AbbrevGeo = 1086
Public Const RES_Unit_AbbrevMag = 1087
Public Const RES_Unit_AbbrevMhz = 1088
Public Const RES_Unit_AbbrevKhz = 1089
Public Const RES_Unit_PerPixel = 1090
Public Const RES_Unit_PerUnit = 1091
'Public Const RES_Unit_ASL = 1092
'Public Const RES_Unit_AGL = 1093
'Public Const RES_Unit_AbbrevASL = 1094
'Public Const RES_Unit_AbbrevAGL = 1095

' Header
'Public Const RES_Hdr_Frequency = 1109
Public Const RES_Hdr_Horizontal = 1110
Public Const RES_Hdr_Vertical = 1111
Public Const RES_Hdr_Altitude = 1112
Public Const RES_Hdr_MagVar = 1113
Public Const RES_Hdr_Rotation = 1116

Public Const RES_Hdr_None = 1141
Public Const RES_Hdr_MetersWide = 1142
Public Const RES_Hdr_Copyright = 1143
Public Const RES_Hdr_AuthorName = 1144

' Runway
Public Const RES_Rwy_Length = 1150
Public Const RES_Rwy_Width = 1151
Public Const RES_Rwy_ID = 1152
Public Const RES_Rwy_Threshold_Length = 1170
Public Const RES_Rwy_Overrun_Length = 1171
Public Const RES_Rwy_Strobes = 1176
Public Const RES_Rwy_HDistance = 1180
Public Const RES_Rwy_VDistance = 1181
Public Const RES_Rwy_RowSeparation = 1182
Public Const RES_Rwy_GlideSlope = 1183
Public Const RES_Rwy_SignsOffset = 1184
Public Const RES_Rwy_InnerMarker = 1191
Public Const RES_Rwy_MiddleMarker = 1192
Public Const RES_Rwy_OuterMarker = 1193

Public Const RES_Rwy_VASIlbl = 1185
Public Const RES_Rwy_PAPIlbl = 1186
Public Const RES_Rwy_Intensity = 1200
Public Const RES_Rwy_Position = 1205
Public Const RES_Rwy_Markers = 1210
Public Const RES_Rwy_OneSided = 1218
Public Const RES_Rwy_Closed = 1219
Public Const RES_Rwy_STOL = 1220
Public Const RES_Rwy_ApprLights = 1240
Public Const RES_Rwy_VASI = 1255
Public Const RES_Rwy_PAPI = 1268
Public Const RES_Rwy_Tab4 = 1283
Public Const RES_Rwy_Tab6 = 1284

' Shape
Public Const RES_Shp_Scale = 1303
Public Const RES_Shp_Spacing = 1305
Public Const RES_Shp_LineWidth = 1307
Public Const RES_Shp_Object = 1309
Public Const RES_Shp_Altitude = 1310
Public Const RES_Shp_Width = 1311
Public Const RES_Shp_LineObjWidth = 1312
Public Const RES_Shp_Type = 1313
Public Const RES_Shp_NightOnly = 1316
Public Const RES_Shp_FlatAltitude = 1318
Public Const RES_Shp_V1 = 1320
Public Const RES_Shp_Z = 1321
Public Const RES_Shp_TaxiType = 1322
Public Const RES_Shp_ArcRadius = 1323
Public Const RES_Shp_Layer1 = 1330
Public Const RES_Shp_Road1 = 1350
Public Const RES_Shp_CmbNone = 1355
Public Const RES_Shp_CmbFlat = 1360
Public Const RES_Shp_TaxiwayLine1 = 1370

Public Const RES_Bldg_Length = 1400
Public Const RES_Bldg_Width = 1401
Public Const RES_Bldg_Altitude = 1402
Public Const RES_Bldg_Height = 1406
Public Const RES_Bldg_Repeat = 1408
Public Const RES_Bldg_NotToScale = 1409
Public Const RES_Bldg_RLength = 1411
Public Const RES_Bldg_RWidth = 1412
Public Const RES_Bldg_RoofLight1 = 1415
Public Const RES_Bldg_Shape1 = 1420
Public Const RES_Bldg_ShapeP = 1427
Public Const RES_Bldg_Levels1 = 1430
Public Const RES_Bldg_Texture1 = 1440

' Macro
Public Const RES_Macro_Range = 1451
Public Const RES_Macro_Scale = 1452
Public Const RES_Macro_Altitude = 1453
Public Const RES_Macro_V1 = 1454
Public Const RES_Macro_V2 = 1455
Public Const RES_Macro_Params = 1457
Public Const RES_Macro_Parameters = 1462
Public Const RES_Macro_SelectCaption = 1480

' Radio
Public Const RES_Rdo_Name = 1600
Public Const RES_Rdo_ID = 1601
Public Const RES_Rdo_Text = 1602
Public Const RES_Rdo_FrequencyATIS = 1603
Public Const RES_Rdo_FrequencyVOR = 1604
Public Const RES_Rdo_FrequencyNDB = 1605
Public Const RES_Rdo_Range = 1606
Public Const RES_Rdo_Runway = 1607
Public Const RES_Rdo_BeamWidth = 1612
Public Const RES_Rdo_LocalizerPos = 1630

' Code
Public Const RES_Cde_Horz = 1703
Public Const RES_Cde_Vert = 1704

' Background
Public Const RES_Back_Image = 1750
Public Const RES_Back_ZoomX = 1754
Public Const RES_Back_ZoomY = 1755

' Exclusion, SurfaceArea
Public Const RES_Exc_Horz = 1800
Public Const RES_Exc_Vert = 1801
Public Const RES_Suf_Height = 1803
Public Const RES_Suf_Type = 1820

' Tower
Public Const RES_Tow_Height = 1850
Public Const RES_Tow_Frequency1 = 1851

' Point
Public Const RES_Pnt_Lighting = 1950
Public Const RES_Pnt_Style = 1951
Public Const RES_Pnt_NormalPoly = 1960
Public Const RES_Pnt_NormalLine = 1965

' Main
Public Const RES_Main_DoSave = 2000

Public Const RES_Main_Loading = 2010
Public Const RES_Main_Saving = 2011
Public Const RES_Main_Autosaving = 2012
Public Const RES_Main_Compiling = 2013
Public Const RES_Main_CopyingFiles = 2014

Public Const RES_Main_CurPos = 2020
Public Const RES_Main_CurPosXY = 2021
Public Const RES_Main_Distance = 2022
Public Const RES_Main_Object = 2023

Public Const RES_Main_SaveCaption = 2030
Public Const RES_Main_OpenCaption = 2031
Public Const RES_Main_CompileCaption = 2033
Public Const RES_Main_LinkCaption = 2034
Public Const RES_Main_LinkOutputCaption = 2034

Public Const RES_SaveFilter = 2050
Public Const RES_OpenFilter = 2051
Public Const RES_MacroFilter = 2052
Public Const RES_CompileFilter = 2053
Public Const RES_LinkFilter = 2054

'Public Const RES_Main_Object = 2013
'Public Const RES_Main_Airport2 = 2014

' About box
Public Const RES_About_Version = 2102
Public Const RES_About_Copyright = 2103

' Symbol
Public Const RES_Sym_ShortcutKey = 2142

' Color
Public Const RES_Col_Night = 2164

' Zoom
Public Const RES_Zoom_Value = 2180

' Transform
Public Const RES_Trans_Rotate = 2193

' Texture
Public Const RES_TEX_Texture = 2200
Public Const RES_TEX_Background = 2201

Public Const RES_TEX_Browse = 2220
Public Const RES_TEX_SaveAsBitmap = 2221
Public Const RES_TEX_PictureFilter = 2222
Public Const RES_TEX_TextureFilter = 2223
Public Const RES_TEX_BitmapFilter = 2224

Public Const RES_Dist_ChangeFolder = 2250
Public Const RES_Dist_AddFileCaption = 2251
Public Const RES_Dist_AddFileFilter = 2252
Public Const RES_CMB_Method = 2260
Public Const RES_CMB_Destination = 2265

' Preferences
Public Const RES_OPT_GeneralMain = 2310
Public Const RES_OPT_RememberWindowState = 2311
Public Const RES_OPT_NeatRecentFiles = 2312
Public Const RES_OPT_ShowHeaderProperties = 2313
Public Const RES_OPT_Remember = 2314
Public Const RES_OPT_ShowFractionalMinutes = 2315
Public Const RES_OPT_UnitOfMeasureMain = 2316
Public Const RES_OPT_UnitOfMeasure1 = 2317
Public Const RES_OPT_OrientationMain = 2319
Public Const RES_OPT_Orientation1 = 2320
Public Const RES_OPT_OldStyleMenus = 2322
Public Const RES_OPT_SaveCompressed = 2323
Public Const RES_OPT_ShowExportWizard = 2324
Public Const RES_OPT_UseMacroDefaults = 2325

Public Const RES_OPT_FSVersionMain = 2330
Public Const RES_OPT_FSVersion1 = 2331

Public Const RES_OPT_ExportMain = 2340
Public Const RES_OPT_EditConfig = 2341
Public Const RES_OPT_AutoCompress = 2342
Public Const RES_OPT_KeepSourceFile = 2343
Public Const RES_OPT_SaveBeforeCompile = 2344

Public Const RES_OPT_AppearanceMain = 2350
Public Const RES_OPT_CrosshairMain = 2351
Public Const RES_OPT_Crosshair1 = 2352
Public Const RES_OPT_FocusCircles = 2354
Public Const RES_OPT_PointCircles = 2355
Public Const RES_OPT_SnapPoints = 2356
Public Const RES_OPT_FillPolygons = 2357
Public Const RES_OPT_ThickLines = 2358
Public Const RES_OPT_FillObjects = 2359
Public Const RES_OPT_ShowCompass = 2360

Public Const RES_OPT_VisibleMain = 2370
  
Public Const RES_OPT_FSFolder = 2400
Public Const RES_OPT_TexFolder = 2401
Public Const RES_OPT_Compiler = 2402
Public Const RES_OPT_Compress = 2403
Public Const RES_OPT_TextEditor = 2404
Public Const RES_OPT_MacroFolder = 2405
Public Const RES_OPT_MacroPicFolder = 2406

Public Const RES_OPT_lblMacros = 2435
Public Const RES_OPT_lblFavorites = 2436

Public Const RES_OPT_lblAutoSave = 2440
Public Const RES_OPT_lblGrid = 2442

Public Const RES_OPT_Macro = 2450
Public Const RES_OPT_Tool = 2451

Public Const RES_OPT_Scheme1 = 2460

Public Const RES_OPT_ExecutableFilter = 2470
Public Const RES_OPT_Favorites = 2471

Public Const RES_TIP_FIRST = 2800

' Error message
Public Const RES_ERR_Bound = 3000
Public Const RES_ERR_Bound2 = 3001
Public Const RES_ERR_Units = 3002
Public Const RES_ERR_Numeric = 3003
Public Const RES_ERR_SpecifyText = 3004
Public Const RES_ERR_SpecifyName = 3005
Public Const RES_ERR_SpecifyID = 3006
Public Const RES_ERR_RunwayID = 3007
Public Const RES_ERR_SpecifyFile = 3008
Public Const RES_ERR_RotUnits = 3009

Public Const RES_ERR_InvalidFormat = 3010
Public Const RES_ERR_HigherFSSCVersion = 3011
Public Const RES_ERR_HigherFSVersion = 3012
Public Const RES_ERR_CityObjects = 3013
Public Const RES_ERR_BadBackgroundRef = 3014
Public Const RES_ERR_Link = 3015

Public Const RES_ERR_NoFile = 3020
Public Const RES_ERR_FileExists = 3021
Public Const RES_ERR_DirExists = 3022
Public Const RES_ERR_TextureError = 3023
Public Const RES_ERR_DriveExists = 3024
Public Const RES_ERR_NoBackground = 3025
Public Const RES_ERR_MacroPic1 = 3026
Public Const RES_ERR_MacroPic2 = 3027
Public Const RES_ERR_DirCreate = 3028
Public Const RES_ERR_MacroOpen = 3029

Public Const RES_ERR_Poles = 3030
Public Const RES_ERR_NoMagVar = 3031
Public Const RES_ERR_LatParse = 3032
Public Const RES_ERR_Tower = 3033
Public Const RES_ERR_FlatArea = 3034
Public Const RES_ERR_BitmapConvert = 3035

Public Const RES_ERR_Compile = 3040
Public Const RES_ERR_CompilerPath = 3041
Public Const RES_ERR_MacroCompileFail = 3042
Public Const RES_ERR_CopyFail = 3043
Public Const RES_ERR_MakeFolderFail = 3044
Public Const RES_ERR_AFDData = 3045

Public Const RES_ERR_LocateTexture = 3050
Public Const RES_ERR_DotSpacingReq = 3051

Public Const RES_ERR_MacroPathInvalid = 3060
Public Const RES_ERR_NoAirport = 3061

' Object Names
Public Const RES_Obj_Header = 3200
Public Const RES_Obj_Runway = 3201
Public Const RES_Obj_Polygon = 3202
Public Const RES_Obj_Taxiway = 3203
Public Const RES_Obj_Road = 3204
Public Const RES_Obj_River = 3205
Public Const RES_Obj_Line = 3206
Public Const RES_Obj_TaxiwayLine = 3207
Public Const RES_Obj_Building = 3208
Public Const RES_Obj_Macro = 3209
Public Const RES_Obj_ATIS = 3210
Public Const RES_Obj_VOR = 3211
Public Const RES_Obj_NDB = 3212
Public Const RES_Obj_Tower = 3213
Public Const RES_Obj_MenuEntry = 3214
Public Const RES_Obj_Background = 3215
Public Const RES_Obj_Flatten = 3216
Public Const RES_Obj_SurfaceArea = 3217
Public Const RES_Obj_Exclusion = 3218
Public Const RES_Obj_Code = 3219
Public Const RES_Obj_Point = 3220
Public Const Res_Obj_Point2 = 3250
Public Const Res_Obj_Point3 = 3251

' Runway
Public Const RES_Rwy_Surface = 3300

' Synthetic Scenery
Public Const RES_Syn_Transparent = 3400

' Menu
Public Const RES_Menu_ShortcutKey = 3606
Public Const RES_Menu_Properties = 3608
Public Const RES_Menu_Properties2 = 3609

' Toolbar tooltips
Public Const RES_Toolbar1 = 3700
Public Const RES_Toolbar2 = 3750

' No localization
Public Const RES_License = 300
Public Const RES_ERR_Language = 310
Public Const RES_Lang_DialogBox1 = 320
Public Const RES_Lang_DialogBox2 = 321
Public Const RES_Lang_Browser = 322

Public Const RES_UnlocalizedObjectNames = 500
Public Const RES_SynData = 1000
Public Const RES_Bldg1Data = 1100
Public Const RES_Bldg2Data = 1200
Public Const RES_BldgRData = 1300
Public Const RES_Regions = 2000
Public Const RES_LangCode = 2100
