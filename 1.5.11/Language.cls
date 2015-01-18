VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Language Class
' Handles localization of strings and text

Private Type LangType
  Name As String
  File As String
  RealFile As String
  Exists As Boolean
End Type

' Language list
Private Languages() As LangType
Private LangInd As Integer

' String list
Private ResStrings(1000 To 3799) As String

Public Event Changed()

Public HelpFile As String

' Given an index, add the corresponding text strings
' to a control
Public Sub AddItems(Ctrl As Control, ByVal Low As Integer, ByVal Count As Integer)
  Dim I As Integer
  For I = Low To Low + Count - 1
    Ctrl.AddItem GetString(I)
  Next I
End Sub

' Fill the specified listbox with the names of
' available languages
Public Sub FillList(X As ListBox)
  Dim I As Integer
  With X
    For I = 0 To UBound(Languages)
      .AddItem Languages(I).Name
      .ItemData(I) = -Languages(I).Exists
    Next I
    '.AddItem "Search for More Languages"
    .ListIndex = LangInd
  End With
End Sub

' Format an error string using the current language
' specifying replacement strings
Public Function FormatErrorMessage(ByVal IDLabel As String, ByVal Low As Single, ByVal High As Single, Optional ByVal Units As Integer)
  Dim LowStr As String
  If Abs(Low) < 1 And Low <> 0 Then
    LowStr = Format$(Low, "0.0#########")
  Else
    LowStr = CStr(Low)
  End If
  If Units <= 0 Then
    FormatErrorMessage = ResolveString(RES_ERR_Bound2, IDLabel, LowStr, CStr(High))
  Else
    FormatErrorMessage = ResolveString(RES_ERR_Bound, IDLabel, LowStr, CStr(High), GetString(Units))
  End If
End Function

' Get the closest available language to the current system
Public Function GetClosestLanguage() As String
  Dim UserLanguage As String, I As Integer, _
    MainLang As String, OtherLang As String
    
  UserLanguage = LocaleInfo(LOCALE_SABBREVLANGNAME)
  For I = 0 To UBound(Languages)
    With Languages(I)
      If Mid$(.RealFile, 6, 3) = UserLanguage Then
        ' Found
        GetClosestLanguage = .Name
        Exit Function
      ElseIf Mid$(.RealFile, 6, 2) = Left$(UserLanguage, 2) Then
        If Mid$(.RealFile, 8, 1) = "." Then
          ' Found main language
          MainLang = .Name
        Else
          ' Other related dialect
          OtherLang = .Name
        End If
      End If
    End With
  Next I
  If MainLang = "" And OtherLang = "" Then
    MainLang = "English (United States)"
  ElseIf MainLang = "" Then
    MainLang = OtherLang
  End If
  GetClosestLanguage = MainLang
End Function

' Get the list of available languages
Public Sub GetLanguageList()
  Dim File As String, FileNum As Integer, Cnt As Integer, _
    Temp As String, Ver As Long, Num As Integer
  
  ReDim Languages(0)
  FileNum = FreeFile
  
  File = MultiDir("lang_*.dat", App.Path)

  Do Until File = ""
    Open File For Binary As #FileNum
    
    Temp = String$(33, 0)
    Get #FileNum, , Temp
    If Temp = "FS Scenery Creator Language File" & Chr$(26) Then
      Get #FileNum, , Ver
     If Ver >= 104000 Then
        ' Copyright
        GetText FileNum
        Get #FileNum, , Num
        If Num = RES_LangName Then
          Temp = GetText(FileNum)
          ReDim Preserve Languages(Cnt)
          With Languages(Cnt)
            .Name = Temp
            .File = File
            .RealFile = File
            .Exists = True
            Get #FileNum, , Num
            Temp = GetText(FileNum)
            If Num = 9999 Then
              If Temp = "None" Then
                .Exists = False
              Else
                .File = "lang_" & Temp & ".dat"
              End If
            End If
            Cnt = Cnt + 1
          End With
        End If
     End If
    End If
    Close #FileNum
    File = MultiDir()
  Loop
End Sub

' Get a string based on an ID
Public Function GetString(ByVal ID As Integer) As String
  If ID < LBound(ResStrings) Or ID > UBound(ResStrings) Then
    GetString = "Language File Error"
  Else
    GetString = ResStrings(ID)
  End If
End Function

' Get a piece of text from a binary file
Public Function GetText(ByVal FileNum As Integer) As String
  GetBinaryData FileNum, GetText
End Function

' Return the current name of the language
Public Property Get Name() As String
  If LangInd <> -999 Then Name = Languages(LangInd).Name
End Property

' Set a language
Public Property Let Name(ByVal Lang As String)
  Dim I As Integer, File As String, X As Integer, _
    Temp As String, Num As Integer, Ver As Long
  
  LangInd = 0
  For I = 0 To UBound(Languages)
    If Languages(I).Name = Lang Then LangInd = I: Exit For
  Next I
  File = Languages(LangInd).File
  
  If File = "" Then GoTo NoLibrary:
  ' FSSC Native Format
  X = FreeFile
  Open File For Binary As #X
  Temp = String$(33, 0)
  Get #X, , Temp
  If Temp = "FS Scenery Creator Language File" & Chr$(26) Then
    Get #X, , Ver
    If Ver >= 104000 Then
      Temp = GetText(X)
      ' Temp = Copyright message. Not needed right now

      Do
        Get #X, , Num
        If Num = 9999 Then Exit Do
        ResStrings(Num) = GetText(X)
      Loop
   End If
  End If
  Close #X
  
  App.Title = GetString(RES_Title)
  HelpFile = AddDir(App.Path, GetString(RES_LangHelpName))
  App.HelpFile = HelpFile & "::/what_" & Left$(Right$(File, 7), 3) & ".txt"
  RaiseEvent Changed
  Exit Property
NoLibrary:
  If Lang = "English (United States)" Then
    MsgBox LoadResString(RES_ERR_Language), vbCritical
    LangInd = -999
  Else
    MsgBox Lang & Mid$(LoadResString(311), 3), vbCritical
    Name = "English (United States)"
  End If
End Property

' Processes a form to localize controls
Public Sub PrepareForm(X As Form)
  Dim Ctrl As Control, I As Integer
  On Error Resume Next

  If Val(X.Tag) > 0 Then X.Caption = GetString(Val(X.Tag))
  For Each Ctrl In X
    If TypeOf Ctrl Is TabStrip Then
      For I = 1 To Ctrl.Tabs.Count
        If Val(Ctrl.Tabs(I).Tag) > 0 Then
          Ctrl.Tabs(I).Caption = "  " & GetString(Val(Ctrl.Tabs(I).Tag)) & "  "
        End If
      Next I
    ElseIf TypeOf Ctrl Is ListView Then
      For I = 1 To Ctrl.ColumnHeaders.Count
        If Val(Ctrl.ColumnHeaders(I).Tag) > 0 Then
          Ctrl.ColumnHeaders(I).Text = GetString(Val(Ctrl.ColumnHeaders(I).Tag))
        End If
      Next I
    ElseIf Not TypeOf Ctrl Is TextBox Then
      If Val(Ctrl.Tag) > 0 Then
        Ctrl.Caption = GetString(Val(Ctrl.Tag))
      End If
    End If
  Next
End Sub

' Get a string based on an ID and edit the string
' with the given parameters
Public Function ResolveString(ByVal ID As Integer, ParamArray Args() As Variant)
  Dim myStr As String, Var As Integer, Pos As Long
  myStr = GetString(ID)
  Pos = InStr(myStr, "%")
  On Error GoTo RSErr:
  Do Until Pos = 0
    Var = Val(Mid$(myStr, Pos + 1, 1))
    myStr = Left$(myStr, Pos - 1) & Args(Var - 1) & Mid$(myStr, Pos + 2)
    Pos = InStr(Pos + Len(Args(Var - 1)), myStr, "%")
  Loop
  ResolveString = myStr
  Exit Function
RSErr:
  MsgBox "Error with language.cls with ID " & ID & vbCrLf & "myStr: " & myStr & vbCrLf & "# of arguments:" & UBound(Args), vbInformation
  Exit Function
End Function

Private Sub Class_Initialize()
  GetLanguageList
End Sub
