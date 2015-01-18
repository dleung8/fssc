VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compile Language"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Txts 
      Height          =   285
      Index           =   3
      Left            =   1560
      TabIndex        =   9
      Top             =   1680
      Width           =   3255
   End
   Begin VB.TextBox Txts 
      Height          =   285
      Index           =   2
      Left            =   1560
      TabIndex        =   7
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox Txts 
      Height          =   285
      Index           =   1
      Left            =   1560
      TabIndex        =   4
      Top             =   720
      Width           =   3255
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   375
      Index           =   1
      Left            =   5040
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtLog 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   2760
      Width           =   6015
   End
   Begin VB.CommandButton cmdCompile 
      Caption         =   "&Compile"
      Default         =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   375
      Index           =   0
      Left            =   5040
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Txts 
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label lbls 
      Caption         =   "Help output dir:"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   8
      Top             =   1710
      Width           =   1215
   End
   Begin VB.Label lbls 
      Caption         =   "Text output dir:"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   6
      Top             =   1230
      Width           =   1215
   End
   Begin VB.Label lbls 
      Caption         =   "Help source file:"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   750
      Width           =   1215
   End
   Begin VB.Label lblLog 
      Caption         =   "Output log:"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label lbls 
      Caption         =   "Text source file:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   270
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type IType
  Keyword As String
  File As String
  Title As String
End Type

Private Sub AddText(S As String)
  txtLog.Text = txtLog.Text & S & vbCrLf & vbCrLf
  DoEvents
End Sub

Private Sub CompileHelpFile()
  Dim File As String, Title As String
  Dim DestPath As String, FileTitle As String
  Dim InputStr As String
  Dim TempStr As String, TempVal As Integer
  Dim NextAsDef As Boolean
  Dim CurIndex As Integer, I As Integer
  Dim NoHeader As Boolean
  Dim NoFooter As Boolean
  
  Dim Y As Integer
  
  Dim Keywords() As IType, Res() As String
  Dim KeywordIndex As Integer
  
  Dim CloseWindowFlag As Boolean
  
  ReDim Keywords(100)
  
  AddText "Compiling Help File..."
  
  DestPath = Txts(3).Text
  FileTitle = LCase$(MakeFileNameNeat(Txts(1).Text))
  If Right$(DestPath, 1) <> "\" Then
    DestPath = DestPath & "\"
  End If
  
  MakePath DestPath
  Open Txts(1).Text For Input As #1
  Open DestPath & FileTitle & ".hhp" For Output As #3
  Print #3, "[Options]"
  Print #3, "Compatibility = 1.1 Or later"
  Print #3, "Compiled File = " & FileTitle & ".chm"
  Print #3, "Index file = " & FileTitle & ".hhk"
  Print #3, "Contents File = " & FileTitle & ".hhc"
  Print #3, "Default Window = Main"
  Print #3, "Default topic = Index.html"
  Print #3, "Language=0x409 English (United States)"
  Print #3, "Title=FS Scenery Creator Help"
  Print #3,
  Print #3, "[WINDOWS]"
  Print #3, "Main=""FS Scenery Creator Help"",""" & FileTitle & ".hhc"",""" & FileTitle & ".hhk"",""index.html"",""index.html"",,,,,0x2120,,0x304e,,,,,,0,,0"
  Print #3, "Tutorial=""Tutorial"",,,,,,,,,0x0,,0x0,,,,,,,,0"
  Print #3,
  Print #3, "[Files]"
  
  Open DestPath & FileTitle & ".hhc" For Output As #4
  
  Print #4, "<!DOCTYPE HTML PUBLIC ""-//IETF//DTD HTML//EN"">"
  Print #4, "<HTML>"
  Print #4, "<HEAD>"
  Print #4, "<!-- Sitemap 1.0 -->"
  Print #4, "</HEAD><BODY>"

  Open DestPath & FileTitle & ".hhk" For Output As #5
  
  Print #5, "<!DOCTYPE HTML PUBLIC ""-//IETF//DTD HTML//EN"">"
  Print #5, "<HTML>"
  Print #5, "<HEAD>"
  Print #5, "<!-- Sitemap 1.0 -->"
  Print #5, "</HEAD><BODY>"
  Print #5, "<UL>"

  Do Until EOF(1)
    Line Input #1, InputStr
    If Left$(InputStr, 5) = "File:" Then
      On Error Resume Next
      
      If CloseWindowFlag Then
        Print #2, "<OBJECT ID=CloseMethod TYPE=""application/x-oleobject"" CLASSID=""clsid:adb880a6-d8ff-11cf-9377-00aa003b7a11""><PARAM NAME=""Command"" VALUE=""Close""></OBJECT>"
        CloseWindowFlag = False
      End If
      Print #2, "</FONT></BODY></HTML>"
      On Error GoTo 0
      Close #2
      TempStr = Trim$(Mid$(InputStr, 6))
      Print #3, TempStr
      MakeMyDir DestPath, TempStr
      Open DestPath & TempStr For Output As #2
      Print #2, "<HTML>"
      File = TempStr
    ElseIf Left$(InputStr, 6) = "Title:" Then
      TempStr = Trim$(Mid$(InputStr, 7))
      Print #2, "<HEAD><TITLE>" & TempStr & "</TITLE></HEAD>"
      Print #2, "<BODY><FONT FACE=""Arial"" SIZE=2>"
      Print #2, "<H3>" & TempStr & "</H3>" & vbCrLf
      Title = TempStr
    ElseIf Left$(InputStr, 6) = "Level:" Then
      TempVal = (Trim$(Mid$(InputStr, 7)))
      If TempVal - CurIndex > 1 Then
        MsgBox "Error: Level Error at " & File & "."
      End If
      If TempVal - CurIndex = 1 Then
        Print #4, String$(CurIndex, vbTab) & "<UL>"
      Else
        For I = CurIndex - 1 To TempVal Step -1
          Print #4, String$(I, vbTab) & "</UL>"
        Next I
      End If
      CurIndex = TempVal
      Print #4, String$(CurIndex, vbTab) & "<LI> <OBJECT type=""text/sitemap"">"
      Print #4, String$(CurIndex + 1, vbTab) & "<param name=""Name"" value=""" & Title & """>"
      Print #4, String$(CurIndex + 1, vbTab) & "<param name=""Local"" value=""" & File & """>"
      Print #4, String$(CurIndex + 1, vbTab) & "</OBJECT>"
    ElseIf Left$(InputStr, 1) = ";" Then
      ' Comment
    ElseIf Left$(InputStr, 9) = "Keywords:" Then
      Res = Split(Mid$(InputStr, 10), ",")
      If Res(0) <> "" Then
        For I = 0 To UBound(Res)
          If KeywordIndex > UBound(Keywords) Then ReDim Preserve Keywords(KeywordIndex * 2)
          With Keywords(KeywordIndex)
            .Keyword = Trim$(Res(I))
            .File = File
            .Title = Title
          End With
          KeywordIndex = KeywordIndex + 1
        Next I
      End If
    Else
      If Len(Trim$(InputStr)) > 0 Then
        If InStr(InputStr, "<A PARAM=""Close"">") > 0 Then
          InputStr = Replace(InputStr, "<A PARAM=""Close"">", "<A HREF=""JavaScript:CloseMethod.Click()"">")
          CloseWindowFlag = True
        End If
        InputStr = Replace(Replace(InputStr, "<WARN>", "<B><FONT COLOR=""RED"">"), "</WARN>", "</FONT></B>")
        ProcessString InputStr
        If NextAsDef Then
          Print #2, "<DD>" & InputStr & "</DD>"
          Print #2, "</DL>" & vbCrLf
          NextAsDef = False
        ElseIf Right$(InputStr, 1) = ":" Then
          Print #2, "<DL>"
          Print #2, "<DT><B>" & InputStr & "</B></DT>"
          NextAsDef = True
        ElseIf Left$(InputStr, 4) = "<UL>" Or Left$(InputStr, 4) = "<LI>" Then
          Print #2, InputStr
        ElseIf Left$(InputStr, 5) = "</UL>" Then
          Print #2, InputStr & vbCrLf
        ElseIf Right$(InputStr, 2) = ":." Then
          Print #2, "<P>" & Left$(InputStr, Len(InputStr) - 1) & "</P>" & vbCrLf
        Else
          Print #2, "<P>" & InputStr & "</P>" & vbCrLf
        End If
      End If
    End If
  Loop
  Print #2, "</FONT></BODY></HTML>"
  For I = CurIndex - 1 To 0 Step -1
    Print #4, String$(I, vbTab) & "</UL>"
  Next I
  Print #4, "</BODY></HTML>"
  
  ReDim Preserve Keywords(KeywordIndex - 1)
  Sort Keywords
  
  CurIndex = 1
  For I = 0 To UBound(Keywords)
    With Keywords(I)
      Y = InStr(Keywords(I).Keyword, "\")
      If Y > 0 Then
        If CurIndex = 1 Then
          Print #5, vbTab & "<UL>"
          CurIndex = 2
        End If
      Else
        If CurIndex = 2 Then
          Print #5, vbTab & "</UL>"
          CurIndex = 1
        End If
      End If
      If I > 0 Then
        If Keywords(I - 1).Keyword = .Keyword Then
          NoHeader = True
        End If
      End If
      If I < UBound(Keywords) Then
        If Keywords(I + 1).Keyword = .Keyword Then
          NoFooter = True
        End If
      End If
      
      If Not NoHeader Then
        Print #5, String$(CurIndex, vbTab) & "<LI> <OBJECT type=""text/sitemap"">"
        If Y > 0 Then
          Print #5, String$(CurIndex + 1, vbTab) & "<param name=""Name"" value=""" & Mid$(.Keyword, Y + 1) & """>"
        Else
          Print #5, String$(CurIndex + 1, vbTab) & "<param name=""Name"" value=""" & .Keyword & """>"
        End If
      End If
      
      Print #5, String$(CurIndex + 1, vbTab) & "<param name=""Name"" value=""" & .Title & """>"
      Print #5, String$(CurIndex + 1, vbTab) & "<param name=""Local"" value=""" & .File & """>"
      
      If Not NoFooter Then
        Print #5, String$(CurIndex + 1, vbTab) & "</OBJECT>"
      End If
      
      NoHeader = False
      NoFooter = False
    End With
  Next I
  
  Print #5, "</UL>"
  Print #5, "</BODY></HTML>"
  
  Print #3,
  Print #3, "[TEXT POPUPS]"
  Print #3, "what_" & Right$(FileTitle, 3) & ".txt"
  
  Close

  AddText "Done Help File"

End Sub

Private Sub CompileTextFile()
  Dim File As String, _
    FileNum1 As Integer, FileNum2 As Integer, _
    DestFile As String, _
    Temp As String, Temp2 As String, Temp3 As String, _
    Num As String, Y As Integer

  Dim LastNum As Integer

  Dim Ver1 As Integer, Ver2 As Integer, Ver3 As Integer, _
    FileDescrip As String
    
  Dim StrCount As Integer

  FileNum1 = FreeFile

  txtLog.Text = ""
  File = Txts(0).Text
  AddText "Opening " & File
  Open File For Input As #FileNum1

  Do Until EOF(FileNum1)
    Line Input #FileNum1, Temp
    If InStr(Temp, "http://") = 0 Then
      Y = InStr(Temp, "//")
      If Y > 0 Then Temp = Left$(Temp, Y - 1)
    End If

    Do
      Y = InStr(Temp, Chr$(9))
      If Y > 0 Then Mid$(Temp, Y, 1) = " "
    Loop Until Y = 0

    Do
      Y = InStr(Temp, "\n")
      If Y > 0 Then Mid$(Temp, Y, 2) = Chr$(13) & Chr$(10)
    Loop Until Y = 0

    Do
      Y = InStr(Temp, "\t")
      If Y > 0 Then Temp = Left$(Temp, Y - 1) & Chr$(9) & Mid$(Temp, Y + 2)
    Loop Until Y = 0
    
    Do
      Y = InStr(Temp, "\" & Chr$(34))
      If Y > 0 Then Temp = Left$(Temp, Y - 1) & Mid$(Temp, Y + 1)
    Loop Until Y = 0

    Temp = Trim$(Temp)
    Temp2 = ReadNext(Temp, " ", True)
    If Temp2 = "FILEVERSION" Then
      Ver1 = Val(ReadNext(Temp, ",", True))
      Ver2 = Val(ReadNext(Temp, ",", True))
      Ver3 = Val(ReadNext(Temp, ",", True))
      AddText "File version: " & Ver1 & "." & Ver2 & "." & Ver3
    ElseIf Temp2 = "VALUE" Then
      Temp2 = ReadNext(Temp, ",", True)
      Temp3 = Mid$(Temp, 2, Len(Temp) - 4)
      If InStr(Temp2, "InternalName") > 0 Then
        DestFile = AddDir(Txts(2).Text, Temp3 & ".dat")
        FileNum2 = FreeFile
        AddText "Output file: " & DestFile
        If FileExists(DestFile) Then Kill DestFile
        Open DestFile For Binary As #FileNum2
        Put #FileNum2, , "FS Scenery Creator Language File" & Chr$(26)
        Put #FileNum2, , CLng(Ver1 * 100000 + Ver2 * 1000 + Ver3 * 10)
      ElseIf InStr(Temp2, "LegalCopyright") > 0 Then
        Put #FileNum2, , CInt(Len(Temp3))
        Put #FileNum2, , Temp3
        AddText "Copyright: " & Temp3
      End If
    ElseIf Val(Temp2) > 0 Then
      If Val(Temp2) <= LastNum Then
        MsgBox "Error: String ID's not sorted. Found: " & Temp2 & " after " & LastNum
        Me.MousePointer = vbDefault
        Close
        Exit Sub
      End If
      Put #FileNum2, , CInt(Temp2)
      LastNum = Val(Temp2)
      Temp3 = Temp
      If UnquoteString(Temp3) Then
        ProcessString Temp3
        Put #FileNum2, , CInt(Len(Temp3))
        Put #FileNum2, , Temp3
        StrCount = StrCount + 1
      Else
        MsgBox "Error: Missing quotes at string " & LastNum
        Me.MousePointer = vbDefault
        Close
        Exit Sub
      End If
    End If
  Loop
  AddText "Strings compiled: " & StrCount
  Close
End Sub

Private Sub MakeMyDir(ByRef Base As String, ByVal Add As String)
  Dim Y As Long, TempStr As String
  Y = InStrRev(Add, "\")
  If Right$(Base, 1) <> "\" Then Base = Base & "\"
  TempStr = Base & Left$(Add, Y)
  MakePath TempStr
End Sub

Private Sub MakePath(ByVal strDirName As String)
  Dim strPath As String
  Dim intOffset As Integer
  Dim intAnchor As Integer
  Dim strOldPath As String

  On Error Resume Next

  '
  'Add trailing backslash
  '
  If Right$(strDirName, 1) <> "\" Then
    strDirName = strDirName & "\"
  End If

  strOldPath = CurDir$
  intAnchor = 0

  '
  'Loop and make each subdir of the path separately.
  '
  intOffset = InStr(intAnchor + 1, strDirName, "\")
  intAnchor = intOffset 'Start with at least one backslash, i.e. "C:\FirstDir"
  Do
    intOffset = InStr(intAnchor + 1, strDirName, "\")
    intAnchor = intOffset

    If intAnchor > 0 Then
      strPath = Left$(strDirName, intOffset - 1)
      ' Determine if this directory already exists
      Err = 0
      ChDir strPath
      If Err Then
        ' We must create this directory
        Err = 0
        MkDir strPath
      End If
    End If
  Loop Until intAnchor = 0
  ChDir strOldPath
  Err = 0
End Sub

Private Sub ProcessString(ByRef myStr As String)
  myStr = Replace(Replace(myStr, " :", Chr$(160) + ":"), " !", Chr$(160) + "!")
End Sub

Private Sub Sort(Arr() As IType)
  Dim I As Integer, J As Integer, Min As Integer
  Dim InputStr As IType
  For I = 0 To UBound(Arr) - 1
    Min = I
    For J = I + 1 To UBound(Arr)
      If Arr(J).Keyword < Arr(Min).Keyword Then Min = J
    Next J
    If Min <> I Then
      InputStr = Arr(I)
      Arr(I) = Arr(Min)
      Arr(Min) = InputStr
    End If
  Next I
End Sub

Private Function UnquoteString(S As String) As Boolean
  If Left$(S, 1) = Chr$(34) And Right$(S, 1) = Chr$(34) Then
    S = Mid$(S, 2, Len(S) - 2)
    UnquoteString = True
  End If
End Function

Private Sub cmdBrowse_Click(Index As Integer)
  Dim X As String
  Filter = "Text files(*.txt)|*.txt"
  DefExt = "txt"
  FilterIndex = 1
  X = OpenDialog(, Txts(Index).Text)
  If X <> "" Then Txts(Index).Text = X
End Sub

Private Sub cmdCompile_Click()
  Me.MousePointer = vbHourglass
  
  If Txts(0).Text <> "" Then CompileTextFile
  If Txts(1).Text <> "" Then CompileHelpFile
  
  AddText "Compilation successful."
  Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
  AppKey = "Software\Leung\FS Scenery Creator\LangConv"
  Txts(0).Text = RegGetKey("TextSource", "")
  Txts(1).Text = RegGetKey("HelpSource", "")
  Txts(2).Text = RegGetKey("TextDir", "")
  Txts(3).Text = RegGetKey("HelpDir", "")
End Sub

Private Sub Form_Unload(Cancel As Integer)
  RegSetKey "TextSource", Txts(0).Text
  RegSetKey "HelpSource", Txts(1).Text
  RegSetKey "TextDir", Txts(2).Text
  RegSetKey "HelpDir", Txts(3).Text
End Sub

Private Sub Txts_GotFocus(Index As Integer)
  SelectText Txts(Index)
End Sub
