VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' FileData Class
' Handles the reading and writing of FSSC files

Private Type FileDataType
  Key As String
  Data As String
End Type

Public Section As String
Public Comment As String
Public FileNum As Integer

Private Data() As FileDataType
Private Cnt As Integer

' Get a data element "Key" from the file
' Accelerates the loading process by keeping track of the
' pointer of the last load (efficient when elements are
' in the order expected by FSSC)
Public Sub GetData(ByVal Key As String, ByRef myData As Variant, Default As Variant)
  Dim I As Integer, Index As Long
  If StrComp(Data(Cnt).Key, Key, vbTextCompare) = 0 Then
    Index = Cnt
  Else
    Index = -1
    For I = 0 To UBound(Data)
      If StrComp(Data(I).Key, Key, vbTextCompare) = 0 Then
        Index = I
        Cnt = I
        Exit For
      End If
    Next I
  End If
  If Index > -1 Then
    Select Case VarType(myData)
      Case vbBoolean
        myData = CBool(Data(Index).Data)
      Case vbString
        myData = Data(Index).Data
      Case Else
        myData = Val(Data(Index).Data)
    End Select
    
    If Cnt < UBound(Data) Then
      Cnt = Cnt + 1
    Else
      Cnt = 0
    End If
  Else
    myData = Default
  End If
End Sub

' Clear the buffer
Public Sub InitializeForWrite(ByVal SectionName As String, ByVal SectionComment As String, Optional ByVal Count As Long = 25)
  Section = SectionName
  Comment = SectionComment
  ReDim Data(Count)
  Cnt = 0
End Sub

' Loads the next section from the file
Public Sub ReadSection()
  Dim I As Integer, Temp As String
  ReDim Data(20)
  Section = ""
  Cnt = 0
  Do
    LineInputEx FileNum, Temp
    If Trim$(Temp) <> "" And Left$(Trim$(Temp), 1) <> ";" Then Exit Do
  Loop Until EOF(FileNum)
  If EOF(FileNum) Then Exit Sub
  Section = Trim$(ReadNext(Temp, ";"))
  Section = Mid$(Section, 2, Len(Section) - 2)
  I = 0
  Do
    LineInputEx FileNum, Temp
    If Trim$(Temp) = "" Then Exit Do
    If I > UBound(Data) Then ReDim Preserve Data(I * 2)
    With Data(I)
      .Key = Trim$(ReadNext(Temp, "="))
      .Data = Trim$(Temp)
    End With
    I = I + 1
  Loop Until EOF(FileNum)
  If I > 0 Then
    ReDim Preserve Data(I - 1)
  Else
    ReDim Data(0)
  End If
End Sub

' Save a key and its data into the buffer
Public Sub SaveData(ByVal myKey As String, myData As Variant, Default As Variant)
  If Options.SaveCompressed Then
    If myData = Default Then Exit Sub
  End If
  With Data(Cnt)
    .Key = myKey
    Select Case VarType(myData)
      Case vbSingle, vbDouble
        .Data = Trim$(Str$(myData))
      Case vbBoolean
        .Data = CStr(-myData)
      Case Else ' vbInteger, vbLong, vbString
        .Data = CStr(myData)
    End Select
  End With
  Cnt = Cnt + 1
  If UBound(Data) < Cnt Then
    ReDim Preserve Data(Cnt * 2)
  End If
End Sub

' Write the current data buffer into the file
Public Sub WriteSection()
  Dim I As Integer
  If Not Options.SaveCompressed Then Print #FileNum, "; " & Comment
  Print #FileNum, "["; Section; "]"
  For I = 0 To Cnt
    With Data(I)
      If .Key = "" Then Exit For
      Print #FileNum, .Key & "=" & .Data
    End With
  Next I
  Print #FileNum,
End Sub
