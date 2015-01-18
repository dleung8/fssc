VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAPTData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Airport Class
' Handles the low level aspect of loading Airport files

Option Explicit

Private Type AirportType
  Key As String
  Data As String
End Type

Public Section As String
Public FileNum As Integer

Private Data() As AirportType
Private Cnt As Integer

' Convert a string detail type into the complexity enum
Public Function GetComplexityData() As Byte
  GetComplexityData = MatchText(GetData("Detail"), "Very Sparse", "Sparse", "Normal", "Dense", "Very Dense")
End Function

' Get a data element "Key" from the file
' Accelerates the loading process by keeping track of the
' pointer of the last load (efficient when elements are
' in the order expected by FSSC)
Public Function GetData(ByVal Key As String) As String
  Dim I As Integer
  If StrComp(Data(Cnt).Key, Key, vbTextCompare) = 0 Then
    GetData = Data(Cnt).Data
  Else
    For I = 0 To UBound(Data)
      If StrComp(Data(I).Key, Key, vbTextCompare) = 0 Then
        GetData = Data(I).Data
        Cnt = I
        Exit For
      End If
    Next I
  End If
  If Cnt < UBound(Data) And I <= UBound(Data) Then
    Cnt = Cnt + 1
  Else
    Cnt = 0
  End If
End Function

' Converts all numerical data values to the
' current locale
Private Sub ProcessData()
  Dim Sep As String, SepOther As String, _
    I As Integer, Y As Long
  Sep = LocaleInfo(LOCALE_SDECIMAL)
  SepOther = IIf(Sep = ".", ",", ".")

  For I = 0 To UBound(Data)
    Y = 1
    With Data(I)
      Do
        Y = InStr(Y + 1, .Data, SepOther)
        If Y > 1 Then
          If InStr(DigitsSigns, Mid$(.Data, Y - 1, 1)) > 0 And InStr(Digits, Mid$(.Data, Y + 1, 1)) > 0 Then
            Mid$(.Data, Y, 1) = Sep
          End If
        End If
      Loop While Y > 1
    End With
  Next I
End Sub

' Loads the next section from the file
Public Sub ReadSection()
  Dim I As Integer, Temp As String
  ReDim Data(20)
  Section = ""
  Cnt = 0
  Do Until EOF(FileNum)
    LineInputEx FileNum, Temp
    If Trim$(Temp) <> "" And Left$(Trim$(Temp), 1) <> ";" Then Exit Do
  Loop
  If EOF(FileNum) Then Exit Sub
  Section = ReadNext(Temp, " ")
  I = 0
  Do Until EOF(FileNum)
    LineInputEx FileNum, Temp
    If Trim$(Temp) = "" Then Exit Do
    If Left$(Temp, 1) = "-" Then
      I = I + 1
      If I > UBound(Data) Then ReDim Preserve Data(I * 2)
      With Data(I)
        Temp = Mid$(Temp, 2)
        .Key = Trim$(ReadNext(Temp, "="))
        .Data = Temp
      End With
    End If
  Loop
  ReDim Preserve Data(I)
  ProcessData
End Sub
'
'' Moves the data pointer forward one element (for
'' acceleration of key access
'Public Sub SkipNext()
'  If Cnt < UBound(Data) Then Cnt = Cnt + 1
'End Sub
