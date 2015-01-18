VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCompile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mLines() As String
Private Cnt As Integer

Public Sub Add(ByVal LineStr As String)
  If Cnt > UBound(mLines) Then
    ReDim Preserve mLines(Cnt * 2)
  End If
  mLines(Cnt) = LineStr
  Cnt = Cnt + 1
End Sub

Public Function GetStr() As String
  ReDim Preserve mLines(Cnt - 1)
  GetStr = Join(mLines, vbNewLine)
End Function

Private Sub Class_Initialize()
  ReDim mLines(100)
End Sub

Private Sub Class_Terminate()
  Erase mLines
End Sub
