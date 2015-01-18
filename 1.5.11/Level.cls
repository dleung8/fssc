VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Level Class
' Provides functions for maintaining building levels
'   for the Building object

Public TexID As Byte, _
       Height As Integer, _
       X As Single, _
       Y As Single, _
       Z As Single

Public Property Get DataStr() As String
  DataStr = TexID & " " & Height & " " & StrEx(X) & " " & StrEx(Y) & " " & StrEx(Z)
End Property

Public Property Let DataStr(ByVal Data As String)
  TexID = Val(ReadNext(Data, " "))
  Height = Val(ReadNext(Data, " "))
  X = Val(ReadNext(Data, " "))
  Y = Val(ReadNext(Data, " "))
  Z = Val(ReadNext(Data, " "))
End Property

' Reverse X and Y
Public Property Let DataStrOld(ByVal Data As String)
  TexID = Val(ReadNext(Data, " "))
  Height = Val(ReadNext(Data, " "))
  Y = Val(ReadNext(Data, " "))
  X = Val(ReadNext(Data, " "))
  Z = Val(ReadNext(Data, " "))
End Property

Public Sub LoadAPTData(ByVal Data As String)
  Height = ValEx(ReadNext(Data, " "))
  TexID = ValEx(ReadNext(Data, " "))
  X = ValEx(ReadNext(Data, " "))
  Y = ValEx(ReadNext(Data, " "))
  Z = ValEx(ReadNext(Data, " "))
End Sub

' Standard LoadData function to load the data
Public Sub LoadBinaryData(ByVal FileNum As Integer)
  Dim intX As Integer, intZ As Integer, _
    intY As Integer, sngH As Single
  GetBinaryData FileNum, TexID, sngH, intX, intZ, intY
  X = intX / 256
  Z = intZ / 256
  Y = intY / 256
  Height = CInt(sngH)
End Sub

' Simulates = operator
Public Sub CopyTo(Destination As clsLevel)
  With Destination
    .TexID = TexID
    .Height = Height
    .X = X
    .Y = Y
    .Z = Z
  End With
End Sub

' Resets the default values
Public Sub SetDefault()
  TexID = 0
  Height = 0
  X = 1
  Z = 1
  Y = 1
End Sub

