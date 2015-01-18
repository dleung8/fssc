VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GSubclass"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub AttachMessage(iwp As ISubclass, ByVal hwnd As Long, ByVal iMsg As Long)
Attribute AttachMessage.VB_Description = "Start subclassing for a particular message."
  MSubclass.AttachMessage iwp, hwnd, iMsg
End Sub

Public Function CallOldWindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  CallOldWindowProc = MSubclass.CallOldWindowProc(hwnd, iMsg, wParam, lParam)
End Function

Public Property Get CurrentMessage() As Long
  CurrentMessage = MSubclass.CurrentMessage
End Property

Public Sub DetachMessage(iwp As ISubclass, ByVal hwnd As Long, ByVal iMsg As Long)
Attribute DetachMessage.VB_Description = "Stop subclassing for a particular message."
  MSubclass.DetachMessage iwp, hwnd, iMsg
End Sub

Public Sub AdminMessage(iwp As ISubclass, ByVal hwnd As Long, ByVal iMsg As Long, ByVal Attach As Boolean)
  If Attach Then
    MSubclass.AttachMessage iwp, hwnd, iMsg
  Else
    MSubclass.DetachMessage iwp, hwnd, iMsg
  End If
End Sub
