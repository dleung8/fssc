Attribute VB_Name = "basCDialog"
' [basCDialog]
' Helper module used for CommonDialog hook
Option Explicit

' (In basFSSC)
'Public cDialog As clsCDialog

' Call the class hook function
Public Function DialogHook(ByVal hDlg As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  If Not cDialog Is Nothing Then cDialog.clsDialogHook hDlg, Msg, wParam, lParam
End Function

