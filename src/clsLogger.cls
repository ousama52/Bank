VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Writes tblLog rows. Call it as clsLogger.Info "proc", "text" (predeclared instance).
' A log write must never break the action that is being logged, so every failure is swallowed.

Public Sub Info(ByVal procName As String, ByVal message As String)
    WriteLine modConstants.LOG_INFO, procName, message
End Sub

Public Sub Warn(ByVal procName As String, ByVal message As String)
    WriteLine modConstants.LOG_WARN, procName, message
End Sub

Public Sub Fail(ByVal procName As String, ByVal message As String)
    WriteLine modConstants.LOG_ERROR, procName, message
End Sub

Private Sub WriteLine(ByVal level As String, ByVal procName As String, ByVal message As String)
    Dim entry(1 To 1, 1 To 4) As Variant
    On Error GoTo Failed
    entry(1, 1) = modDateUtils.TimeStamp()
    entry(1, 2) = level
    entry(1, 3) = procName
    entry(1, 4) = message
    modSheetUtils.AppendRows modConstants.TBL_LOG, entry
    Exit Sub
Failed:
    Debug.Print level & " | " & procName & " | " & message
End Sub
