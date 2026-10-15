VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Used in the ErrHandler block of every entry macro:  clsErrorHandler.LogError "RunDeposit", Err

Public Sub LogError(ByVal procName As String, ByVal e As ErrObject)
    Dim num As Long, desc As String, src As String
    ' Copy first: logging itself may reset the Err object.
    num = e.Number
    desc = e.Description
    src = e.Source
    On Error Resume Next
    clsLogger.Fail procName, "Error " & num & ": " & desc & " (source: " & src & ")"
    Debug.Print "ERROR in " & procName & ": " & num & " " & desc
End Sub
