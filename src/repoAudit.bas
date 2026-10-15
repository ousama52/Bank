Attribute VB_Name = "repoAudit"
Option Explicit

' tblAudit columns (6): AuditID, Timestamp, UserID, Action, Target, Detail

' Adds an audit row to the unit of work (written with everything else when it commits).
Public Sub Add(ByVal uow As clsUnitOfWork, ByVal action As String, ByVal target As String, ByVal detail As String)
    Dim r(1 To 6) As Variant
    r(1) = uow.NewId(modConstants.TBL_AUDIT, "AuditID", "AU", 6)
    r(2) = modDateUtils.TimeStamp()
    r(3) = uow.UserID
    r(4) = action
    r(5) = target
    r(6) = detail
    uow.AppendRow modConstants.TBL_AUDIT, r
End Sub
