Attribute VB_Name = "modMessages"
Option Explicit

' Fixed message texts and the two clsResult builders used by every service.
' Later phases add their own messages here.

Public Const MSG_VAULT_SHORT As String = "Branch vault is short. Ask head office for a cash shipment."
Public Const MSG_FOREIGN_CCY As String = "This action works with MMK accounts only."
Public Const MSG_SETUP_FIRST As String = "A sheet or table is missing. Run RunSetup first."
Public Const MSG_NO_USER As String = "Choose an active user in Control C4."
Public Const MSG_BAD_AMOUNT As String = "Amount must be a whole number of MMK greater than zero."
Public Const MSG_SAME_ACCOUNT As String = "The two accounts must be different."
Public Const MSG_MAKER_CHECKER As String = "The approver must be a different user from the requester."
Public Const MSG_NO_BALANCE As String = "Balance is not enough."
Public Const MSG_MIN_BALANCE As String = "The balance after this would be below the minimum balance."
Public Const MSG_FD_NO_DEBIT As String = "A fixed deposit cannot be withdrawn or transferred out. Use maturity or early break."
Public Const MSG_EOD_ROLLED_BACK As String = "End of day failed a control check. Everything from this run was rolled back and the date was not advanced."

Public Function OkResult(ByVal message As String, Optional ByVal count As Long = 0, Optional ByVal refId As String = "") As clsResult
    Dim r As clsResult
    Set r = New clsResult
    r.Success = True
    r.Message = message
    r.Count = count
    r.RefID = refId
    Set OkResult = r
End Function

Public Function FailResult(ByVal message As String) As clsResult
    Dim r As clsResult
    Set r = New clsResult
    r.Success = False
    r.Message = message
    Set FailResult = r
End Function

Public Function UnknownAccount(ByVal accountNo As String) As String
    UnknownAccount = "Unknown account " & accountNo & "."
End Function

Public Function UnknownCustomer(ByVal customerId As String) As String
    UnknownCustomer = "Unknown customer " & customerId & "."
End Function

Public Function UnknownBranch(ByVal branchId As String) As String
    UnknownBranch = "Unknown branch " & branchId & "."
End Function

Public Function UnknownUser(ByVal userId As String) As String
    UnknownUser = "Unknown or inactive user " & userId & "."
End Function

Public Function AccountStatusRefuses(ByVal accountNo As String, ByVal status As String) As String
    AccountStatusRefuses = "Account " & accountNo & " is " & status & ". This action is refused."
End Function

Public Function RoleDenied(ByVal role As String, ByVal action As String) As String
    RoleDenied = "Role " & role & " is not allowed to " & action & "."
End Function

Public Function EndOfDayAlready(ByVal d As Date) As String
    EndOfDayAlready = "End of day for " & modDateUtils.ToIso(d) & " was already run."
End Function

Public Function PendingWarning(ByVal n As Long) As String
    PendingWarning = n & " approval(s) are still Pending."
End Function
