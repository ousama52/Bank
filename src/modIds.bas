Attribute VB_Name = "modIds"
Option Explicit

' IDs are always "highest existing number + 1" (never reuse). A unit of work that needs many IDs
' reads the highest number once with a Max... function and counts up in memory.

Public Function FormatId(ByVal prefix As String, ByVal n As Long, ByVal width As Long) As String
    FormatId = prefix & Format$(n, String$(width, "0"))
End Function

' Highest number among keys that are `prefix` followed by digits only. 0 when there are none.
Public Function MaxNumber(ByVal tableName As String, ByVal keyColumn As String, ByVal prefix As String) As Long
    Dim col As Variant, i As Long, k As String, rest As String, n As Long, best As Long
    col = modSheetUtils.ReadColumn(tableName, keyColumn)
    If IsEmpty(col) Then Exit Function
    For i = 1 To UBound(col, 1)
        k = CStr(col(i, 1))
        If Left$(k, Len(prefix)) = prefix Then
            rest = Mid$(k, Len(prefix) + 1)
            If Len(rest) > 0 And Len(rest) <= 9 And Not (rest Like "*[!0-9]*") Then
                n = CLng(rest)
                If n > best Then best = n
            End If
        End If
    Next i
    MaxNumber = best
End Function

Public Function NextId(ByVal tableName As String, ByVal keyColumn As String, ByVal prefix As String, ByVal width As Long) As String
    NextId = FormatId(prefix, MaxNumber(tableName, keyColumn, prefix) + 1, width)
End Function

' ---- Single IDs ----

Public Function NextCustomerId() As String
    NextCustomerId = NextId(modConstants.TBL_CUSTOMERS, "CustomerID", "C", 6)
End Function

Public Function NextTransferId() As String
    NextTransferId = NextId(modConstants.TBL_INTERBANK, "TransferID", "IB", 6)
End Function

Public Function NextApprovalId() As String
    NextApprovalId = NextId(modConstants.TBL_APPROVALS, "ApprovalID", "AP", 6)
End Function

Public Function NextAuditId() As String
    NextAuditId = NextId(modConstants.TBL_AUDIT, "AuditID", "AU", 6)
End Function

' ---- Journals and lines (bulk: MaxJournalNo once, then JournalId(max + 1), JournalId(max + 2) ...) ----

Public Function MaxJournalNo() As Long
    MaxJournalNo = MaxNumber(modConstants.TBL_JOURNAL, "JournalID", "J")
End Function

Public Function JournalId(ByVal n As Long) As String
    JournalId = FormatId("J", n, 8)
End Function

Public Function MaxLineNo() As Long
    MaxLineNo = MaxNumber(modConstants.TBL_JLINES, "LineID", "L")
End Function

Public Function LineId(ByVal n As Long) As String
    LineId = FormatId("L", n, 7)
End Function

' ---- Transactions and groups: the sequence restarts every business date ----

Public Function MaxTxnNo(ByVal d As Date) As Long
    MaxTxnNo = MaxNumber(modConstants.TBL_TRANSACTIONS, "TxnID", "TX" & modDateUtils.DateKey(d) & "-")
End Function

Public Function TxnId(ByVal d As Date, ByVal n As Long) As String
    TxnId = FormatId("TX" & modDateUtils.DateKey(d) & "-", n, 5)
End Function

' GroupID is stored in Transactions, Journal and Approvals; the highest of the first two counts.
Public Function MaxGroupNo(ByVal d As Date) As Long
    Dim a As Long, b As Long, p As String
    p = "G" & modDateUtils.DateKey(d) & "-"
    a = MaxNumber(modConstants.TBL_TRANSACTIONS, "GroupID", p)
    b = MaxNumber(modConstants.TBL_JOURNAL, "GroupID", p)
    If a > b Then MaxGroupNo = a Else MaxGroupNo = b
End Function

Public Function GroupId(ByVal d As Date, ByVal n As Long) As String
    GroupId = FormatId("G" & modDateUtils.DateKey(d) & "-", n, 5)
End Function

' ---- Account numbers: 3 digit branch number + 7 digit serial. The serial runs over the whole bank. ----

Public Function MaxAccountSerial() As Long
    Dim col As Variant, i As Long, k As String, n As Long, best As Long
    col = modSheetUtils.ReadColumn(modConstants.TBL_ACCOUNTS, "AccountNo")
    If IsEmpty(col) Then Exit Function
    For i = 1 To UBound(col, 1)
        k = CStr(col(i, 1))
        If Len(k) = 10 And Not (k Like "*[!0-9]*") Then
            n = CLng(Right$(k, 7))
            If n > best Then best = n
        End If
    Next i
    MaxAccountSerial = best
End Function

Public Function AccountNoFor(ByVal branchId As String, ByVal serial As Long) As String
    Dim num As String
    num = Mid$(branchId, 3)
    If UCase$(Left$(branchId, 2)) <> "BR" Or Len(num) = 0 Or num Like "*[!0-9]*" Then
        Err.Raise vbObjectError + 1010, "modIds.AccountNoFor", "Bad branch ID: " & branchId
    End If
    AccountNoFor = Format$(CLng(num), "000") & Format$(serial, "0000000")
End Function

Public Function NextAccountNo(ByVal branchId As String) As String
    NextAccountNo = AccountNoFor(branchId, MaxAccountSerial() + 1)
End Function
