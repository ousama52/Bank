Attribute VB_Name = "repoAccount"
Option Explicit

' Columns of tblAccounts (16): AccountNo, CustomerID, BranchID, ProductCode, Currency, Status, OpenDate, TermMonths,
' MaturityDate, LockedRate, PayoutAccountNo, Balance, AccruedInterest, LastAccrualDate, LastActivityDate, ClosedDate
Private Const NCOLS As Long = 16

' An optional Decimal: blank stays Empty.
Private Function OptDec(ByVal v As Variant) As Variant
    If Len(CStr(v)) = 0 Then Exit Function
    OptDec = modMoney.ToDec(v)
End Function

Private Function FromRow(ByVal data As Variant, ByVal i As Long, ByVal prods As Object) As clsAccount
    Dim a As clsAccount, p As clsProduct
    Set a = New clsAccount
    a.AccountNo = CStr(data(i, 1))
    a.CustomerID = CStr(data(i, 2))
    a.BranchID = CStr(data(i, 3))
    a.ProductCode = CStr(data(i, 4))
    a.Ccy = CStr(data(i, 5))
    a.Status = CStr(data(i, 6))
    a.OpenDate = modDateUtils.ToDateOrEmpty(data(i, 7))
    a.TermMonths = CLng(modMoney.ToDec(data(i, 8)))
    a.MaturityDate = modDateUtils.ToDateOrEmpty(data(i, 9))
    a.LockedRate = OptDec(data(i, 10))
    a.PayoutAccountNo = CStr(data(i, 11))
    a.Balance = modMoney.ToDec(data(i, 12))
    a.AccruedInterest = modMoney.ToDec(data(i, 13))
    a.LastAccrualDate = modDateUtils.ToDateOrEmpty(data(i, 14))
    a.LastActivityDate = modDateUtils.ToDateOrEmpty(data(i, 15))
    a.ClosedDate = modDateUtils.ToDateOrEmpty(data(i, 16))
    If prods.Exists(a.ProductCode) Then
        Set p = prods(a.ProductCode)
        a.ProductType = p.ProductType
        a.DepositGL = repoProduct.DepositGL(p.ProductType)
        a.MinBalance = p.MinBalance
    Else
        a.MinBalance = CDec(0)
    End If
    Set FromRow = a
End Function

Private Sub PutRow(ByRef out As Variant, ByVal r As Long, ByVal a As clsAccount)
    out(r, 1) = a.AccountNo
    out(r, 2) = a.CustomerID
    out(r, 3) = a.BranchID
    out(r, 4) = a.ProductCode
    out(r, 5) = a.Ccy
    out(r, 6) = a.Status
    out(r, 7) = a.OpenDate
    If a.TermMonths > 0 Then out(r, 8) = a.TermMonths Else out(r, 8) = Empty
    out(r, 9) = a.MaturityDate
    out(r, 10) = a.LockedRate
    If Len(a.PayoutAccountNo) > 0 Then out(r, 11) = a.PayoutAccountNo Else out(r, 11) = Empty
    out(r, 12) = a.Balance
    out(r, 13) = a.AccruedInterest
    out(r, 14) = a.LastAccrualDate
    out(r, 15) = a.LastActivityDate
    out(r, 16) = a.ClosedDate
End Sub

' Dictionary AccountNo -> clsAccount, in table order.
Public Function LoadAccounts() As Object
    Dim data As Variant, d As Object, i As Long, a As clsAccount, prods As Object
    Set d = CreateObject("Scripting.Dictionary")
    Set prods = repoProduct.LoadProducts()
    data = modSheetUtils.ReadTable(modConstants.TBL_ACCOUNTS)
    If Not IsEmpty(data) Then
        For i = 1 To UBound(data, 1)
            If Len(CStr(data(i, 1))) > 0 Then
                Set a = FromRow(data, i, prods)
                d.Add a.AccountNo, a
            End If
        Next i
    End If
    Set LoadAccounts = d
End Function

' Writes the whole table back in one assignment when anything changed (rows keep their order, new accounts go last).
Public Sub Save(ByVal accts As Object)
    Dim data As Variant, n As Long, i As Long, c As Long, k As Variant, a As clsAccount
    Dim anyChange As Boolean, newCount As Long, out As Variant, r As Long, key As String
    For Each k In accts.Keys
        Set a = accts(k)
        If a.Dirty Or a.IsNew Then anyChange = True
        If a.IsNew Then newCount = newCount + 1
    Next k
    If Not anyChange Then Exit Sub
    data = modSheetUtils.ReadTable(modConstants.TBL_ACCOUNTS)
    If Not IsEmpty(data) Then n = UBound(data, 1)
    ReDim out(1 To n + newCount, 1 To NCOLS)
    For i = 1 To n
        key = CStr(data(i, 1))
        If accts.Exists(key) Then
            Set a = accts(key)
            If a.Dirty Then
                PutRow out, i, a
            Else
                For c = 1 To NCOLS
                    out(i, c) = data(i, c)
                Next c
            End If
        Else
            For c = 1 To NCOLS
                out(i, c) = data(i, c)
            Next c
        End If
    Next i
    r = n
    For Each k In accts.Keys
        Set a = accts(k)
        If a.IsNew Then
            r = r + 1
            PutRow out, r, a
        End If
    Next k
    modSheetUtils.WriteTable modConstants.TBL_ACCOUNTS, out
    For Each k In accts.Keys
        Set a = accts(k)
        a.Dirty = False
        a.IsNew = False
    Next k
End Sub
