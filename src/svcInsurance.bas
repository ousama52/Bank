Attribute VB_Name = "svcInsurance"
Option Explicit

' Stage 8: deposit insurance premium. At every quarter end the bank books an expense (GL 5340) on the insured deposits.
' Insured = min(customer's deposits and wallet balances in MMK, DepositInsuranceCap). Once per quarter.

' Deposit and wallet balances per customer in MMK (closed accounts have no balance).
Public Function CustomerTotals(ByVal uow As clsUnitOfWork) As Object
    Dim res As Object, k As Variant, a As clsAccount, v As Variant, mid As Variant, buy As Variant, sell As Variant, msg As String
    Set res = CreateObject("Scripting.Dictionary")
    For Each k In uow.Accounts().Keys
        Set a = uow.Accounts()(k)
        If a.Status <> modConstants.STATUS_CLOSED And a.Balance > 0 Then
            v = a.Balance
            If a.Ccy <> modConstants.CCY_MMK Then
                If svcFx.MidFor(uow, a.Ccy, mid, buy, sell, msg) Then v = modMoney.RoundCcy(a.Balance * mid, modConstants.CCY_MMK) Else v = CDec(0)
            End If
            If res.Exists(a.CustomerID) Then res(a.CustomerID) = res(a.CustomerID) + v Else res(a.CustomerID) = v
        End If
    Next k
    Set CustomerTotals = res
End Function

' Total insured amount of all customers.
Public Function TotalInsured(ByVal uow As clsUnitOfWork) As Variant
    Dim tot As Object, k As Variant, cap As Variant, s As Variant
    Set tot = CustomerTotals(uow)
    cap = repoConfig.GetDec(modConstants.CFG_INS_CAP)
    s = CDec(0)
    For Each k In tot.Keys
        If tot(k) > cap Then s = s + cap Else s = s + tot(k)
    Next k
    TotalInsured = s
End Function

Private Function QuarterLabel(ByVal d As Date) As String
    QuarterLabel = "Deposit insurance " & Year(d) & "Q" & ((Month(d) - 1) \ 3 + 1)
End Function

' Called at month end. Does nothing unless d is a quarter end.
Public Function QuarterEnd(ByVal uow As clsUnitOfWork, ByVal d As Date, ByVal stats As Object) As clsResult
    Dim ex As Variant, i As Long, label As String, prem As Variant, insured As Variant, lines As Collection, jr As clsResult, id As String, ho As String
    ho = modConstants.HEAD_OFFICE_ID
    If Month(d) Mod 3 <> 0 Or Not modDateUtils.IsLastDayOfMonth(d) Then Set QuarterEnd = modMessages.OkResult("", 0): Exit Function
    If Not modSheetUtils.TableExists(modConstants.TBL_EXPENSES) Or Not repoConfig.Has(modConstants.CFG_INS_RATE) Then Set QuarterEnd = modMessages.OkResult("", 0): Exit Function
    label = QuarterLabel(d)
    ex = uow.WorkingTable(modConstants.TBL_EXPENSES)
    If Not IsEmpty(ex) Then
        For i = 1 To UBound(ex, 1)
            If CStr(ex(i, 7)) = label Then Set QuarterEnd = modMessages.OkResult("", 0): Exit Function
        Next i
    End If
    insured = TotalInsured(uow)
    prem = modMoney.RoundCcy(insured * repoConfig.GetDec(modConstants.CFG_INS_RATE) / 4, modConstants.CCY_MMK)
    If prem <= 0 Then Set QuarterEnd = modMessages.OkResult("", 0): Exit Function
    Set lines = New Collection
    svcLedger.AddDr lines, ho, modConstants.GL_INS_PREMIUM, prem
    svcLedger.AddCr lines, ho, modConstants.GL_ACCRUED, prem
    Set jr = svcLedger.PostJournal(uow, modConstants.JT_EXPENSE, label, uow.NewGroupId(), lines)
    If Not jr.Success Then Set QuarterEnd = jr: Exit Function
    id = uow.NewId(modConstants.TBL_EXPENSES, "ExpenseID", "EX", 6)
    uow.AppendRow modConstants.TBL_EXPENSES, repoCost.ExpenseRow(id, ho, d, modConstants.GL_INS_PREMIUM, prem, "Deposit Insurance Fund", label, "HeadOffice", "Posted", "SYSTEM", "", jr.RefID, "")
    repoAudit.Add uow, "DepositInsurance", id, label & ": insured " & modMoney.FormatAmount(insured) & ", premium " & modMoney.FormatAmount(prem)
    stats("InsurancePremium") = prem
    Set QuarterEnd = modMessages.OkResult("", 1)
End Function
