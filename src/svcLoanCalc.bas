Attribute VB_Name = "svcLoanCalc"
Option Explicit

' Loan schedule maths. Everything is Decimal; amounts round to whole MMK with modMoney.RoundCcy.
' A schedule is a 2D array (1-based): row n = installment, columns 1 = due date, 2 = principal, 3 = interest.

' Due date of installment k: k months after the disbursement date (clamped to the end of a shorter month).
Public Function DueDate(ByVal disburse As Date, ByVal k As Long) As Date
    DueDate = modDateUtils.AddMonths(disburse, k)
End Function

' (1 + r)^n in Decimal by repeated multiplication (n is at most a few hundred).
Private Function PowDec(ByVal onePlusR As Variant, ByVal n As Long) As Variant
    Dim v As Variant, i As Long
    v = CDec(1)
    For i = 1 To n
        v = v * onePlusR
    Next i
    PowDec = v
End Function

' Level payment of a declining balance loan: RoundCcy(P x r / (1 - (1 + r)^-n)) with r = annual rate / 12.
Public Function AnnuityPayment(ByVal principal As Variant, ByVal annualRate As Variant, ByVal n As Long) As Variant
    Dim r As Variant, p As Variant
    r = annualRate / 12
    If r = 0 Then
        AnnuityPayment = modMoney.RoundCcy(principal / n, modConstants.CCY_MMK)
        Exit Function
    End If
    p = PowDec(1 + r, n)
    AnnuityPayment = modMoney.RoundCcy(principal * r * p / (p - 1), modConstants.CCY_MMK)
End Function

' Flat: total interest = RoundCcy(P x rate x n / 12); installments 1..n-1 carry RoundCcy(P / n) principal and RoundCcy(total / n) interest,
' the last installment takes the remainders.
Private Function FlatSchedule(ByVal principal As Variant, ByVal rate As Variant, ByVal n As Long, ByVal disburse As Date) As Variant
    Dim s() As Variant, totalInt As Variant, pp As Variant, ip As Variant, i As Long, sumP As Variant, sumI As Variant
    ReDim s(1 To n, 1 To 3)
    totalInt = modMoney.RoundCcy(principal * rate * n / 12, modConstants.CCY_MMK)
    pp = modMoney.RoundCcy(principal / n, modConstants.CCY_MMK)
    ip = modMoney.RoundCcy(totalInt / n, modConstants.CCY_MMK)
    sumP = CDec(0)
    sumI = CDec(0)
    For i = 1 To n
        s(i, 1) = DueDate(disburse, i)
        If i < n Then
            s(i, 2) = pp
            s(i, 3) = ip
        Else
            s(i, 2) = principal - sumP
            s(i, 3) = totalInt - sumI
        End If
        sumP = sumP + s(i, 2)
        sumI = sumI + s(i, 3)
    Next i
    FlatSchedule = s
End Function

' Declining balance (annuity): installments 1..n-1 interest = RoundCcy(opening x r), principal = payment - interest;
' the last installment clears the remaining balance.
Private Function DecliningSchedule(ByVal principal As Variant, ByVal rate As Variant, ByVal n As Long, ByVal disburse As Date) As Variant
    Dim s() As Variant, pay As Variant, r As Variant, bal As Variant, i As Long, intr As Variant
    ReDim s(1 To n, 1 To 3)
    r = rate / 12
    pay = AnnuityPayment(principal, rate, n)
    bal = principal
    For i = 1 To n
        intr = modMoney.RoundCcy(bal * r, modConstants.CCY_MMK)
        s(i, 1) = DueDate(disburse, i)
        s(i, 3) = intr
        If i < n Then
            s(i, 2) = pay - intr
        Else
            s(i, 2) = bal
        End If
        bal = bal - s(i, 2)
    Next i
    DecliningSchedule = s
End Function

Public Function Build(ByVal principal As Variant, ByVal rate As Variant, ByVal method As String, ByVal n As Long, ByVal disburse As Date) As Variant
    If method = "Flat" Then
        Build = FlatSchedule(principal, rate, n, disburse)
    Else
        Build = DecliningSchedule(principal, rate, n, disburse)
    End If
End Function

' Recalculates the remaining installments after an extra principal payment (same number of installments).
' balance = principal still outstanding; startDue = due dates of the remaining rows are kept by the caller.
' Declining: new level payment for the new balance. Flat: principal spread evenly, interest per installment unchanged.
Public Function Reschedule(ByVal balance As Variant, ByVal rate As Variant, ByVal method As String, ByVal k As Long, ByVal flatInterest As Variant) As Variant
    Dim s() As Variant, i As Long, pp As Variant, sumP As Variant, pay As Variant, bal As Variant, intr As Variant, r As Variant
    ReDim s(1 To k, 1 To 2)
    If method = "Flat" Then
        pp = modMoney.RoundCcy(balance / k, modConstants.CCY_MMK)
        sumP = CDec(0)
        For i = 1 To k
            If i < k Then s(i, 1) = pp Else s(i, 1) = balance - sumP
            s(i, 2) = flatInterest
            sumP = sumP + s(i, 1)
        Next i
    Else
        r = rate / 12
        pay = AnnuityPayment(balance, rate, k)
        bal = balance
        For i = 1 To k
            intr = modMoney.RoundCcy(bal * r, modConstants.CCY_MMK)
            s(i, 2) = intr
            If i < k Then s(i, 1) = pay - intr Else s(i, 1) = bal
            bal = bal - s(i, 1)
        Next i
    End If
    Reschedule = s
End Function
