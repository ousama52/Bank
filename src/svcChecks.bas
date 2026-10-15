Attribute VB_Name = "svcChecks"
Option Explicit

' Control checks. RunChecks returns a 2D array (1-based): number, check, result (PASS / FAIL / WARN), detail.
' End of day runs it and rolls the day back when any check FAILs. Later stages add checks in svcChecksExtra.

Public Sub AddRow(ByVal rowsList As Collection, ByVal num As Long, ByVal name As String, ByVal status As String, ByVal detail As String)
    rowsList.Add Array(num, name, status, detail)
End Sub

Private Function Verdict(ByVal bad As Long, ByVal okText As String, ByVal badText As String) As String
    If bad = 0 Then Verdict = okText Else Verdict = badText
End Function

Private Function PassFail(ByVal bad As Long) As String
    If bad = 0 Then PassFail = "PASS" Else PassFail = "FAIL"
End Function

Private Sub AddTo(ByVal d As Object, ByVal key As String, ByVal v As Variant)
    If d.Exists(key) Then d(key) = d(key) + v Else d(key) = v
End Sub

' Checks 1 to 3, all from the journal lines.
Private Sub CheckJournals(ByVal rowsList As Collection, ByVal lns As Variant)
    Dim perJ As Object, perJB As Object, i As Long, net As Variant, j As String, b As String, k As Variant
    Dim bad1 As Long, bad2 As Long, gl1900 As Variant, sample1 As String, sample2 As String
    Set perJ = CreateObject("Scripting.Dictionary")
    Set perJB = CreateObject("Scripting.Dictionary")
    gl1900 = CDec(0)
    If Not IsEmpty(lns) Then
        For i = 1 To UBound(lns, 1)
            net = modMoney.ToDec(lns(i, repoJournal.LNC_DEBIT)) - modMoney.ToDec(lns(i, repoJournal.LNC_CREDIT))
            j = CStr(lns(i, repoJournal.LNC_JOURNAL))
            b = CStr(lns(i, repoJournal.LNC_BRANCH))
            AddTo perJ, j, net
            AddTo perJB, j & "|" & b, net
            If CStr(lns(i, repoJournal.LNC_GL)) = modConstants.GL_INTERBRANCH Then gl1900 = gl1900 + net
        Next i
    End If
    For Each k In perJ.Keys
        If perJ(k) <> 0 Then
            bad1 = bad1 + 1
            If Len(sample1) < 60 Then sample1 = sample1 & CStr(k) & " "
        End If
    Next k
    For Each k In perJB.Keys
        If perJB(k) <> 0 Then
            bad2 = bad2 + 1
            If Len(sample2) < 60 Then sample2 = sample2 & CStr(k) & " "
        End If
    Next k
    AddRow rowsList, 1, "Every journal: total debit = total credit", PassFail(bad1), Verdict(bad1, perJ.Count & " journals checked", bad1 & " unbalanced: " & sample1)
    AddRow rowsList, 2, "Every journal, every branch: debit = credit", PassFail(bad2), Verdict(bad2, perJB.Count & " journal-branch pairs checked", bad2 & " unbalanced: " & sample2)
    AddRow rowsList, 3, "GL 1900 inter-branch clearing nets to 0", PassFail(IIf(gl1900 = 0, 0, 1)), "Net " & modMoney.FormatAmount(gl1900)
End Sub

' Check 4: account balance = statement rows = journal lines, and deposit GL totals = sum of balances.
Private Sub CheckAccounts(ByVal rowsList As Collection, ByVal lns As Variant)
    Dim accts As Object, txn As Variant, perT As Object, perL As Object, glNet As Object, i As Long, k As Variant, a As clsAccount
    Dim bad As Long, sample As String, sums As Object, gl As String, badGL As Long, detailGL As String
    Dim glv As Variant, s As Variant, n As Variant, t As Variant, l As Variant, isFx As Boolean
    Set accts = repoAccount.LoadAccounts()
    Set perT = CreateObject("Scripting.Dictionary")
    Set perL = CreateObject("Scripting.Dictionary")
    Set glNet = CreateObject("Scripting.Dictionary")
    Set sums = CreateObject("Scripting.Dictionary")
    txn = repoTransaction.LoadTransactions()
    If Not IsEmpty(txn) Then
        For i = 1 To UBound(txn, 1)
            If CStr(txn(i, 6)) = modConstants.DIR_CREDIT Then
                AddTo perT, CStr(txn(i, 4)), modMoney.ToDec(txn(i, 7))
            Else
                AddTo perT, CStr(txn(i, 4)), -modMoney.ToDec(txn(i, 7))
            End If
        Next i
    End If
    If Not IsEmpty(lns) Then
        For i = 1 To UBound(lns, 1)
            gl = CStr(lns(i, repoJournal.LNC_GL))
            isFx = (Len(CStr(lns(i, 6))) > 0 And CStr(lns(i, 6)) <> modConstants.CCY_MMK)
            If Not isFx And (gl = modConstants.GL_SAVINGS Or gl = modConstants.GL_CURRENT Or gl = modConstants.GL_FIXED) Then
                AddTo glNet, gl, modMoney.ToDec(lns(i, repoJournal.LNC_CREDIT)) - modMoney.ToDec(lns(i, repoJournal.LNC_DEBIT))
            End If
            If Len(CStr(lns(i, repoJournal.LNC_ACCOUNT))) > 0 Then
                If isFx Then
                    If modMoney.ToDec(lns(i, repoJournal.LNC_CREDIT)) > 0 Then
                        AddTo perL, CStr(lns(i, repoJournal.LNC_ACCOUNT)), modMoney.ToDec(lns(i, 10))
                    Else
                        AddTo perL, CStr(lns(i, repoJournal.LNC_ACCOUNT)), -modMoney.ToDec(lns(i, 10))
                    End If
                Else
                    AddTo perL, CStr(lns(i, repoJournal.LNC_ACCOUNT)), modMoney.ToDec(lns(i, repoJournal.LNC_CREDIT)) - modMoney.ToDec(lns(i, repoJournal.LNC_DEBIT))
                End If
            End If
        Next i
    End If
    For Each k In accts.Keys
        Set a = accts(k)
        If Len(a.DepositGL) > 0 And a.Ccy = modConstants.CCY_MMK Then AddTo sums, a.DepositGL, a.Balance
        t = CDec(0)
        l = CDec(0)
        If perT.Exists(CStr(k)) Then t = perT(CStr(k))
        If perL.Exists(CStr(k)) Then l = perL(CStr(k))
        If a.Balance <> t Or a.Balance <> l Then
            bad = bad + 1
            If Len(sample) < 60 Then sample = sample & CStr(k) & " "
        End If
    Next k
    AddRow rowsList, 4, "Each account: Balance = statement rows = journal lines", PassFail(bad), Verdict(bad, accts.Count & " accounts checked", bad & " differ: " & sample)
    For Each glv In Array(modConstants.GL_SAVINGS, modConstants.GL_CURRENT, modConstants.GL_FIXED)
        gl = CStr(glv)
        s = CDec(0)
        n = CDec(0)
        If sums.Exists(gl) Then s = sums(gl)
        If glNet.Exists(gl) Then n = glNet(gl)
        If s <> n Then badGL = badGL + 1
        detailGL = detailGL & gl & ": balances " & modMoney.FormatAmount(s) & " / GL " & modMoney.FormatAmount(n) & "; "
    Next glv
    AddRow rowsList, 4, "Each deposit GL = total of account balances", PassFail(badGL), detailGL
End Sub

' Check 5: GL 2900 = total of Submitted outgoing transfers.
Private Sub CheckInterbank(ByVal rowsList As Collection, ByVal lns As Variant)
    Dim gl As Variant, i As Long, ib As Variant, sub1 As Variant
    gl = CDec(0)
    sub1 = CDec(0)
    If Not IsEmpty(lns) Then
        For i = 1 To UBound(lns, 1)
            If CStr(lns(i, repoJournal.LNC_GL)) = modConstants.GL_IB_PAYABLE Then gl = gl + modMoney.ToDec(lns(i, repoJournal.LNC_CREDIT)) - modMoney.ToDec(lns(i, repoJournal.LNC_DEBIT))
        Next i
    End If
    ib = repoInterbank.LoadTransfers()
    If Not IsEmpty(ib) Then
        For i = 1 To UBound(ib, 1)
            If CStr(ib(i, repoInterbank.IBC_DIR)) = modConstants.DIR_OUT And CStr(ib(i, repoInterbank.IBC_STATUS)) = modConstants.STATUS_SUBMITTED Then
                sub1 = sub1 + modMoney.ToDec(ib(i, repoInterbank.IBC_AMOUNT))
            End If
        Next i
    End If
    AddRow rowsList, 5, "GL 2900 = total of Submitted outgoing transfers", PassFail(IIf(gl = sub1, 0, 1)), "GL " & modMoney.FormatAmount(gl) & " / Submitted " & modMoney.FormatAmount(sub1)
End Sub

' Check 6: no vault below 0 (FAIL); below VaultMin is only a warning.
Private Sub CheckVaults(ByVal rowsList As Collection, ByVal lns As Variant)
    Dim net As Object, i As Long, branches As Object, k As Variant, b As clsBranch, v As Variant, bad As Long, warn As Long, info As String
    Set net = CreateObject("Scripting.Dictionary")
    If Not IsEmpty(lns) Then
        For i = 1 To UBound(lns, 1)
            If CStr(lns(i, repoJournal.LNC_GL)) = modConstants.GL_VAULT Then
                AddTo net, CStr(lns(i, repoJournal.LNC_BRANCH)), modMoney.ToDec(lns(i, repoJournal.LNC_DEBIT)) - modMoney.ToDec(lns(i, repoJournal.LNC_CREDIT))
            End If
        Next i
    End If
    Set branches = repoBranch.LoadBranches()
    For Each k In branches.Keys
        Set b = branches(k)
        v = CDec(0)
        If net.Exists(CStr(k)) Then v = net(CStr(k))
        If v < 0 Then
            bad = bad + 1
            info = info & CStr(k) & " below 0; "
        ElseIf v < b.VaultMin Then
            warn = warn + 1
        End If
    Next k
    If bad > 0 Then
        AddRow rowsList, 6, "No branch vault below 0", "FAIL", info
    ElseIf warn > 0 Then
        AddRow rowsList, 6, "No branch vault below 0", "WARN", "No vault is negative, but " & warn & " branch(es) are below VaultMin"
    Else
        AddRow rowsList, 6, "No branch vault below 0", "PASS", branches.Count & " vaults checked"
    End If
End Sub

Public Function RunChecks() As Variant
    Dim rowsList As Collection, lns As Variant, out() As Variant, i As Long, c As Long
    Set rowsList = New Collection
    lns = repoJournal.LoadLines()
    CheckJournals rowsList, lns
    CheckAccounts rowsList, lns
    CheckInterbank rowsList, lns
    CheckVaults rowsList, lns
    svcChecksExtra.RunExtraChecks rowsList, lns
    ReDim out(1 To rowsList.Count, 1 To 4)
    For i = 1 To rowsList.Count
        For c = 0 To 3
            out(i, c + 1) = rowsList(i)(c)
        Next c
    Next i
    RunChecks = out
End Function

Public Function FailCount(ByVal results As Variant) As Long
    Dim i As Long
    For i = 1 To UBound(results, 1)
        If results(i, 3) = "FAIL" Then FailCount = FailCount + 1
    Next i
End Function

Public Function FailText(ByVal results As Variant) As String
    Dim i As Long
    For i = 1 To UBound(results, 1)
        If results(i, 3) = "FAIL" Then FailText = FailText & "#" & results(i, 1) & " " & results(i, 2) & " (" & results(i, 4) & "); "
    Next i
End Function
