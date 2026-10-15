Attribute VB_Name = "svcPolicy"
Option Explicit

' Stage 8: central bank policy rate. Deposit rates get NEW rows in tblRateTiers (effective today); loan product rates change in place
' (existing loans keep their locked rate, fixed deposits keep theirs). Needs a HeadOffice approval (PolicyRate).

Private Function Clamp(ByVal v As Variant, ByVal lo As Variant, ByVal hi As Variant) As Variant
    If v < lo Then v = lo
    If v > hi Then v = hi
    Clamp = modMoney.RoundTo(v, 4)
End Function

Private Function RuleFor(ByVal rules As Variant, ByVal target As String, ByVal code As String, ByRef lo As Variant, ByRef hi As Variant) As Boolean
    Dim i As Long
    If IsEmpty(rules) Then Exit Function
    For i = 1 To UBound(rules, 1)
        If CStr(rules(i, 2)) = target And StrComp(CStr(rules(i, 3)), code, vbTextCompare) = 0 And CStr(rules(i, 4)) = modConstants.FLAG_YES Then
            lo = modMoney.ToDec(rules(i, 5))
            hi = modMoney.ToDec(rules(i, 6))
            RuleFor = True
            Exit Function
        End If
    Next i
End Function

' Control: nothing but Config.PolicyRate. Creates the approval.
Public Function ApplyPolicyRate(ByVal user As clsUser, ByVal inp As Object) As clsResult
    Dim uow As clsUnitOfWork, msg As String, r As clsResult, args As Object, delta As Variant
    msg = svcSecurity.Denied(user, svcSecurity.ACT_POLICY)
    If Len(msg) = 0 Then
        If Not modSheetUtils.TableExists(modConstants.TBL_RATERULES) Or Not repoConfig.Has("PolicyRate") Then msg = modMessages.MSG_SETUP_FIRST
    End If
    If Len(msg) > 0 Then Set ApplyPolicyRate = modMessages.FailResult(msg): Exit Function
    delta = repoConfig.GetDec("PolicyRate") - repoConfig.GetDec(modConstants.CFG_POLICY_APPLIED)
    If delta = 0 Then Set ApplyPolicyRate = modMessages.FailResult("The policy rate is already applied (no change since the last run)."): Exit Function
    Set uow = New clsUnitOfWork
    uow.Init repoConfig.BusinessDate(), user.UserID
    Set args = modArgs.NewArgs()
    args("policy") = CStr(repoConfig.GetDec("PolicyRate"))
    Set r = svcApproval.Request(uow, modConstants.REQ_POLICY_RATE, "", "", CDec(0), modConstants.HEAD_OFFICE_ID, args, user.UserID)
    uow.Commit
    Set ApplyPolicyRate = modMessages.OkResult("Policy rate change of " & Format$(CDbl(delta), "0.0000") & " requested. " & r.Message, 1, r.RefID)
End Function

' Run on approval: the change is computed again from the policy rate at that time.
Public Function ExecPolicy(ByVal uow As clsUnitOfWork, ByVal args As Object, ByVal createdBy As String, ByVal approvedBy As String) As clsResult
    Dim rules As Variant, tiers As Variant, lp As Variant, cfg As Variant, delta As Variant, policy As Variant, d As Date
    Dim i As Long, j As Long, lo As Variant, hi As Variant, newRate As Variant, seen As Object, key As String, best As Variant, bd As Date, td As Date, n As Long, detail As String, found As Boolean
    policy = repoConfig.GetDec("PolicyRate")
    delta = policy - repoConfig.GetDec(modConstants.CFG_POLICY_APPLIED)
    If delta = 0 Then Set ExecPolicy = modMessages.FailResult("The policy rate is already applied."): Exit Function
    d = uow.BusinessDate
    rules = modSheetUtils.ReadTable(modConstants.TBL_RATERULES)
    tiers = uow.WorkingTable(modConstants.TBL_RATETIERS)
    If Not IsEmpty(tiers) Then
        Set seen = CreateObject("Scripting.Dictionary")
        ' latest row per product, term and tier start
        For i = 1 To UBound(tiers, 1)
            If modDateUtils.TryToDate(tiers(i, 7), td) Then
                If td <= d Then
                    key = CStr(tiers(i, 2)) & "|" & CStr(tiers(i, 3)) & "|" & CStr(tiers(i, 4))
                    If Not seen.Exists(key) Then
                        seen(key) = i
                    ElseIf modDateUtils.TryToDate(tiers(seen(key), 7), bd) Then
                        If td >= bd Then seen(key) = i
                    End If
                End If
            End If
        Next i
        Dim k As Variant
        For Each k In seen.Keys
            i = seen(k)
            If RuleFor(rules, "Deposit", CStr(tiers(i, 2)), lo, hi) Then
                newRate = Clamp(modMoney.ToDec(tiers(i, 6)) + delta, lo, hi)
                If newRate <> modMoney.ToDec(tiers(i, 6)) Then
                    uow.AppendRow modConstants.TBL_RATETIERS, Array(uow.NewId(modConstants.TBL_RATETIERS, "RateID", "R", 4), CStr(tiers(i, 2)), tiers(i, 3), tiers(i, 4), tiers(i, 5), newRate, d)
                    n = n + 1
                    detail = detail & CStr(tiers(i, 2)) & "/" & CStr(tiers(i, 3)) & "m " & Format$(CDbl(tiers(i, 6)), "0.0000") & ">" & Format$(CDbl(newRate), "0.0000") & "; "
                End If
            End If
        Next k
    End If
    If modSheetUtils.TableExists(modConstants.TBL_LOANPRODUCTS) Then
        lp = uow.WorkingTable(modConstants.TBL_LOANPRODUCTS)
        If Not IsEmpty(lp) Then
            For j = 1 To UBound(lp, 1)
                If RuleFor(rules, "Loan", CStr(lp(j, 1)), lo, hi) Then
                    newRate = Clamp(modMoney.ToDec(lp(j, 4)) + delta, lo, hi)
                    If newRate <> modMoney.ToDec(lp(j, 4)) Then
                        detail = detail & "Loan " & CStr(lp(j, 1)) & " " & Format$(CDbl(lp(j, 4)), "0.0000") & ">" & Format$(CDbl(newRate), "0.0000") & "; "
                        lp(j, 4) = newRate
                        n = n + 1
                        found = True
                    End If
                End If
            Next j
            If found Then uow.ReplaceTable modConstants.TBL_LOANPRODUCTS, lp
        End If
    End If
    cfg = uow.WorkingTable(modConstants.TBL_CONFIG)
    found = False
    For i = 1 To UBound(cfg, 1)
        If StrComp(CStr(cfg(i, 1)), modConstants.CFG_POLICY_APPLIED, vbTextCompare) = 0 Then cfg(i, 2) = CStr(policy): found = True: Exit For
    Next i
    If found Then uow.ReplaceTable modConstants.TBL_CONFIG, cfg Else uow.AppendRow modConstants.TBL_CONFIG, Array(modConstants.CFG_POLICY_APPLIED, CStr(policy))
    repoAudit.Add uow, "PolicyRateApplied", CStr(policy), "Delta " & Format$(CDbl(delta), "0.0000") & ", " & n & " change(s): " & Left$(detail, 800)
    Set ExecPolicy = modMessages.OkResult("Policy rate " & Format$(CDbl(policy), "0.0000") & " applied: " & n & " rate change(s).", n, "")
End Function
