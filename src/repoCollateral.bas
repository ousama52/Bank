Attribute VB_Name = "repoCollateral"
Option Explicit

' Column positions and row builders for the Stage 4 tables (read and changed through clsUnitOfWork.WorkingTable).

' tblCollateralTypes (8): TypeCode, Name, MaxLTV, Haircut, AnnualValueChange, RevalueMonths, SellingCostPct, MonthlyHoldingCostPct
Public Const CTC_CODE As Long = 1
Public Const CTC_MAXLTV As Long = 3
Public Const CTC_HAIRCUT As Long = 4
Public Const CTC_VALUECHANGE As Long = 5
Public Const CTC_REVALUE As Long = 6
Public Const CTC_SELLCOST As Long = 7
Public Const CTC_HOLDCOST As Long = 8

' tblCollateral (12): CollateralID, CustomerID, TypeCode, Description, Location, BranchID, Custody, LoanID, LienStatus, RegisteredDate, LatestValue, LatestValueDate
Public Const COC_ID As Long = 1
Public Const COC_CUSTOMER As Long = 2
Public Const COC_TYPE As Long = 3
Public Const COC_DESC As Long = 4
Public Const COC_BRANCH As Long = 6
Public Const COC_CUSTODY As Long = 7
Public Const COC_LOAN As Long = 8
Public Const COC_LIEN As Long = 9
Public Const COC_VALUE As Long = 11
Public Const COC_VALUEDATE As Long = 12

Public Const LIEN_FREE As String = "Free"
Public Const LIEN_PLEDGED As String = "Pledged"
Public Const LIEN_RELEASED As String = "Released"
Public Const LIEN_REPOSSESSED As String = "Repossessed"
Public Const LIEN_SOLD As String = "Sold"
Public Const LIEN_REDEEMED As String = "Redeemed"

' tblAppraisals (7): AppraisalID, CollateralID, AppraisalDate, AppraisedValue, Appraiser, Purpose, Note
Public Const APPR_COLLATERAL As Long = 2
Public Const APPR_DATE As Long = 3
Public Const APPR_VALUE As Long = 4
Public Const APPR_PURPOSE As Long = 6

' tblRepossessions (16): RepoID, LoanID, CollateralID, HeldBranchID, NoticeDate, RepoDate, PrincipalAtRepo, InterestAtRepo, ProvisionUsed,
' ClaimTotal, FairValueLessCosts, CarryingValue, OriginalCarryingValue, Deficiency, Status, ClosedDate
Public Const RPC_ID As Long = 1
Public Const RPC_LOAN As Long = 2
Public Const RPC_COLLATERAL As Long = 3
Public Const RPC_BRANCH As Long = 4
Public Const RPC_NOTICE As Long = 5
Public Const RPC_REPODATE As Long = 6
Public Const RPC_PRINCIPAL As Long = 7
Public Const RPC_INTEREST As Long = 8
Public Const RPC_PROVISION As Long = 9
Public Const RPC_CLAIM As Long = 10
Public Const RPC_FVLC As Long = 11
Public Const RPC_CARRYING As Long = 12
Public Const RPC_ORIGINAL As Long = 13
Public Const RPC_DEFICIENCY As Long = 14
Public Const RPC_STATUS As Long = 15
Public Const RPC_CLOSED As Long = 16

Public Const RS_NOTICE As String = "Notice"
Public Const RS_HELD As String = "Held"
Public Const RS_SOLD As String = "Sold"
Public Const RS_RENTED As String = "Rented"
Public Const RS_INBUSINESS As String = "InBusiness"
Public Const RS_REDEEMED As String = "Redeemed"
Public Const RS_CLOSED As String = "Closed"

' tblAssetSales (12): SaleID, RepoID, SaleDate, Buyer, SalePrice, SellingCosts, NetProceeds, AppliedToClaim, SurplusToCustomer, CarryingValueAtSale, RecoveryOrLoss, JournalID
' tblAssetRentals (7): RentalID, RepoID, Tenant, MonthlyRent, StartDate, EndDate, Status
' tblAssetOps (7): OpID, RepoID, OpDate, OpType, Amount, Note, JournalID

Public Function FindCollateralRow(ByVal data As Variant, ByVal idCol As Long, ByVal id As String) As Long
    Dim i As Long
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If StrComp(CStr(data(i, idCol)), id, vbTextCompare) = 0 Then
            FindCollateralRow = i
            Exit Function
        End If
    Next i
End Function

' Collateral type as a 1D array (copy of the table row), or Empty.
Public Function TypeRow(ByVal code As String) As Variant
    Dim data As Variant, i As Long, c As Long, r(1 To 8) As Variant
    If Not modSheetUtils.TableExists(modConstants.TBL_COLLTYPES) Then Exit Function
    data = modSheetUtils.ReadTable(modConstants.TBL_COLLTYPES)
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If StrComp(CStr(data(i, CTC_CODE)), code, vbTextCompare) = 0 Then
            For c = 1 To 8
                r(c) = data(i, c)
            Next c
            TypeRow = r
            Exit Function
        End If
    Next i
End Function

' Value x (1 - haircut) x (1 - selling cost): what the bank expects to get after costs.
Public Function FairValueLessCosts(ByVal value As Variant, ByVal trow As Variant) As Variant
    FairValueLessCosts = modMoney.RoundCcy(value * (1 - modMoney.ToDec(trow(CTC_HAIRCUT))) * (1 - modMoney.ToDec(trow(CTC_SELLCOST))), modConstants.CCY_MMK)
End Function

' Value x min(type MaxLTV, loan product MaxLTV): the most the item lets the bank lend.
Public Function Lendable(ByVal value As Variant, ByVal trow As Variant, ByVal productLtv As Variant) As Variant
    Dim ltv As Variant
    ltv = modMoney.ToDec(trow(CTC_MAXLTV))
    If productLtv < ltv Then ltv = productLtv
    Lendable = modMoney.RoundCcy(value * ltv, modConstants.CCY_MMK)
End Function

Public Function CollateralRow(ByVal id As String, ByVal customerId As String, ByVal typeCode As String, ByVal description As String, ByVal location As String, _
                              ByVal branchId As String, ByVal custody As String, ByVal d As Date, ByVal value As Variant) As Variant
    Dim r(1 To 12) As Variant
    r(1) = id
    r(2) = customerId
    r(3) = typeCode
    r(4) = description
    If Len(location) > 0 Then r(5) = location
    r(6) = branchId
    r(7) = custody
    r(9) = LIEN_FREE
    r(10) = d
    r(11) = value
    r(12) = d
    CollateralRow = r
End Function

Public Function AppraisalRow(ByVal id As String, ByVal collateralId As String, ByVal d As Date, ByVal value As Variant, ByVal appraiser As String, _
                             ByVal purpose As String, ByVal note As String) As Variant
    Dim r(1 To 7) As Variant
    r(1) = id
    r(2) = collateralId
    r(3) = d
    r(4) = value
    r(5) = appraiser
    r(6) = purpose
    If Len(note) > 0 Then r(7) = note
    AppraisalRow = r
End Function

Public Function OpRow(ByVal id As String, ByVal repoId As String, ByVal d As Date, ByVal opType As String, ByVal amount As Variant, ByVal note As String, ByVal journalId As String) As Variant
    Dim r(1 To 7) As Variant
    r(1) = id
    r(2) = repoId
    r(3) = d
    r(4) = opType
    r(5) = amount
    If Len(note) > 0 Then r(6) = note
    r(7) = journalId
    OpRow = r
End Function
