Attribute VB_Name = "repoAsset"
Option Explicit

' tblFixedAssets columns (16): AssetID, BranchID, AssetType, Description, AcquireDate, Cost, UsefulLifeMonths, ResidualValue, DepMethod,
' AccumDepAt20260930, Status, then (added by Stage 2 setup) AccumDepreciation, LastDepPeriod, LoadedOn, DisposedOn, DisposalProceeds
Public Const FAC_ID As Long = 1
Public Const FAC_BRANCH As Long = 2
Public Const FAC_TYPE As Long = 3
Public Const FAC_DESC As Long = 4
Public Const FAC_ACQUIRED As Long = 5
Public Const FAC_COST As Long = 6
Public Const FAC_LIFE As Long = 7
Public Const FAC_RESIDUAL As Long = 8
Public Const FAC_METHOD As Long = 9
Public Const FAC_ACCUM0930 As Long = 10
Public Const FAC_STATUS As Long = 11
Public Const FAC_ACCUM As Long = 12
Public Const FAC_LASTPERIOD As Long = 13
Public Const FAC_LOADED As Long = 14
Public Const FAC_DISPOSED As Long = 15
Public Const FAC_PROCEEDS As Long = 16
Public Const FAC_COLS As Long = 16

Public Const ASSET_INUSE As String = "InUse"
Public Const ASSET_SOLD As String = "Sold"

' The next asset ID in the same style as the existing ones (prefix letters, then a zero padded number).
Public Function NextAssetId(ByVal data As Variant) As String
    Dim i As Long, s As String, p As Long, digits As String, best As Long, n As Long, prefix As String, width As Long
    prefix = "FA"
    width = 4
    If Not IsEmpty(data) Then
        For i = 1 To UBound(data, 1)
            s = CStr(data(i, FAC_ID))
            p = Len(s)
            Do While p > 0
                If Mid$(s, p, 1) Like "#" Then p = p - 1 Else Exit Do
            Loop
            digits = Mid$(s, p + 1)
            If Len(digits) > 0 And Len(digits) <= 9 Then
                n = CLng(digits)
                If n >= best Then
                    best = n
                    prefix = Left$(s, p)
                    width = Len(digits)
                End If
            End If
        Next i
    End If
    NextAssetId = prefix & Format$(best + 1, String$(width, "0"))
End Function

' True when the table has the Stage 2 columns.
Public Function IsUpgraded(ByVal data As Variant) As Boolean
    If IsEmpty(data) Then
        IsUpgraded = True
    Else
        IsUpgraded = (UBound(data, 2) >= FAC_COLS)
    End If
End Function
