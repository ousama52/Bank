Attribute VB_Name = "repoProduct"
Option Explicit

' Columns of tblProducts: ProductCode, Name, Type, Currency, MinBalance, MinOpening, MaintenanceFeeCode, IsActive

' Dictionary ProductCode -> clsProduct.
Public Function LoadProducts() As Object
    Dim data As Variant, i As Long, p As clsProduct, d As Object
    Set d = CreateObject("Scripting.Dictionary")
    data = modSheetUtils.ReadTable(modConstants.TBL_PRODUCTS)
    If Not IsEmpty(data) Then
        For i = 1 To UBound(data, 1)
            If Len(CStr(data(i, 1))) > 0 Then
                Set p = New clsProduct
                p.ProductCode = CStr(data(i, 1))
                p.Name = CStr(data(i, 2))
                p.ProductType = CStr(data(i, 3))
                p.Ccy = CStr(data(i, 4))
                p.MinBalance = modMoney.ToDec(data(i, 5))
                p.MinOpening = modMoney.ToDec(data(i, 6))
                p.MaintenanceFeeCode = CStr(data(i, 7))
                p.IsActive = (CStr(data(i, 8)) = modConstants.FLAG_YES)
                d.Add p.ProductCode, p
            End If
        Next i
    End If
    Set LoadProducts = d
End Function

Public Function FindProduct(ByVal productCode As String) As clsProduct
    Dim d As Object
    Set d = LoadProducts()
    If d.Exists(productCode) Then Set FindProduct = d(productCode)
End Function

' Deposit GL by product type (posting rules table).
Public Function DepositGL(ByVal productType As String) As String
    Select Case productType
        Case modConstants.PT_SAVINGS: DepositGL = modConstants.GL_SAVINGS
        Case modConstants.PT_CURRENT: DepositGL = modConstants.GL_CURRENT
        Case modConstants.PT_FIXED: DepositGL = modConstants.GL_FIXED
        Case modConstants.PT_WALLET: DepositGL = modConstants.GL_WALLETS
    End Select
End Function
