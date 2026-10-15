Attribute VB_Name = "repoCustomer"
Option Explicit

' Columns of tblCustomers: CustomerID, Name, CustomerType, Phone, Address, HomeBranchID, KYCLevel, TaxExempt, JoinDate, Status

' Dictionary CustomerID -> clsCustomer.
Public Function LoadCustomers() As Object
    Dim data As Variant, i As Long, c As clsCustomer, d As Object
    Set d = CreateObject("Scripting.Dictionary")
    data = modSheetUtils.ReadTable(modConstants.TBL_CUSTOMERS)
    If Not IsEmpty(data) Then
        For i = 1 To UBound(data, 1)
            If Len(CStr(data(i, 1))) > 0 Then
                Set c = New clsCustomer
                c.CustomerID = CStr(data(i, 1))
                c.Name = CStr(data(i, 2))
                c.HomeBranchID = CStr(data(i, 6))
                If IsNumeric(data(i, 7)) Then c.KYCLevel = CLng(data(i, 7))
                c.TaxExempt = (CStr(data(i, 8)) = modConstants.FLAG_YES)
                c.Status = CStr(data(i, 10))
                d.Add c.CustomerID, c
            End If
        Next i
    End If
    Set LoadCustomers = d
End Function

Public Function FindCustomer(ByVal customerId As String) As clsCustomer
    Dim d As Object
    Set d = LoadCustomers()
    If d.Exists(customerId) Then Set FindCustomer = d(customerId)
End Function
