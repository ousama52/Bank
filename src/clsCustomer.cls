VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' A customer (row of tblCustomers).
Public CustomerID As String
Public Name As String
Public HomeBranchID As String
Public KYCLevel As Long
Public TaxExempt As Boolean
Public Status As String
