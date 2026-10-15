VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Plain data holder for one row of tblAccounts. Optional dates and rates stay Variant (Empty = blank).
Public AccountNo As String
Public CustomerID As String
Public BranchID As String
Public ProductCode As String
Public Ccy As String
Public Status As String
Public OpenDate As Variant
Public TermMonths As Long
Public MaturityDate As Variant
Public LockedRate As Variant
Public PayoutAccountNo As String
Public Balance As Variant
Public AccruedInterest As Variant
Public LastAccrualDate As Variant
Public LastActivityDate As Variant
Public ClosedDate As Variant
' Filled from the product when loaded (not stored in the table).
Public ProductType As String
Public DepositGL As String
Public MinBalance As Variant
Public IsNew As Boolean
Public Dirty As Boolean
