VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTxnRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' One row of tblTransactions (the customer statement).
Public TxnID As String
Public GroupID As String
Public TxnDate As Date
Public AccountNo As String
Public TxnType As String
Public Direction As String
Public Amount As Variant
Public BalanceAfter As Variant
Public CounterAccountNo As String
Public BranchID As String
Public Description As String
Public Reference As String
Public JournalID As String
Public Status As String
Public CreatedBy As String
Public ApprovedBy As String
