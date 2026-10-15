VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsApproval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' One row of tblApprovals.
Public ApprovalID As String
Public RequestType As String
Public AccountNo As String
Public CounterAccountNo As String
Public Amount As Variant
Public BranchID As String
Public Note As String
Public RequestedBy As String
Public RequestedAt As Date
Public Status As String
Public ApprovedBy As String
Public DecidedAt As Variant
Public DecisionNote As String
Public TxnGroupID As String
