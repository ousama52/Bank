VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsJournalLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' One ledger line. Debit and Credit are Decimal (Variant).
Public BranchID As String
Public GLCode As String
Public AccountNo As String
Public Ccy As String
Public Debit As Variant
Public Credit As Variant
' Foreign legs (Stage 6): ForeignAmount in the line currency, FxRate = MMK per 1 unit. Empty for MMK lines.
Public FxRate As Variant
Public ForeignAmount As Variant
