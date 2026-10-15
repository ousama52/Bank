VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' A fee definition (row of tblFees).
Public FeeCode As String
Public Method As String
Public FlatAmount As Variant
Public Rate As Variant
Public MinFee As Variant
Public MaxFee As Variant
Public Ccy As String
Public IncomeGL As String
Public Taxable As Boolean
