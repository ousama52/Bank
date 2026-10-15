VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOtherBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' A simulated other bank (row of tblOtherBanks).
Public BankCode As String
Public Name As String
Public SettleDays As Long
Public FailRate As Variant
