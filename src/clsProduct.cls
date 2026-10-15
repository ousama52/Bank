VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' A deposit product (row of tblProducts).
Public ProductCode As String
Public Name As String
Public ProductType As String
Public Ccy As String
Public MinBalance As Variant
Public MinOpening As Variant
Public MaintenanceFeeCode As String
Public IsActive As Boolean
