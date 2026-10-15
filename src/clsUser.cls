VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' A bank user (row of tblUsers).
Public UserID As String
Public Name As String
Public Role As String
Public BranchID As String
Public IsActive As Boolean
