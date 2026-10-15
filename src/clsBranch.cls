VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBranch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' A branch (row of tblBranches).
Public BranchID As String
Public Name As String
Public BranchType As String
Public Status As String
Public VaultMin As Variant
Public VaultMax As Variant
