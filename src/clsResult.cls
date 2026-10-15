VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' What every service returns. Entry macros show Message; nothing below Entry uses MsgBox.
Public Success As Boolean
Public Message As String
Public Count As Long
' Optional: the group, approval or transfer ID the action created.
Public RefID As String
