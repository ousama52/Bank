VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUnitOfWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' One business action (or one end-of-day run) builds ALL its changes here, in memory:
' account balances, journal rows, statement rows, and rows for any other table.
' Nothing touches a worksheet until Commit. An error before Commit leaves the workbook unchanged.

Public BusinessDate As Date
Public UserID As String
Public LastJournalID As String

Private m_accts As Object
Private m_chart As Object
Private m_branches As Object
Private m_glNet As Object
Private m_fxNet As Object
Private m_counters As Object
Private m_appends As Object
Private m_replaced As Object

Private Sub Class_Initialize()
    Set m_counters = CreateObject("Scripting.Dictionary")
    Set m_appends = CreateObject("Scripting.Dictionary")
    Set m_replaced = CreateObject("Scripting.Dictionary")
End Sub

Public Sub Init(ByVal d As Date, ByVal userId As String)
    BusinessDate = d
    Me.UserID = userId
    repoConfig.Reload
End Sub

' ---- Cached reference data and accounts ----

Public Function Accounts() As Object
    If m_accts Is Nothing Then Set m_accts = repoAccount.LoadAccounts()
    Set Accounts = m_accts
End Function

Public Function GetAccount(ByVal accountNo As String) As clsAccount
    If Accounts().Exists(accountNo) Then Set GetAccount = Accounts()(accountNo)
End Function

Public Sub AddAccount(ByVal a As clsAccount)
    a.IsNew = True
    a.Dirty = True
    Accounts().Add a.AccountNo, a
End Sub

Public Function Chart() As Object
    If m_chart Is Nothing Then Set m_chart = repoChart.LoadChart()
    Set Chart = m_chart
End Function

Public Function Branches() As Object
    If m_branches Is Nothing Then Set m_branches = repoBranch.LoadBranches()
    Set Branches = m_branches
End Function

' ---- GL balances (debit minus credit), kept current as journals are posted ----

Private Function GLNetDict() As Object
    If m_glNet Is Nothing Then Set m_glNet = repoJournal.LoadGLNet()
    Set GLNetDict = m_glNet
End Function

Public Function GLNet(ByVal branchId As String, ByVal glCode As String) As Variant
    Dim key As String
    key = branchId & "|" & glCode
    If GLNetDict().Exists(key) Then GLNet = GLNetDict()(key) Else GLNet = CDec(0)
End Function

Public Function GLNetAll(ByVal glCode As String) As Variant
    Dim k As Variant, total As Variant
    total = CDec(0)
    For Each k In GLNetDict().Keys
        If Mid$(CStr(k), InStr(CStr(k), "|") + 1) = glCode Then total = total + GLNetDict()(k)
    Next k
    GLNetAll = total
End Function

Public Sub AddGLNet(ByVal branchId As String, ByVal glCode As String, ByVal net As Variant)
    Dim key As String, g As Object
    Set g = GLNetDict()
    key = branchId & "|" & glCode
    If g.Exists(key) Then
        g(key) = g(key) + net
    Else
        g(key) = net
    End If
End Sub

' ---- Foreign groups ("BranchID|GLCode|Currency" -> Array(foreign signed, base signed)), kept current as journals are posted ----

Public Function FxGroups() As Object
    If m_fxNet Is Nothing Then Set m_fxNet = repoJournal.LoadFxNet()
    Set FxGroups = m_fxNet
End Function

Public Sub AddFx(ByVal branchId As String, ByVal glCode As String, ByVal ccy As String, ByVal foreign As Variant, ByVal base As Variant)
    Dim g As Object, key As String, cur As Variant
    Set g = FxGroups()
    key = branchId & "|" & glCode & "|" & ccy
    If g.Exists(key) Then
        cur = g(key)
        g(key) = Array(cur(0) + foreign, cur(1) + base)
    Else
        g(key) = Array(foreign, base)
    End If
End Sub

' Foreign balance (debit positive) of one GL in one branch and currency.
Public Function FxBalance(ByVal branchId As String, ByVal glCode As String, ByVal ccy As String) As Variant
    Dim g As Object, key As String
    Set g = FxGroups()
    key = branchId & "|" & glCode & "|" & ccy
    If g.Exists(key) Then FxBalance = g(key)(0) Else FxBalance = CDec(0)
End Function

' Foreign balance of one GL over all branches.
Public Function FxBalanceAll(ByVal glCode As String, ByVal ccy As String) As Variant
    Dim g As Object, k As Variant, parts() As String, t As Variant
    Set g = FxGroups()
    t = CDec(0)
    For Each k In g.Keys
        parts = Split(CStr(k), "|")
        If parts(1) = glCode And parts(2) = ccy Then t = t + g(k)(0)
    Next k
    FxBalanceAll = t
End Function

' ---- IDs: highest existing number + 1, counted up in memory so one action can take many ----

Private Function NextNumber(ByVal counterKey As String, ByVal tableName As String, ByVal keyColumn As String, ByVal prefix As String) As Long
    If Not m_counters.Exists(counterKey) Then
        If counterKey Like "GROUP|*" Then
            m_counters(counterKey) = modIds.MaxGroupNo(BusinessDate)
        Else
            m_counters(counterKey) = modIds.MaxNumber(tableName, keyColumn, prefix)
        End If
    End If
    m_counters(counterKey) = m_counters(counterKey) + 1
    NextNumber = m_counters(counterKey)
End Function

Public Function NewId(ByVal tableName As String, ByVal keyColumn As String, ByVal prefix As String, ByVal width As Long) As String
    NewId = modIds.FormatId(prefix, NextNumber(tableName & "|" & prefix, tableName, keyColumn, prefix), width)
End Function

Public Function NewJournalId() As String
    NewJournalId = NewId(modConstants.TBL_JOURNAL, "JournalID", "J", 8)
End Function

Public Function NewLineId() As String
    NewLineId = NewId(modConstants.TBL_JLINES, "LineID", "L", 7)
End Function

Public Function NewTxnId() As String
    NewTxnId = NewId(modConstants.TBL_TRANSACTIONS, "TxnID", "TX" & modDateUtils.DateKey(BusinessDate) & "-", 5)
End Function

Public Function NewGroupId() As String
    Dim p As String
    p = "G" & modDateUtils.DateKey(BusinessDate) & "-"
    NewGroupId = modIds.FormatId(p, NextNumber("GROUP|" & p, "", "", p), 5)
End Function

' ---- Rows for other tables ----

Public Sub AppendRow(ByVal tableName As String, ByVal rowValues As Variant)
    If Not m_appends.Exists(tableName) Then m_appends.Add tableName, New Collection
    m_appends(tableName).Add rowValues
End Sub

' The table as it will look after this unit commits (its replaced copy if there is one). Empty when it holds no data.
Public Function WorkingTable(ByVal tableName As String) As Variant
    If m_replaced.Exists(tableName) Then
        WorkingTable = m_replaced(tableName)
    Else
        WorkingTable = modSheetUtils.ReadTable(tableName)
    End If
End Function

' Registers a changed copy of existing rows (status changes, etc.). New rows still go through AppendRow.
Public Sub ReplaceTable(ByVal tableName As String, ByVal data As Variant)
    m_replaced(tableName) = data
End Sub

Public Function HasChanges() As Boolean
    HasChanges = (m_appends.Count > 0) Or (m_replaced.Count > 0) Or Not (m_accts Is Nothing)
End Function

' Writes everything: accounts, changed tables, then new rows. Call it once, after all checks passed.
Public Sub Commit()
    Dim k As Variant
    If Not m_accts Is Nothing Then repoAccount.Save m_accts
    For Each k In m_replaced.Keys
        modSheetUtils.WriteTable CStr(k), m_replaced(k)
    Next k
    For Each k In m_appends.Keys
        modSheetUtils.AppendRows CStr(k), modSheetUtils.RowsTo2D(m_appends(k))
    Next k
    m_replaced.RemoveAll
    m_appends.RemoveAll
End Sub
