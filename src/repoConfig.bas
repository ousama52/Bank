Attribute VB_Name = "repoConfig"
Option Explicit

' tblConfig as key/value. Values are text; numbers and dates are parsed without locale.
Private m_cache As Object

Public Sub Reload()
    Set m_cache = Nothing
End Sub

Private Function Cache() As Object
    Dim data As Variant, i As Long
    If m_cache Is Nothing Then
        Set m_cache = CreateObject("Scripting.Dictionary")
        m_cache.CompareMode = 1
        data = modSheetUtils.ReadTable(modConstants.TBL_CONFIG)
        If Not IsEmpty(data) Then
            For i = 1 To UBound(data, 1)
                If Len(CStr(data(i, 1))) > 0 Then m_cache(CStr(data(i, 1))) = data(i, 2)
            Next i
        End If
    End If
    Set Cache = m_cache
End Function

Public Function Has(ByVal key As String) As Boolean
    Has = Cache().Exists(key)
End Function

Public Function GetText(ByVal key As String, Optional ByVal defaultValue As String = "") As String
    If Cache().Exists(key) Then
        GetText = Trim$(CStr(Cache()(key)))
    Else
        GetText = defaultValue
    End If
End Function

Public Function GetDec(ByVal key As String) As Variant
    Dim d As Variant
    If Not Cache().Exists(key) Then Err.Raise vbObjectError + 1200, "repoConfig.GetDec", "Config key " & key & " is missing. Run RunSetup."
    If Not modMoney.TryToDec(Cache()(key), d) Then Err.Raise vbObjectError + 1201, "repoConfig.GetDec", "Config key " & key & " is not a number."
    GetDec = d
End Function

Public Function GetLong(ByVal key As String) As Long
    GetLong = CLng(GetDec(key))
End Function

Public Function GetBool(ByVal key As String) As Boolean
    GetBool = (StrComp(GetText(key), modConstants.FLAG_YES, vbTextCompare) = 0)
End Function

Public Function BusinessDate() As Date
    Dim d As Date
    If Not Cache().Exists(modConstants.CFG_BUSDATE) Then Err.Raise vbObjectError + 1200, "repoConfig.BusinessDate", "Config key BusinessDate is missing. Run RunSetup."
    If Not modDateUtils.TryToDate(Cache()(modConstants.CFG_BUSDATE), d) Then Err.Raise vbObjectError + 1202, "repoConfig.BusinessDate", "Config BusinessDate is not a date."
    BusinessDate = d
End Function

' Changes or adds one key. The caller writes the audit row.
Public Sub SetValue(ByVal key As String, ByVal text As String)
    Dim data As Variant, i As Long, newRow(1 To 1, 1 To 2) As Variant
    data = modSheetUtils.ReadTable(modConstants.TBL_CONFIG)
    If Not IsEmpty(data) Then
        For i = 1 To UBound(data, 1)
            If StrComp(CStr(data(i, 1)), key, vbTextCompare) = 0 Then
                data(i, 2) = text
                modSheetUtils.WriteTable modConstants.TBL_CONFIG, data
                Reload
                Exit Sub
            End If
        Next i
    End If
    newRow(1, 1) = key
    newRow(1, 2) = text
    modSheetUtils.AppendRows modConstants.TBL_CONFIG, newRow
    Reload
End Sub

Public Sub SetBusinessDate(ByVal d As Date)
    SetValue modConstants.CFG_BUSDATE, modDateUtils.ToIso(d)
End Sub
