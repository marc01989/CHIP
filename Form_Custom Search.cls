VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Custom Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim formType As String

Private Sub Form_Load()

    If Not IsNull(Forms![Custom Search].OpenArgs) Then
        formType = Forms![Custom Search].OpenArgs
    Else
        Call LogError(0 & " " & "Form Type not passed to custom search form", "Custom Search; Form_Load()")
        DoCmd.Close acForm, "Custom Search", acSaveNo
        Exit Sub
    End If


End Sub

Private Sub btnSearch_Click()

    Dim sqlStr As String
    Dim searchType As String
    Dim searchVal As String
    searchVal = Clean(txtSearchValue.Value)
    
    Select Case optSearchType
        Case 1 'UCI
            searchType = "[uci]"
            
        Case 2 'Audit UFI
            searchType = "[audit_ufi]"
            
        Case 3 'Dependent Name
            searchType = "[dependent_name]"
    End Select

    Select Case formType
        Case "Apps"
            sqlStr = "SELECT * FROM tblHeader WHERE " & searchType & " = '" & searchVal & "' AND audit_type = 'Application' ORDER BY header_id;"
        Case "Rens"
            sqlStr = "SELECT * FROM tblHeader WHERE " & searchType & " = '" & searchVal & "' AND audit_type = 'Renewal' ORDER BY header_id;"
    End Select
    
    Select Case formType
        Case "Apps"
            Call Forms("Apps").SearchHeaderTbl(sqlStr)
            DoCmd.Close acForm, "Custom Search", acSaveNo
        Case "Rens"
            Call Forms("Rens").SearchHeaderTbl(sqlStr)
            DoCmd.Close acForm, "Custom Search", acSaveNo
    End Select

End Sub


