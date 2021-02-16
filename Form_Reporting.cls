VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_REPORTING"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database




Private Sub btnExit_Click()
    DoCmd.Close acForm, "Reporting", acSaveNo
End Sub
Private Sub btnAuditFindingsApp_Click()

On Error GoTo err1:

    Dim tableName As String
    tableName = "tblAuditFindingsApp"
    Dim stDocName As String
    stDocName = "CHIP Enrollment Audit Findings Detail Report - App"
    If Application.CurrentProject.AllReports(stDocName).IsLoaded = True Then DoCmd.Close acReport, stDocName, acSaveNo
    
    DoCmd.SetWarnings False
    If Not IsNull(DLookup("Name", "MSysObjects", "Name='" & tableName & "'")) Then
        DoCmd.Close acTable, tableName, acSaveYes
        DoCmd.DeleteObject acTable = acDefault, tableName
        Debug.Print "Table" & tableName & "deleted..."
    End If
    
    DoCmd.OpenQuery ("qryAuditFindingsApp")
    If DCount("header_id", "tblAuditFindingsApp", "header_id IS NOT NULL") > 0 Then
        DoCmd.OpenReport stDocName, acPreview
    Else
        MsgBox "No results"
    End If


err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "Reporting; btnAuditFindingsApp_Click()")
            If MsgBox("Cannot open report.", vbCritical + vbOKOnly, "Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
    
End Sub

Private Sub btnAuditFindingsRen_Click()

On Error GoTo err1:


    Dim tableName As String
    tableName = "tblAuditFindingsRen"
    Dim stDocName As String
    stDocName = "CHIP Enrollment Audit Findings Detail Report - Ren"
    If Application.CurrentProject.AllReports(stDocName).IsLoaded = True Then DoCmd.Close acReport, stDocName, acSaveNo
    
    DoCmd.SetWarnings False
    If Not IsNull(DLookup("Name", "MSysObjects", "Name='" & tableName & "'")) Then
        DoCmd.Close acTable, tableName, acSaveYes
        DoCmd.DeleteObject acTable = acDefault, tableName
        Debug.Print "Table" & tableName & "deleted..."
    End If
    
    DoCmd.OpenQuery ("qryAuditFindingsRen")
    If DCount("header_id", "tblAuditFindingsRen", "header_id IS NOT NULL") > 0 Then
        DoCmd.OpenReport stDocName, acPreview
    Else
        MsgBox "No results"
    End If


err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "Reporting; btnAuditFindingsRen_Click()")
            If MsgBox("Cannot open report.", vbCritical + vbOKOnly, "Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
    
End Sub










Private Sub btnErrorAssessmentApp_Click()

On Error GoTo err1:
    Dim tableName As String
    tableName = "tblAssessmentFindingsApp"
    Dim stDocName As String
    stDocName = "Error Assessment Report By Specialist - App"
    If Application.CurrentProject.AllReports(stDocName).IsLoaded = True Then DoCmd.Close acReport, stDocName, acSaveNo
    
    DoCmd.SetWarnings False
    If Not IsNull(DLookup("Name", "MSysObjects", "Name='" & tableName & "'")) Then
        DoCmd.Close acTable, tableName, acSaveYes
        DoCmd.DeleteObject acTable = acDefault, tableName
        Debug.Print "Table" & tableName & "deleted..."
    End If
    
    DoCmd.OpenQuery ("qryAssessmentFindingsApp")
    If DCount("header_id", "tblAssessmentFindingsApp", "header_id IS NOT NULL") > 0 Then
        DoCmd.OpenReport stDocName, acPreview
    Else
        MsgBox "No results"
    End If
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "Reporting; btnErrorAssessmentApp_Click()")
            If MsgBox("Cannot open report.", vbCritical + vbOKOnly, "Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
    
End Sub

Private Sub btnErrorAssessmentRen_Click()

On Error GoTo err1:

    Dim tableName As String
    tableName = "tblAssessmentFindingsRen"
    Dim stDocName As String
    stDocName = "Error Assessment Report By Specialist - Ren"
    If Application.CurrentProject.AllReports(stDocName).IsLoaded = True Then DoCmd.Close acReport, stDocName, acSaveNo
    
    DoCmd.SetWarnings False
    If Not IsNull(DLookup("Name", "MSysObjects", "Name='" & tableName & "'")) Then
        DoCmd.Close acTable, tableName, acSaveYes
        DoCmd.DeleteObject acTable = acDefault, tableName
        Debug.Print "Table" & tableName & "deleted..."
    End If
    
    DoCmd.OpenQuery ("qryAssessmentFindingsRen")
    If DCount("header_id", "tblAssessmentFindingsRen", "header_id IS NOT NULL") > 0 Then
        DoCmd.OpenReport stDocName, acPreview
    Else
        MsgBox "No results"
    End If
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "Reporting; btnErrorAssessmentRen_Click()")
            If MsgBox("Cannot open report.", vbCritical + vbOKOnly, "Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
End Sub








Private Sub btnSpecialistAccuracyApp_Click()
On Error GoTo Err_btnSpecialistAccuracyApp_Click

    Dim stDocName As String

    stDocName = "CHIP Audit Specialist Accuracy Summary"
    DoCmd.OpenReport stDocName, acPreview

Exit_btnSpecialistAccuracyApp_Click:
    Exit Sub

Err_btnSpecialistAccuracyApp_Click:
    MsgBox "Error opening report"
    Resume Exit_btnSpecialistAccuracyApp_Click
End Sub

Private Sub btnSpecialistAccuracyRen_Click()
On Error GoTo Err_btnSpecialistAccuracyRen_Click
    Dim stDocName As String
    
    stDocName = "CHIP Renewal Audit Specialist Accuracy Summary"
    DoCmd.OpenReport stDocName, acPreview


Exit_btnSpecialistAccuracyRen_Click:
    Exit Sub

Err_btnSpecialistAccuracyRen_Click:
    MsgBox "Error opening report"
    Resume Exit_btnSpecialistAccuracyRen_Click
End Sub

Private Sub btnAuditorProd_Click()
On Error GoTo Err_btnAuditorProd_Click
    Dim stDocName As String
    
    stDocName = "CHIP Enrollment_Auditor Daily Production"
    DoCmd.OpenReport stDocName, acPreview


Exit_btnAuditorProd_Click:
    Exit Sub

Err_btnAuditorProd_Click:
    MsgBox "Error opening report"
    Resume Exit_btnAuditorProd_Click
End Sub


Private Sub Command193_Click()
    'COB root cause export
    
On Error GoTo err1:
    Dim exportPath As String
    Dim userName As String: userName = Environ("USERNAME")
    Dim timestamp As String: timestamp = Format(Now(), "mm_dd_yyyy_hhmmss")
    
    'use this cloud path for exporting error reports, replacing username with that of logged in user
    'C:\Users\carusomj\UPMC\QA – Gov’t Products - General\Chip Error Report\CHIP_COB_Holding\
    exportPath = "C:\Users\" & userName & "\UPMC\QA – Gov’t Products - General\Chip Error Report\CHIP_COB_Holding\"
    
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "qryRootCauseExportCOB", exportPath & "CHIPCOB_ErrorProcessing_" & timestamp
    
    If MsgBox("Export complete - would you like to view the file?", vbYesNo) = vbYes Then
        Shell "explorer.exe " & exportPath, vbNormalFocus
    End If
    
err1:
    Select Case Err.Number
        Case 0
        Case 2001
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "Reporting; Command193_Click()")
                If MsgBox("Error processing report.", vbCritical + vbOKOnly, "Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
    
End Sub

Private Sub Command195_Click()
    'CHIP root cause export
    
On Error GoTo err1:
    Dim exportPath As String
    Dim userName As String: userName = Environ("USERNAME")
    Dim timestamp As String: timestamp = Format(Now(), "mm_dd_yyyy_hhmmss")
    
    'use this cloud path for exporting error reports, replacing username with that of logged in user
    'C:\Users\carusomj\UPMC\QA – Gov’t Products - General\Chip Error Report\CHIP_Enrollment_Holding\
    exportPath = "C:\Users\" & userName & "\UPMC\QA – Gov’t Products - General\Chip Error Report\CHIP_Enrollment_Holding\"
    
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "qryRootCauseExportCHIP", exportPath & "CHIPEnrollmt_ErrorProcessing_" & timestamp
    
    If MsgBox("Export complete - would you like to view the file?", vbYesNo) = vbYes Then
        Shell "explorer.exe " & exportPath, vbNormalFocus
    End If
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "Reporting; Command195_Click()")
                If MsgBox("Error processing report.", vbCritical + vbOKOnly, "Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
    
End Sub

Private Sub btnCOBSpecialistAccuracy_Click()
'COB Audit Specialist Accuracy Summary

On Error GoTo Err_btnCOBSpecialistAccuracy_Click

    Dim stDocName As String

    stDocName = "COB Audit Specialist Accuracy Summary"
    DoCmd.OpenReport stDocName, acPreview

Exit_btnCOBSpecialistAccuracy_Click:
    Exit Sub

Err_btnCOBSpecialistAccuracy_Click:
    MsgBox "Error opening report"
    Resume Exit_btnCOBSpecialistAccuracy_Click
End Sub


Private Sub btnAuditFindingsCOB_Click()
On Error GoTo err1:

    Dim tableName As String
    tableName = "tblAuditFindingsCOB"
    Dim stDocName As String: stDocName = "COB Audit Findings Detail Report"
    If Application.CurrentProject.AllReports(stDocName).IsLoaded = True Then DoCmd.Close acReport, stDocName, acSaveNo
    
    DoCmd.SetWarnings False
    If DCount("header_id", "tblAuditFindingsCOB", "header_id is not null") > 0 Then
        DoCmd.RunSQL ("DELETE * FROM tblAuditFindingsCOB;")
    End If
    
    DoCmd.OpenQuery ("qryAuditFindingsCOB")
    If DCount("header_id", "tblAuditFindingsCOB", "header_id IS NOT NULL") > 0 Then
        DoCmd.OpenReport stDocName, acPreview
    Else
        MsgBox "No results"
    End If


err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "Reporting; btnAuditFindingsCOB_Click()")
            If MsgBox("Cannot open report.", vbCritical + vbOKOnly, "Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select


End Sub

Private Sub Command204_Click()
'CHIP enrollment rep report

On Error GoTo err1:

    DoCmd.OpenQuery "qryCHIPEnrollmentRepReport", acViewNormal, acReadOnly

err1:
    Select Case Err.Number
        Case 0
        Case Else
            MsgBox ("Report output failed")
    End Select

End Sub


