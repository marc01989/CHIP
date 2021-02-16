VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Sampling Page"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database



'REFERENCE REQUIRED FOR FILE DIALOG WINDOW - Microsoft Office 16.0 Object Library
'REFERENCE REQUIRED REGULAR EXPRESSION - Microsoft VBScript Regular Expression 5.5


Private Sub btnRunReport_Click()

    'VALIDATION
    If CheckSubmission(1) = True Then
        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            txtErrorBar.SetFocus
            Exit Sub
        End If
    End If
    

    Select Case optReports
    
        Case 1 'import universe using select week_of from combobox
            Dim weekOf1 As String: weekOf1 = txtWeekOf.Value
            
            If DCount("[week_of]", "tblUniverseCHIP", "FORMAT([week_of], 'mm/dd/yyyy') = '" & weekOf1 & "'") > 0 Then
                If MsgBox("Sample records already imported for this date." & vbCrLf & vbCrLf & "Do you want to overwrite?", vbQuestion & vbYesNo) = vbYes Then
                    DoCmd.RunSQL ("DELETE * FROM tblUniverseCHIP WHERE [week_of] = '" & weekOf1 & "';")
                    DoCmd.OpenQuery ("qappUniverse")
                    MsgBox ("Import complete")
                Else
                    Exit Sub
                End If
                
            Else
                If MsgBox("Would you like to import the CHIP sample from Excel sheet?", vbYesNo + vbQuestion, "Import?") = vbYes Then
                    DoCmd.OpenQuery ("qappUniverse")
                    MsgBox ("Import complete")
                Else
                    Exit Sub
                End If
            End If
    
        Case 2 'append to staging, excluding emails from [HPEnrollDocType]. Updates network id to full name of specialist
            DoCmd.RunSQL ("DELETE * FROM tblStaging;")
            DoCmd.OpenQuery ("qappStaging")
            DoCmd.OpenQuery ("qupdApplication")
            DoCmd.OpenQuery ("qupdRenewal")
            DoCmd.OpenQuery ("qupdNetworkId")
            
        Case 3 'find network id's that aren't on tblEmployees
            If CurrentData.AllQueries("qryUnmatchedSpecialist").IsLoaded Then DoCmd.Close acQuery, "qryUnmatchedSpecialist", acSaveNo
            DoCmd.OpenQuery "qryUnmatchedSpecialist", acViewNormal, acReadOnly
        
        Case 4 'update the network id's with name
            DoCmd.OpenQuery ("qupdNetworkId")
            
        Case 5 'daily totals
            DoCmd.OpenQuery ("qryWeeklyTotals")
            
        Case 6 'opens sample log
            'Call GenerateApp
            If CurrentData.AllQueries("qrySampleLog").IsLoaded Then DoCmd.Close acQuery, "qrySampleLog", acSaveNo
            DoCmd.OpenQuery "qrySampleLog", acViewNormal, acReadOnly
            
        Case 7
            Call GenerateSampleCHIP
            
        Case 8
            Dim fDialog As FileDialog
            Dim filePath As String, fileLocation As String
            Dim weekOf As String: weekOf = Format(txtWeekOf, "yyyymmdd")
            Dim weekStart As String: weekStart = txtWeekOf.Value
            Dim weekEnd As String
            weekStart = Format(weekStart, "mm.dd.yyyy")
            weekEnd = Format(DateAdd("d", 6, txtWeekOf.Value), "mm.dd.yyyy")
            
            Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
            fileLocation = "X:\QA - Tech Team\QA - Tech Team(DeptUsers)\Government_Products\Weekly\GPW_001_CHIP Enrollment Samples\"
            With fDialog
                .AllowMultiSelect = False
                .Title = "Select a location to export files"
                .InitialFileName = fileLocation
                .Filters.Clear
                If .Show = 0 Then
                    End
                Else
                    filePath = .SelectedItems(1)
                End If
            End With
            
            DoCmd.OutputTo acOutputReport, "SPECIALIST SAMPLE REPORT - APPLICATION", acFormatPDF, filePath & "\SPECIALIST SAMPLE REPORT - APPLICATION_" & weekOf & ".pdf"
            DoCmd.OutputTo acOutputReport, "SPECIALIST SAMPLE REPORT - RENEWAL", acFormatPDF, filePath & "\SPECIALIST SAMPLE REPORT - RENEWAL_" & weekOf & ".pdf"
            DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "qrySampleLog", filePath & "\CHIP Sample Log " & weekStart & " to " & weekEnd & ".xlsx", True
            MsgBox ("Export Complete")
            
            
        Case 9
            sql = "INSERT INTO [tblHeader] ( [week_of], [applicant], [application_source], [sample_ufi], [ssn], [enrollment_rep], [additional_id], [sample_type], [audit_type], [specialist_comments], [audit_status]  ) " & _
                " SELECT [Start Date], [Applicant Name], [Application Source], UFI, SSN, [Enrollment Rep], [Receipt Number], Type, Type, Notes, 'Not Started' " & _
                " FROM tblRandomSample;"
            DoCmd.RunSQL (sql)
            
            sql = "INSERT INTO tblReconciliation (week_of) VALUES ('" & txtWeekOf & "');"
            DoCmd.RunSQL (sql)
            
    End Select
    
End Sub


Public Sub GenerateSampleCHIP()

    'VALIDATION
    If CheckSubmission(1) = True Then
        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            txtErrorBar.SetFocus
            Exit Sub
        End If
    End If
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim qd As DAO.QueryDef
    Dim varNames As Variant
    
On Error GoTo err1:

    'GET DISTINCT LIST OF SPECIALIST NAMES WHO WORKED ITEMS FOR WEEK (txtweekof), PUT INTO varNames ARRAY
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT DISTINCT Assigned FROM tblStaging;")
    
    With rs
        If .RecordCount > 0 Then
            varNames = .GetRows(.RecordCount)
        Else
            MsgBox ("No records found for " & txtWeekOf.Value & ".")
            Exit Sub
        End If
    End With
    
    'DELETE EXISTING DATA FROM tblSample
    DoCmd.RunSQL ("DELETE * FROM tblRandomSample")
    
    'ITERATE OVER varNames, USE QUERY TO GENERATE RANDOM SAMPLE AND APPEND TO TABLE
    For i = 0 To UBound(varNames, 2)
        Set qd = db.QueryDefs("qappRandomSample")
        qd.Parameters("[emp_name]").Value = varNames(0, i)
        qd.Parameters("[week_of1]").Value = txtWeekOf.Value
        qd.Execute
        qd.Close: Set qd = Nothing
    Next i
    
    'CLEANUP
    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing
    MsgBox ("Job complete")
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "Sampling Page; GenerateSampleCHIP()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
    
End Sub


Public Sub GenerateSampleCOB()

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim universeSize As Double
    Dim sampleSize As Integer
    Const SAMPMAX As Double = 73
    Dim monthOf As String: monthOf = txtMonthOf.Value
    Dim lob(1 To 2) As Variant
    lob(1) = "Commercial"
    lob(2) = "Marketplace"
    
On Error GoTo err1:

    'DELETE EXISTING DATA FROM tblSample
    DoCmd.RunSQL ("DELETE * FROM tblRandomSampleCOB")
    
    For i = 1 To 2
        
        'GET COUNT OF TOTAL COB ITEMS BY LOB (MARKETPLACE or COMMERCIAL)
        Set db = CurrentDb
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM tblStagingCOB WHERE type = '" & lob(i) & "';")
        
        
        With rs
            If .RecordCount > 0 Then
                universeSize = .Fields(0)
                
                'CALCULATE SAMPLE SIZE
                sampleSize = -1 * Int(-1 * (SAMPMAX / (1 + (SAMPMAX / universeSize))))
                MsgBox (sampleSize)
                
            Else
                MsgBox ("No records COB found for " & txtMonthOf.Value & ".")
                Exit Sub
            End If
        End With
        
        'APPEND SAMPLE TO tblRandomSampleCOB
        DoCmd.RunSQL ("INSERT INTO tblRandomSampleCOB SELECT TOP " & sampleSize & " * FROM tblStagingCOB WHERE month_of = '" & monthOf & "' AND type = '" & lob(i) & "' ORDER BY random_number;")
    
    Next i
    
    'CLEANUP
    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing
    MsgBox ("Job complete")
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "Sampling Page; GenerateSampleCOB()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
    
End Sub

Public Function CheckSubmission(submitType As Integer) As Boolean

    Dim msgStr As String, headerStr As String, footerStr As String, resultStr As String
    Dim regex As RegExp: Set regex = New RegExp
    headerStr = "<ul>"
    footerStr = "</ul>"

    Select Case submitType
        Case 1
            If IsNull(txtWeekOf) Or txtWeekOf.Value = "" Then
                msgStr = msgStr & "<b>Week Of</b> cannot be blank."
            End If
            With regex
                .Multiline = False
                .Global = True
                .IgnoreCase = True
                .Pattern = "^(0[1-9]|1[012])[\/](0[1-9]|[12][0-9]|3[01])[\/](19|20)\d\d$"
            End With
            If Not IsNull(txtWeekOf) Then
                If regex.test(txtWeekOf.Value) = False Then
                    msgStr = msgStr & "<b>Week Of</b> must be in 'mm/dd/yyyy' format."
                End If
            End If
        Case 2
            If IsNull(txtMonthOf) Or txtMonthOf.Value = "" Then
                msgStr = msgStr & "<b>Month Of</b> cannot be blank."
            End If
    End Select

    If msgStr = "" Then 'no errors
        txtErrorBox.Value = Null
        txtErrorBar.Value = Null
        txtErrorBar.BackColor = RGB(242, 242, 242)
        CheckSubmission = False
    Else 'errors, show msg to user
        txtErrorBox.Value = headerStr & msgStr & footerStr
        txtErrorBar.Value = "Submission Errors"
        txtErrorBar.BackColor = RGB(255, 186, 0)
        CheckSubmission = True
    End If


End Function


Private Sub btnRunCOBReports_Click()
    
    'VALIDATION
    If CheckSubmission(2) = True Then
        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            txtErrorBar.SetFocus
            Exit Sub
        End If
    End If
    

    Select Case optCOBReports
    
        Case 1
            Dim monthOf As String: monthOf = txtMonthOf.Value
            
            If DCount("[month_of]", "tblUniverseCOB", "FORMAT([month_of], 'mm/dd/yyyy') = '" & monthOf & "'") > 0 Then
                If MsgBox("Sample records already imported for this date." & vbCrLf & vbCrLf & "Do you want to overwrite?", vbQuestion & vbYesNo) = vbYes Then
                    DoCmd.RunSQL ("DELETE * FROM tblUniverseCOB WHERE [month_of] = '" & monthOf & "';")
                    DoCmd.OpenQuery ("qappUniverseCOB")
                    MsgBox ("Import complete")
                Else
                    Exit Sub
                End If
                
            Else
                If MsgBox("Would you like to import the COB CHIP sample from Excel sheet?", vbYesNo + vbQuestion, "Import?") = vbYes Then
                    DoCmd.OpenQuery ("qappUniverseCOB")
                    MsgBox ("Import complete")
                Else
                    Exit Sub
                End If
            End If
            
        Case 2
            DoCmd.RunSQL ("DELETE * FROM tblStagingCOB;")
            DoCmd.OpenQuery ("qappStagingCOB")
            DoCmd.OpenQuery ("qupdNetworkIdCOB")
            
        Case 3
            If CurrentData.AllQueries("qryUnmatchedSpecialistCOB").IsLoaded Then DoCmd.Close acQuery, "qryUnmatchedSpecialistCOB", acSaveNo
            DoCmd.OpenQuery "qryUnmatchedSpecialistCOB", acViewNormal, acReadOnly
            
        Case 4
            DoCmd.OpenQuery ("qupdNetworkIdCOB")
            
        Case 5
        
        
        Case 6
            Call GenerateSampleCOB
            
        Case 7
            
            
        Case 8
            sql = "INSERT INTO [tblHeaderCOB] ( [B - Member ID], [A - Employee Number], [D - CHIP Last Name], [C - CHIP First Name], [E - CHIP Start], [F - CHIP Term], [CHIP MMI], [CHIP MMI Added Date], [H - NC Member ID], [J - NC Last Name], [I - NC First Name], [NC Subscriber Name], [NC Subscriber Gender], [NC MMI], [NC MMI Added Date], [K - NC Start], [L - NC Term], [O - NC Group], [NC Sbgrp], [DAYS OF OVERLAP], [Comments], [AssignedToUser], [Type], [audit_status], [month_of] ) " & _
                " SELECT " & _
                " [B - Member ID], [A - Employee Number], [D - CHIP Last Name], [C - CHIP First Name], [E - CHIP Start], [F - CHIP Term], [CHIP MMI], [CHIP MMI Added Date], [H - NC Member ID], [J - NC Last Name], [I - NC First Name], [NC Subscriber Name], [NC Subscriber Gender], [NC MMI], [NC MMI Added Date], [K - NC Start], [L - NC Term], [O - NC Group], [NC Sbgrp], [DAYS OF OVERLAP], [Comments], [AssignedToUser], [Type], 'Not Started', [month_of] " & _
                " FROM tblRandomSampleCOB;"
            DoCmd.RunSQL (sql)
            
    End Select
    
End Sub






















