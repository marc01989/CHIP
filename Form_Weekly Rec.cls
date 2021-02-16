VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Weekly Rec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
    Dim varApps As Variant

Private Sub btnClear_Click()
    Call ClearForm
End Sub

Private Sub Form_Load()
    Call ClearForm
    Call LoadPendingRecs
End Sub
Private Sub LoadPendingRecs()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM tblReconciliation ORDER BY recon_id;", dbReadOnly)
    With rs
        If .RecordCount > 0 Then
            varApps = .GetRows(.RecordCount)
        End If
    End With
    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing
    
    lstPendingRecs.RowSource = ""
    lstPendingRecs.AddItem "Week Of," & "Reconciliation Complete?,"
    For j = 0 To UBound(varApps, 2)
        lstPendingRecs.AddItem varApps(1, j) & "," & varApps(6, j)
    Next j
End Sub

Private Sub btnSearch_Click()

    'VALIDATION
    If DCount("week_of", "tblHeader", "week_of = '" & cboWeeks.Value & "'") = 0 Then
        MsgBox ("No Data")
        Call ClearForm
        Exit Sub
    End If


    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim varSample As Variant
    Dim population As Integer, targetSamp As Integer, actualSamp As Integer, unableToRev As Integer, itemsAudited As Integer
    
    Set db = CurrentDb
    Call ClearForm
    
    'SAMPLE LOG
    Set rs = db.OpenRecordset("SELECT tblEmployees.emp_name, Count(tblUniverseCHIP.assigned) AS [Total Items], " & _
        " IIf(Abs(Int(-([Total Items]*0.05)))>20,20,Abs(Int(-([Total Items]*0.05)))) AS [Sample Size] " & _
        " FROM tblEmployees RIGHT JOIN tblUniverseCHIP ON tblEmployees.network_id = tblUniverseCHIP.Assigned " & _
        " WHERE (tblUniverseCHIP.week_of = '" & cboWeeks.Value & "' AND tblUniverseCHIP.HPEnrollDocType<>'Email') " & _
        " GROUP BY tblEmployees.emp_name;")
    
    With rs
        If .RecordCount > 0 Then
            varSample = .GetRows(.RecordCount)
        End If
        .Close
    End With
    
    lstSampleLog.RowSource = ""
    lstSampleLog.AddItem "Enroll Specialist," & "Total Items," & "Sample Size,"
    For j = 0 To UBound(varSample, 2)
        lstSampleLog.AddItem varSample(0, j) & "," & varSample(1, j) & "," & varSample(2, j)
        population = population + varSample(1, j)
        targetSamp = targetSamp + varSample(2, j)
    Next j
    lstSampleLog.AddItem "TOTAL," & population & "," & targetSamp
    
    'POPULATION
    txtPopulation.Value = population
    
    'TARGET SAMPLE SIZE
    txtTargetSampSize.Value = targetSamp
    
    'ACTUAL SAMPLE SIZE
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM tblHeader WHERE week_of = '" & cboWeeks.Value & "';")
    With rs
        If .RecordCount > 0 Then
            actualSamp = .Fields(0)
        End If
        .Close
    End With
    
    'UNABLE TO REVIEW
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM tblHeader WHERE week_of = '" & cboWeeks.Value & "' AND unable_to_review = 1;")
    With rs
        If .RecordCount > 0 Then
            unableToRev = .Fields(0)
            txtUnableToReview.Value = unableToRev
        End If
        .Close
    End With
    
    'CALCULATE ACTUAL SAMPLE SIZE (COUNT OF RECORDS ON tblHeader MINUS UNABLE TO REVIEW RECORDS
    actualSamp = actualSamp - unableToRev
    txtActualSampleSize.Value = actualSamp
    
    'RESAMPLE NEEDED
    If targetSamp <> actualSamp Then
        txtResample.Value = "Yes"
        txtResample.BackColor = RGB(255, 99, 71)
    Else
        txtResample.Value = "No"
        txtResample.BackColor = RGB(144, 238, 144)
    End If
    
    'ITEMS AUDITED
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM tblHeader WHERE (week_of = '" & cboWeeks.Value & "' AND audit_status = 'Complete' AND unable_to_review <> 1);")
    With rs
        If .RecordCount > 0 Then
            itemsAudited = .Fields(0)
        End If
        .Close
    End With
    
    txtItemsAudited.Value = itemsAudited
    If targetSamp <> itemsAudited Then
        txtItemsAudited.BackColor = RGB(255, 99, 71)
    Else
        txtItemsAudited.BackColor = RGB(144, 238, 144)
    End If
    
    'RECONCILIATION INFO
    Set rs = db.OpenRecordset("SELECT * FROM tblReconciliation WHERE week_of = '" & cboWeeks.Value & "';", dbReadOnly)
    With rs
        If .RecordCount > 0 Then
            txtRecId.Value = .Fields(0)
            chkManagerSignoff.Value = .Fields(2)
            cboManagerName.Value = .Fields(3)
            txtSignoffDate.Value = .Fields(4)
            chkPrinted.Value = .Fields(5)
        End If
        .Close
    End With
    
    Set rs = Nothing
    db.Close: Set db = Nothing
    
End Sub

Private Sub btnSave_Click()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    If MsgBox("Do you want to mark this reconciliation for week of " & cboWeeks.Value & " as complete?", vbQuestion + vbYesNo, "Complete Reconciliation?") = vbYes Then
        Set db = CurrentDb
        Set rs = db.OpenRecordset("tblReconciliation", dbOpenDynaset)
        With rs
            .FindFirst "week_of = '" & cboWeeks.Value & "'"
            If Not .NoMatch Then
                .Edit
                    ![manager_signoff] = chkManagerSignoff
                    ![manager_name] = cboManagerName
                    ![signoff_date] = txtSignoffDate
                    ![summary_printed] = chkPrinted
                    ![rec_complete] = True
                .Update
            End If
        End With
        rs.Close: Set rs = Nothing
        db.Close: Set db = Nothing
        MsgBox ("Reconciliation complete.")
    Else
        Exit Sub
    End If
    Call LoadPendingRecs
End Sub
Private Sub btnPrint_Click()

    'VALIDATION
    If ValidateForm(1) = True Then
         If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            txtErrorBar.SetFocus
             Exit Sub
         End If
    End If

    

'On Error GoTo err1:

    Dim fDialog As FileDialog
    Dim filePath As String, fileLocation As String
    Dim weekOf As String: weekOf = Format(cboWeeks.Value, "yyyymmdd")
    Dim weekStart As String: weekStart = cboWeeks.Value
    Dim weekEnd As String
'    weekStart = Format(weekStart, "mm.dd.yyyy")
'    weekEnd = Format(DateAdd("d", 6, cboWeeks.Value), "mm.dd.yyyy")
    weekStart = Format(weekStart, "mm/dd/yyyy")
    weekEnd = Format(DateAdd("d", 6, cboWeeks.Value), "mm/dd/yyyy")
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

    DoCmd.OutputTo acOutputReport, "Weekly Reconciliaton", acFormatPDF, filePath & "\CHIP Weekly Reconciliation_" & weekOf & ".pdf"
    
    MsgBox ("Export Complete")
    
'err1:
'    Select Case Err.Number
'        Case 0
'        Case Else
'            Call LogError(Err.Number & " " & Err.Description, "Weekly Rec; btnPrint_Click()")
'            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
'            Exit Sub
'    End Select
    
End Sub
Private Sub ClearForm()

    txtResample.BackColor = RGB(251, 252, 252)
    txtItemsAudited.BackColor = RGB(251, 252, 252)
    txtErrorBar.BackColor = RGB(255, 255, 255)
    lstSampleLog.RowSource = ""
    
    With Me
        Dim ctrl As Control
        For Each ctrl In Controls
            If TypeOf ctrl Is TextBox Then
                ctrl.Value = Null
                ctrl.Enabled = True
            ElseIf TypeOf ctrl Is ComboBox And ctrl.Name <> "cboWeeks" Then
                ctrl.Value = Null
                ctrl.Enabled = True
            ElseIf TypeOf ctrl Is CheckBox And ctrl.Name <> "cboWeeks" Then
                ctrl.Value = False
                ctrl.Enabled = True
            End If
        Next
    End With
End Sub

Public Function ValidateForm(submitType As Integer) As Boolean

    Dim msgStr As String, headerStr As String, footerStr As String, ctlName As String
    headerStr = "<ul>": footerStr = "</ul>"
    Dim varFocus As Variant

    Select Case submitType
        Case 1 'PRINT SUMMARY
            If txtRecId.Value = "" Or IsNull(txtRecId) Then
                msgStr = msgStr & "<li>You must select a week before printing the reconciliation summary.</li>"
                ctlName = ctlName & "cboWeeks,"
            End If
            If txtResample.Value = "Yes" Then
                msgStr = msgStr & "<li>Resamples must be pulled before printing form.</li>"
                ctlName = ctlName & "cboWeeks,"
            End If
            If txtItemsAudited.Value <> txtTargetSampSize.Value Then
                msgStr = msgStr & "<li>All items must be audited before printing form.</li>"
                ctlName = ctlName & "cboWeeks,"
            End If
    End Select
    
    If msgStr = "" Then 'no errors
        txtErrorBox.Value = Null
        txtErrorBar.Value = Null
        txtErrorBar.BackColor = RGB(255, 255, 255)
        ValidateForm = False
    Else 'errors, show msg to user
        txtErrorBox.Value = headerStr & msgStr & footerStr
        txtErrorBar.Value = "Submission Errors"
        txtErrorBar.BackColor = RGB(255, 186, 0)
        varFocus = Split(ctlName, ",")
        'Me.Controls(varFocus(0)).SetFocus
        ValidateForm = True
    End If

End Function
