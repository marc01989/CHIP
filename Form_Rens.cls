VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Rens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim varApps As Variant
Dim recordMax As Long, recordCounter As Long
Dim userId As Integer
Dim bolChanged As Boolean

Private Sub btnCustomSearch_Click()
    DoCmd.OpenForm "Custom Search", acNormal, , , , acFormPropertySettings, "Rens"
End Sub

Private Sub Form_Load()
    Dim obj As AccessObject, dbs As Object
    Dim args As String
    cboWeeks.SetFocus
    lstWorkList.RowSource = ""
    txtAuditStatus.BackColor = RGB(251, 252, 252)
    
    'checks if value is passed when opening form, runs main query if so
    If Not IsNull(Forms![Rens].OpenArgs) Then
        args = Forms![Rens].OpenArgs
            If args = 0 Or IsNull(args) Then
                MsgBox ("User not found - please login again")
                Set dbs = Application.CurrentProject
                For Each obj In dbs.AllForms
                    If obj.IsLoaded = True Then
                      DoCmd.Close acForm, obj.Name, acSaveNo
                    End If
                Next obj
                DoCmd.OpenForm "Login", acNormal, , , , acWindowNormal
                Call LogError(0 & " " & "User Id not found or not passed to home screen", "Home; Form_Load()")
                Exit Sub
            Else
                userId = CInt(args)
                'MsgBox (userId)
            End If
    Else
        MsgBox ("User not found - please login again")
        Set dbs = Application.CurrentProject
        For Each obj In dbs.AllForms
            If obj.IsLoaded = True Then
              DoCmd.Close acForm, obj.Name, acSaveNo
            End If
        Next obj
        DoCmd.OpenForm "Login", acNormal, , , , acWindowNormal
        Call LogError(0 & " " & "User Id not found or not passed to home screen", "Home; Form_Load()")
        Exit Sub
    End If
    
    Call SetListeners
End Sub
Private Sub btnClear_Click()
    Call ClearForm
End Sub
Private Sub btnClose_Click()
    If CheckUnsavedChanges = False Then Exit Sub
    DoCmd.Close acForm, "Rens", acSaveNo
End Sub
Private Sub btnSearch_Click()
    'VALIDATION
    If ValidateForm(1) = True Then
        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            Exit Sub
        End If
    End If
    
    Dim weekOf As String: weekOf = cboWeeks.Value
    Dim sqlStr As String: sqlStr = "SELECT * FROM tblHeader WHERE FORMAT(week_of, 'MM/DD/YYYY') = '" & weekOf & "' AND audit_type = 'Renewal' ORDER BY header_id;"
    
    Call SearchHeaderTbl(sqlStr)
End Sub


Public Sub SearchHeaderTbl(sql As String)

    'VALIDATION - CHECK FOR UNSAVED CHANGES
    If CheckUnsavedChanges = False Then Exit Sub
    
    'START CONN
    On Error GoTo err1:
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset(sql, dbReadOnly)
    With rs
        If .RecordCount > 0 Then
            .MoveFirst
            .MoveLast
            .MoveFirst
            recordMax = .RecordCount - 1
            recordCounter = 0
            varApps = .GetRows(.RecordCount)
            
            Call PopulateBoxes(recordCounter)
            Call PopulateList
        Else
            MsgBox ("No records found")
            ClearForm
        End If
    End With
    
    bolChanged = False
    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "Rens; btnSearch_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
    
End Sub
Public Sub PopulateBoxes(i As Long)


On Error GoTo err1:
    If Not IsNull(varApps(3, i)) Then: auditorName = DLookup("auditor", "tblAuditors", "auditor_id = " & varApps(3, i))

    txtId.Value = varApps(0, i)
    'txtWeekOf = varApps(1, i)
    txtRep.Value = varApps(2, i)
    txtAuditor.Value = auditorName
    txtApplicantName.Value = varApps(4, i)
    txtSampleUFI.Value = varApps(5, i)
    txtAdditionalId.Value = varApps(6, i)
    txtMemberNumber.Value = varApps(7, i)
    txtSSN.Value = varApps(8, i)
    txtDependentName.Value = varApps(9, i)
    txtAuditUFI.Value = varApps(10, i)
    txtUCI.Value = varApps(11, i)
    txtAppSource.Value = varApps(12, i)
    'sample_type varApps(13, i)
    'txtAuditType.Value = varApps(14, i)
    cboUnableToReview.Value = varApps(15, i)
    txtComments.Value = varApps(16, i)
    txtSpecialistComments.Value = varApps(17, i)
    txtAuditDate.Value = varApps(18, i)
    txtDateReceived.Value = varApps(19, i)
    txtDateEntered.Value = varApps(20, i)
    txtDateProcessed.Value = varApps(21, i)
    txtAuditStatus = varApps(22, i)
    If varApps(22, i) = "Complete" Then
        txtAuditStatus.BackColor = RGB(144, 238, 144)
    Else
        txtAuditStatus.BackColor = RGB(255, 99, 71)
    End If
    cboPendingResponse.Value = varApps(23, i)
    
    txtRecordCounter.Value = i + 1
    txtRecordMax.Value = recordMax + 1

    Dim recordId As Long: recordId = varApps(0, i)
    Call GetElements(recordId)

err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "Rens; PopulateBoxes()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select

End Sub
Private Sub GetElements(recordId As Long)

    'START CONN
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim ctlName As String
    
    ClearElements
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM [tblElementJunc] WHERE [header_id] = " & recordId & " ORDER BY element_id ASC;", dbReadOnly)
    With rs
        If .RecordCount > 0 Then
            .MoveFirst
            For i = 1 To .RecordCount
                ctlName = .Fields(1)
                Me.Controls(ctlName).Value = .Fields(2)
                .MoveNext
            Next i
        Else
            'exit sub?
        End If
    End With
    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing
    
End Sub
Private Sub btnSubmit_Click()

   'VALIDATION
    If ValidateForm(3) = True Then
        Call LogUserOff("btnSubmit_Click()")
        Exit Sub
    End If
    If ValidateForm(2) = True Then
         If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            txtErrorBar.SetFocus
             Exit Sub
         End If
    End If
       
On Error GoTo err1:
    'BEGIN WRITING DATA
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim recordId As Long: recordId = txtId.Value
    Dim subDate As String: subDate = Now()
    Dim ctlArray As Variant
    Dim ctlName As String
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT element_id FROM tblElements WHERE is_active = true AND type = 'renewal';", dbReadOnly)
    With rs
        If .RecordCount > 0 Then
            .MoveFirst
            .MoveLast
            .MoveFirst
            ctlArray = .GetRows(.RecordCount)
        Else
            MsgBox ("Cannot submit record")
            Exit Sub
        End If
    End With
    rs.Close
    
    Set rs = db.OpenRecordset("tblHeader", dbOpenDynaset)
    With rs
        .FindFirst ("[header_id] = " & recordId)
        If Not .NoMatch Then
            .Edit
                'UPDATE TABLE INFORMATION
                '![record#] = ""
                '![week_of] = ""
                '![enrollment_rep] = txtRep
                ![auditor_id] = userId
                '![applicant] = txtApplicantName
                '![sample_ufi] = txtSampleUFI
                '![additional_id] = txtAdditionalid.value
                ![member_number] = txtMemberNumber
                ![SSN] = txtSSN
                ![dependent_name] = Clean(txtDependentName)
                ![audit_ufi] = txtAuditUFI
                ![uci] = txtUCI.Value
                '![application_source] = txtAppSource
                '[sample_type] = ""
                '![audit_type] = txtAuditType
                ![unable_to_review] = cboUnableToReview
                ![Comments] = txtComments
                '![specialist_comments] = txtSpecialistNotes
                ![date_audited] = subDate
                ![date_received] = txtDateReceived
                ![date_entered] = txtDateEntered
                ![date_processed] = txtDateProcessed
                ![audit_status] = "Complete"
                txtAuditStatus.Value = "Complete"
                txtAuditStatus.BackColor = RGB(144, 238, 144)
                ![pending_response] = cboPendingResponse
            .Update
            .Close
            
            'UPDATE LOCAL ARRAY
            'varApps(0, recordCounter) = ""
            'varApps(1, recordCounter) = ""
            'varApps(2, recordCounter) = ""
            varApps(3, recordCounter) = userId
            'varApps(4, recordCounter) = ""
            'varApps(5, recordCounter) = txtUFI
            'varapps(6, recordcounter) = txtadditionalid
            varApps(7, recordCounter) = txtMemberNumber
            varApps(8, recordCounter) = txtSSN
            varApps(9, recordCounter) = Clean(txtDependentName)
            varApps(10, recordCounter) = txtAuditUFI
            varApps(11, recordCounter) = txtUCI
            'varApps(12, recordCounter) = txtappsource
            'varApps(13, recordCounter) = txtsampletype
            'varApps(14, recordCounter) = txtAuditType
            varApps(15, recordCounter) = cboUnableToReview
            varApps(16, recordCounter) = txtComments
            'varApps(17, recordCounter) = txtspecialistcomments
            varApps(18, recordCounter) = subDate
            varApps(19, recordCounter) = txtDateReceived
            varApps(20, recordCounter) = txtDateEntered
            varApps(21, recordCounter) = txtDateProcessed
            varApps(22, recordCounter) = "Complete"
            varApps(23, recordCounter) = cboPendingResponse
            txtAuditDate.Value = subDate
            txtAuditor.Value = DLookup("auditor", "tblAuditors", "auditor_id = " & userId)
        Else
            MsgBox ("Cannot submit record - please try again")
        End If
    End With
    
   
    Set rs = db.OpenRecordset("tblElementJunc", dbOpenDynaset)
    rs.FindFirst ("[header_id] = " & recordId)
    'delete existing elements
        If Not rs.NoMatch Then
            With rs
                Do While Not .NoMatch
                    .Edit
                    .Delete
                    .FindNext ("[header_id] = " & recordId)
                Loop
            End With
        End If
        
    'add new elements
    With rs
        For i = 0 To UBound(ctlArray, 2)
            ctlName = ctlArray(0, i) 'element id IS ALSO NAME OF CONTROLS ON FORM, EACH FOR AN ELEMENT IN THE LIST

            .AddNew
                ![header_id] = recordId
                ![element_id] = ctlArray(0, i)
                ![result] = Me.Controls(ctlName).Value & vbNullString
            .Update
        Next i
    End With
            
    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing
    Call PopulateList
    bolChanged = False
    MsgBox ("Item submitted")
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "Rens; btnSubmit_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
End Sub
Private Sub btnBack_Click()
    'previous record
    If CheckUnsavedChanges = False Then Exit Sub
    If recordMax = 0 Then Exit Sub
    If recordCounter = 0 Then
        MsgBox ("No more records")
        Exit Sub
    Else
        recordCounter = recordCounter - 1
        Call PopulateBoxes(recordCounter)
    End If
End Sub
Private Sub btnForward_Click()
    'next record
    If CheckUnsavedChanges = False Then Exit Sub
    If recordMax = 0 Then Exit Sub
    If recordCounter = recordMax Then
        MsgBox ("No more records")
        Exit Sub
    Else
        recordCounter = recordCounter + 1
        Call PopulateBoxes(recordCounter)
    End If
End Sub
Private Sub PopulateList()
    lstWorkList.RowSource = ""
    lstWorkList.AddItem "Record#," & "Applicant," & "Enroll Rep," & "Audit Status," & "Pending Response"
    For j = 0 To UBound(varApps, 2)
        lstWorkList.AddItem varApps(0, j) & "," & varApps(4, j) & "," & varApps(2, j) & "," & varApps(22, j) & "," & varApps(23, j)
    Next j
End Sub
Private Sub lstWorkList_DblClick(Cancel As Integer)
    Dim recordId As Long
    Dim row As String
    Dim indexId As Long
    
    If lstWorkList.RowSource = "" Then
        Exit Sub
    Else
        recordId = lstWorkList.Column(0, lstWorkList.ListIndex + 1)
    End If
    
    For j = 0 To UBound(varApps, 2)
        If varApps(0, j) = recordId Then
            indexId = j
            Exit For
        End If
    Next j
    
    recordCounter = j
    Call PopulateBoxes(indexId)
End Sub
Private Sub btnChangeAuditType_Click()

    'VALIDATION
    If IsNull(txtId) Or txtId.Value = "" Then Exit Sub
    
On Error GoTo err1:
    If MsgBox("Do you want to change the audit type from 'Renewal' to 'Application'?", vbQuestion + vbYesNo, "Change Audit Type") = vbYes Then
        Dim db As DAO.Database
        Dim rs As DAO.Recordset
        Set db = CurrentDb
        Set rs = db.OpenRecordset("tblElementJunc", dbOpenDynaset)
        rs.FindFirst ("[header_id] = " & txtId)
        
        'delete existing elements
        If Not rs.NoMatch Then
            With rs
                Do While Not .NoMatch
                    .Edit
                    .Delete
                    .FindNext ("[header_id] = " & txtId)
                Loop
            End With
        End If
        rs.Close:
        
        Set rs = db.OpenRecordset("tblHeader", dbOpenDynaset)
        rs.FindFirst ("[header_id] = " & txtId)
        
        'update type from application to renewal
        If Not rs.NoMatch Then
            With rs
                .Edit
                    ![audit_type] = "Application"
                .Update
            End With
        End If
        rs.Close: Set rs = Nothing
        db.Close: Set db = Nothing
        
        Call btnSearch_Click
        MsgBox ("Audit type changed.")
    Else
        Exit Sub
    End If

    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "Apps; btnSubmit_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
    
End Sub
Private Sub ClearForm()
    If Not IsEmpty(varApps) Then: Set varApps = Nothing
    txtErrorBar.BackColor = RGB(255, 255, 255)
    txtAuditStatus.BackColor = RGB(251, 252, 252)
    recordMax = 0
    lstWorkList.RowSource = ""
    bolChanged = False
    With Me
        Dim ctrl As Control
        For Each ctrl In Controls
            If TypeOf ctrl Is TextBox Then
                ctrl.Value = Null
                ctrl.Enabled = True
            ElseIf TypeOf ctrl Is ComboBox And ctrl.Name <> "cboWeeks" Then
                ctrl.Value = Null
                ctrl.Enabled = True
            End If
        Next
    End With
End Sub
Private Sub ClearElements()
    With Me
        Dim ctrl As Control
        For Each ctrl In Controls
            If Left(ctrl.Name, 1) = "r" Then
                ctrl.Value = Null
            End If
        Next
    End With
End Sub
Private Sub SetListeners()
    Dim ctrl As Control
    For Each ctrl In Controls
        If TypeOf ctrl Is TextBox Then
            ctrl.AfterUpdate = "=MarkChanged()"
        ElseIf TypeOf ctrl Is ComboBox And ctrl.Name <> "cboWeeks" Then
            ctrl.AfterUpdate = "=MarkChanged()"
        End If
    Next
End Sub
Private Function CheckUnsavedChanges() As Boolean
    If bolChanged = True Then
        If MsgBox("You have unsaved changes to a record. Continue?", vbYesNo + vbQuestion, "Unsaved Changes") = vbYes Then
            CheckUnsavedChanges = True
            bolChanged = False
        Else
            CheckUnsavedChanges = False
        End If
    Else
        CheckUnsavedChanges = True
    End If
End Function
Private Function MarkChanged()
    bolChanged = True
End Function
Public Function ValidateForm(submitType As Integer) As Boolean

    Dim msgStr As String, headerStr As String, footerStr As String, ctlName As String
    headerStr = "<ul>": footerStr = "</ul>"
    Dim varFocus As Variant

    Select Case submitType
        Case 1 'SEARCH
            If cboWeeks.Value = "" Or IsNull(cboWeeks.Value) = True Then
                msgStr = msgStr & "<li><b>Week Of</b> cannot be blank.</li>"
                ctlName = ctlName & "cboWeeks,"
            End If

        Case 2 'SUBMIT
            If cboUnableToReview.Value = 1 Then 'unable to review
            ElseIf cboUnableToReview.Value = 0 Then 'able to review, required fields must be completed
                Dim nullField As Boolean: nullField = False
                If R3.Value = "" Or IsNull(R3) Then: nullField = True: ctlName = ctlName & "r3,"
                If r6.Value = "" Or IsNull(r6) Then: nullField = True: ctlName = ctlName & "r6,"
                If r7.Value = "" Or IsNull(r7) Then: nullField = True: ctlName = ctlName & "r7,"
                If r8.Value = "" Or IsNull(r8) Then: nullField = True: ctlName = ctlName & "r8,"
                If r9.Value = "" Or IsNull(r9) Then: nullField = True: ctlName = ctlName & "r9,"
                If r14.Value = "" Or IsNull(r14) Then: nullField = True: ctlName = ctlName & "r14,"
                If r15.Value = "" Or IsNull(r15) Then: nullField = True: ctlName = ctlName & "r15,"
                If r16.Value = "" Or IsNull(r16) Then: nullField = True: ctlName = ctlName & "r16,"
                If r17.Value = "" Or IsNull(r17) Then: nullField = True: ctlName = ctlName & "r17,"
                If r18.Value = "" Or IsNull(r18) Then: nullField = True: ctlName = ctlName & "r18,"
                If r20.Value = "" Or IsNull(r20) Then: nullField = True: ctlName = ctlName & "r20,"
                If r22.Value = "" Or IsNull(r22) Then: nullField = True: ctlName = ctlName & "r22,"
                If r23.Value = "" Or IsNull(r23) Then: nullField = True: ctlName = ctlName & "r23,"
                If r24.Value = "" Or IsNull(r24) Then: nullField = True: ctlName = ctlName & "r24,"
                If r25.Value = "" Or IsNull(r25) Then: nullField = True: ctlName = ctlName & "r25,"
                If nullField = True Then
                    msgStr = msgStr & "<li>All required elements must be answered.</li>"
                    ctlName = ctlName & "cboWeeks,"
                End If
            ElseIf cboUnableToReview.Value = "" Or IsNull(cboUnableToReview) Then
                msgStr = msgStr & "<li><b>Unable to Review</b> cannot be blank.</li>"
                ctlName = ctlName & "cboUnableToReview,"
            End If
            If cboPendingResponse.Value = "" Or IsNull(cboPendingResponse) Then
                msgStr = msgStr & "<li><b>Pending Response</b> cannot be blank.</li>"
                ctlName = ctlName & "cboPendingResponse,"
            End If
            If txtComments.Value = "" Or IsNull(txtComments) _
                And (R3.Value = "No" _
                Or r6.Value = "No" _
                Or r7.Value = "No" _
                Or r8.Value = "No" _
                Or r9.Value = "No" _
                Or r14.Value = "No" _
                Or r15.Value = "No" _
                Or r16.Value = "No" _
                Or r17.Value = "No" _
                Or r18.Value = "No" _
                Or r20.Value = "No" _
                Or r22.Value = "No" _
                Or r23.Value = "No" _
                Or r24.Value = "No" _
                Or r25.Value = "YES") _
            Then
                msgStr = msgStr & "<li><b>Comments</b> cannot be blank when submitting an element with an error.</li>"
                ctlName = ctlName & "txtComments,"
            End If
            
        Case 3 'CHECK USER ID
            If userId = 0 Or IsNull(userId) Then
                msgStr = msgStr & "<li><b>Audit Status</b> must be changed before submitting.</li>"
                ctlName = ctlName & "txtAuditStatus,"
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
