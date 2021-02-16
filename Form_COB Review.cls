VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_COB Review"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim varCom As Variant
Dim recordMax As Long, recordCounter As Long
Dim userId As Integer
Dim bolChanged As Boolean




Private Sub btnClear_Click()
    Call ClearForm
End Sub
Private Sub btnClose_Click()
    If CheckUnsavedChanges = False Then Exit Sub
    DoCmd.Close acForm, "COB Review", acSaveNo
End Sub


Private Sub Form_Load()
    Dim obj As AccessObject, dbs As Object
    Dim args As String
    cboMonth.SetFocus
    txtAuditStatus.BackColor = RGB(255, 255, 255)
    
    'checks if value is passed when opening form, runs main query if so
    If Not IsNull(Forms![COB Review].OpenArgs) Then
        args = Forms![COB Review].OpenArgs
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
                'Call btnClearForm_Click
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
Private Sub btnSearch_Click()
    
    'VALIDATION
    If ValidateForm(1) = True Then
        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            txtErrorBar.SetFocus
            Exit Sub
        End If
    End If
    'CHECK FOR UNSAVED CHANGES
    If CheckUnsavedChanges = False Then Exit Sub
    
On Error GoTo err1:
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim monthOf As String: monthOf = cboMonth.Value
    Dim lob As String: lob = cboLOB.Value
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM tblHeaderCOB WHERE month_of = '" & monthOf & "' AND type = '" & lob & "' ;", dbReadOnly)
    
    With rs
        If .RecordCount > 0 Then
            .MoveFirst
            .MoveLast
            .MoveFirst
            recordMax = .RecordCount - 1
            recordCounter = 0
            varCom = .GetRows(.RecordCount)
            
            Call PopulateBoxes(recordCounter)
        Else
            MsgBox ("No results")
            Call ClearForm
            Exit Sub
        End If
    End With
    
    bolChanged = False
    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "COB Review; btnSearch_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
    
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
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim recordId As Long: recordId = txtId.Value
    Dim subDate As String: subDate = Now()
    Dim ctlArray As Variant
    Const auditStatus As String = "Complete"
    

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT element_id FROM tblElementsCOB WHERE is_active = true;", dbReadOnly)
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
    
    
    Set rs = db.OpenRecordset("tblHeaderCOB", dbOpenDynaset)
    With rs
        .FindFirst ("[header_id] = " & recordId)
        If Not .NoMatch Then
            .Edit
                ![date_audited] = subDate
                ![audit_status] = auditStatus
                ![qa_comments] = txtErrorDescription.Value
                ![auditor_id] = userId
                txtAuditStatus.Value = auditStatus
                txtAuditStatus.BackColor = RGB(144, 238, 144)
                txtAuditDate.Value = subDate
                txtAuditorName.Value = DLookup("auditor", "tblAuditors", "auditor_id = " & userId)
            .Update
            .Close
            
            'UPDATE LOCAL ARRAY
            varCom(25, recordCounter) = subDate
            varCom(26, recordCounter) = auditStatus
            varCom(27, recordCounter) = txtErrorDescription.Value
            varCom(28, recordCounter) = userId
        Else
            MsgBox ("Cannot submit record")
        End If
    End With
    
    Set rs = db.OpenRecordset("tblElementJuncCOB", dbOpenDynaset)
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
                ctlName = ctlArray(0, i) 'element_id IS ALSO NAME OF CONTROLS ON FORM FOR EACH ELEMENT
    
                .AddNew
                    ![header_id] = recordId
                    ![element_id] = ctlArray(0, i)
                    ![result] = Me.Controls(ctlName).Value & vbNullString
                .Update
            Next i
        End With
    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing
    
    bolChanged = False
    MsgBox ("Item submitted")
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "COB Review; btnSubmit_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select

End Sub

Public Sub PopulateBoxes(i As Long)

On Error GoTo err1:
    txtId.Value = varCom(0, i)
    txtMemberNbrCHIP.Value = varCom(1, i)
    txtEmployeeNbrCHIP.Value = varCom(2, i)
    txtLastNameCHIP.Value = varCom(3, i)
    txtFirstNameCHIP.Value = varCom(4, i)
    txtEffDateCHIP.Value = FormatDates(varCom(5, i))
    txtTermDateCHIP.Value = FormatDates(varCom(6, i))
    txtMemberNbrCom.Value = varCom(9, i)
    txtLastNameCom.Value = varCom(10, i)
    txtFirstNameCom.Value = varCom(11, i)
    txtSubscriberNameCom.Value = varCom(12, i)
    txtEffDateCom.Value = FormatDates(varCom(16, i))
    txtTermDateCom.Value = FormatDates(varCom(17, i))
    txtGroupCom.Value = varCom(18, i)
    txtSbgrpCom.Value = varCom(19, i)
    txtSpecialistComments.Value = varCom(21, i)
    txtAssignedToUser.Value = varCom(22, i)
    txtAuditDate.Value = varCom(25, i)
    txtAuditStatus.Value = varCom(26, i)
    txtErrorDescription.Value = varCom(27, i)
    If varCom(26, i) = "Complete" Then
        txtAuditStatus.BackColor = RGB(144, 238, 144)
    Else
        txtAuditStatus.BackColor = RGB(255, 99, 71)
    End If
    If Not IsNull(varCom(28, i) And varCom(28, i) <> "") Then
        txtAuditorName.Value = DLookup("auditor", "tblAuditors", "auditor_id = " & varCom(28, i))
    Else
        txtAuditorName.Value = Null
    End If
    txtRecordCounter.Value = i + 1
    txtRecordMax.Value = recordMax + 1
    
    Dim recordId As Long: recordId = varCom(0, i)
    Call GetElements(recordId)
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "COB Review; PopulateBoxes()")
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
    
On Error GoTo err1:
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM [tblElementJuncCOB] WHERE [header_id] = " & recordId & " ORDER BY element_id ASC;", dbReadOnly)
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
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "COB Review; GetElements()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
    
End Sub
Private Sub btnBack_Click()
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

Private Sub ClearForm()
    If Not IsEmpty(varCom) Then: Set varApps = Nothing
    txtErrorBar.BackColor = RGB(255, 255, 255)
    txtAuditStatus.BackColor = RGB(255, 255, 255)
    recordMax = 0
    bolChanged = False

    With Me
        Dim ctrl As Control
        For Each ctrl In Controls
            If TypeOf ctrl Is TextBox Then
                ctrl.Value = Null
                ctrl.Enabled = True
            ElseIf TypeOf ctrl Is ComboBox And ctrl.Name <> "cboMonth" And ctrl.Name <> "cboLOB" Then
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
            If Left(ctrl.Name, 1) = "E" Then
                ctrl.Value = Null
            End If
        Next
    End With
End Sub
Public Function FormatDates(ByVal inputDate As String) As String
    Dim formattedDate As String
    If inputDate <> "21991231" Then
        formattedDate = Mid(inputDate, 5, 2) & "/" & Right(inputDate, 2) & "/" & Left(inputDate, 4)
        FormatDates = formattedDate
    Else
        FormatDates = ""
    End If
End Function
Private Sub SetListeners()
    Dim ctrl As Control
    For Each ctrl In Controls
        If TypeOf ctrl Is TextBox Then
            ctrl.AfterUpdate = "=MarkChanged()"
        ElseIf TypeOf ctrl Is ComboBox And ctrl.Name <> "cboMonth" And ctrl.Name <> "cboLOB" Then
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
            If cboMonth.Value = "" Or IsNull(cboMonth.Value) = True Then
                msgStr = msgStr & "<li><b>Sample Period</b> cannot be blank.</li>"
                ctlName = ctlName & "cboMonth,"
            End If
             If cboLOB.Value = "" Or IsNull(cboLOB.Value) = True Then
                msgStr = msgStr & "<li><b>LOB</b> cannot be blank.</li>"
                ctlName = ctlName & "cboLOB,"
            End If

        Case 2 'SUBMIT
            Dim nullField As Boolean: nullField = False
            If e1.Value = "" Or IsNull(e1) Then: nullField = True: ctlName = ctlName & "e1,"
            If e2.Value = "" Or IsNull(e2) Then: nullField = True: ctlName = ctlName & "e2,"
            If e3.Value = "" Or IsNull(e3) Then: nullField = True: ctlName = ctlName & "e3,"
            If e4.Value = "" Or IsNull(e4) Then: nullField = True: ctlName = ctlName & "e4,"
            If e7.Value = "" Or IsNull(e7) Then: nullField = True: ctlName = ctlName & "e7,"
            If e8.Value = "" Or IsNull(e8) Then: nullField = True: ctlName = ctlName & "e8,"
            If e10.Value = "" Or IsNull(e10) Then: nullField = True: ctlName = ctlName & "e10,"
            If nullField = True Then
                msgStr = msgStr & "<li>All required elements must be answered.</li>"
                ctlName = ctlName & "cboWeeks,"
            End If
            
           If txtErrorDescription.Value = "" Or IsNull(txtErrorDescription) _
                And (e1.Value = "NO" _
                Or e2.Value = "NO" _
                Or e3.Value = "NO" _
                Or e4.Value = "NO" _
                Or e7.Value = "NO" _
                Or e8.Value = "NO" _
                Or e10.Value = "NO") _
            Then
                msgStr = msgStr & "<li><b>QA Comments</b> cannot be blank when submitting an element with an error.</li>"
                ctlName = ctlName & "txtErrorDescription,"
            End If

            
        Case 3 'CHECK USER ID
            If userId = 0 Or IsNull(userId) Then
                msgStr = msgStr & "<li>Session timed out - please log in again before submitting.</li>"
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

