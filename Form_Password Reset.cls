VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Password Reset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnUpdate_Click()
    
    'check if both boxes are not null
    If ValidateForm(1) = True Then
        Exit Sub
    End If
    
    'start further validation here
        If ValidateForm(2) = True Then
            Exit Sub
        End If
        
        Dim db As DAO.Database
        Dim rs As DAO.Recordset
        Dim userName As String: userName = txtNetworkId.Value
        Set db = CurrentDb
        Set rs = db.OpenRecordset("tblAuditors", dbOpenDynaset)
            With rs
                .FindFirst ("network_id = '" & userName & "'")
                If .NoMatch Then
                    MsgBox ("User not found. Try again")
                Else
                    .Edit
                    rs![password] = Trim(txtPassword.Value)
                    .Update
                        If MsgBox("Password successfully updated.", vbOKOnly, "Success") = vbOK Then:
                        DoCmd.Close acForm, "Password Reset", acSaveNo
                        DoCmd.OpenForm "Login", acNormal, , , , acWindowNormal
                End If
            End With
        rs.Close: Set rs = Nothing
        db.Close: Set db = Nothing
            
End Sub

Public Function ValidateForm(submitType As Integer) As Boolean

    Dim msgStr As String, headerStr As String, footerStr As String, ctlName As String
    Dim varFocus As Variant

    headerStr = "<ul>"
    footerStr = "</ul>"
    
    Select Case submitType
        Case 1
            '--cannot be blank--
            If txtNetworkId.Value = "" Or IsNull(txtNetworkId.Value) = True Then
                msgStr = msgStr & "<li><b>Network Id</b> cannot be blank.</li>"
                ctlName = "txtNetworkId,"
            End If
            If txtPassword.Value = "" Or IsNull(txtPassword.Value) = True Then
                msgStr = msgStr & "<li><b>New Password</b> cannot be blank.</li>"
                ctlName = ctlName & "txtPassword,"
            End If
            
        Case 2
            '--logged in user must match input network it
            Dim activeUser As String, networkIdInput As String
            Dim matchInt As Integer
            networkIdInput = txtNetworkId.Value
            activeUser = Environ$("username")
            matchInt = StrComp(activeUser, txtNetworkId.Value, vbBinaryCompare)
            If matchInt <> 0 Then
                msgStr = "<li><b>You do not have permission to update the password for this user.<b></li>"
                ctlName = "txtNetworkId,"
            End If
            
    End Select
    
    If msgStr = "" Then 'no errors
        txtErrorBox.Value = Null
        txtErrorBar.Value = Null
        txtErrorBar.BackColor = RGB(255, 255, 255)
        ValidateForm = False
    Else 'errors, show msg to user
        txtErrorBox.Value = headerStr & msgStr & footerStr
        txtErrorBar.Value = "Error"
        txtErrorBar.BackColor = RGB(255, 186, 0)
        varFocus = Split(ctlName, ",")
        Me.Controls(varFocus(0)).SetFocus
        ValidateForm = True
    End If

End Function
