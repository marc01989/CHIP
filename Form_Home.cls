VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim userId As Integer






Private Sub Form_Load()

    Dim obj As AccessObject, dbs As Object
    Dim args As String

    'checks if value is passed when opening form, runs main query if so
    If Not IsNull(Forms![Home].OpenArgs) Then
        args = Forms![Home].OpenArgs
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
                Call CheckPermission
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

End Sub
Private Sub CheckPermission()
    Dim permLevel As Integer
    permLevel = DLookup("is_manager", "tblAuditors", "auditor_id = " & userId)
    If permLevel = 1 Then
        btnImportSamples.Visible = True
        lblImportSamples.Visible = True
    Else
        btnImportSamples.Visible = False
        lblImportSamples.Visible = False
    End If
End Sub
Private Sub Command142_Click()
    DoCmd.Quit
End Sub
Private Sub Command164_Click()
    DoCmd.OpenForm "Reporting", acNormal
End Sub
Private Sub Command169_Click()
    DoCmd.OpenForm "Apps", acNormal, , , , acFormPropertySettings, userId
End Sub
Private Sub Command171_Click()
    DoCmd.OpenForm "Rens", acNormal, , , , acFormPropertySettings, userId
End Sub
Private Sub Command184_Click()
    DoCmd.OpenForm "COB Review", acNormal, , , , acFormPropertySettings, userId
End Sub
Private Sub Command186_Click()
    DoCmd.OpenForm "COB Review Mkt", acNormal, , , , acFormPropertySettings, userId
End Sub
Private Sub Command179_Click()
    DoCmd.OpenForm "Weekly Rec", acNormal
End Sub
Private Sub btnImportSamples_Click()
    DoCmd.OpenForm "Sampling Page", acNormal
End Sub
Private Sub btnSysInfo_Click()
    DoCmd.OpenForm "Sys Info", acNormal
End Sub


