Attribute VB_Name = "Module1"
Option Compare Database

Public Sub LogError(strError, modName As String)
'logs captured errors to txt file
'list of errors
'https://msdn.microsoft.com/en-us/library/bb221208(v=office.12).aspx

    Dim strPath As String, comp As String
    Dim fs As Object
    Dim a As Object

    comp = Environ$("username")
    strPath = "X:\Quality Audit\Quality Audit(DeptUsers)\Technical Team\Gov_Prods\CHIP\Databases\db_utilities"

    Set fs = CreateObject("Scripting.FileSystemObject")
        If fs.FileExists(strPath & "\errorLogCHIP.txt") = True Then
            Set a = fs.Opentextfile(strPath & "\errorLogCHIP.txt", 8)
        Else
            Set a = fs.createtextfile(strPath & "\errorLogCHIP.txt")
        End If
    
        a.writeline "--------------------------------------------------------------------------"
        a.writeline "DATE: " & Date + Time
        a.writeline "ERROR: " & strError
        a.writeline "USER: " & comp
        a.writeline "MODULE: " & modName
        a.writeline "VERSION: 2.2"
        a.writeline "--------------------------------------------------------------------------"
        a.Close
    Set fs = Nothing
End Sub

Public Sub LogUserOff(formModule As String)
'logs off user if their userid is not found or is 0. User must log back in to submit changes to backend tables

    Dim obj As AccessObject, dbs As Object
    
    MsgBox ("User not found - please login again")
    Set dbs = Application.CurrentProject
    For Each obj In dbs.AllForms
        If obj.IsLoaded = True Then
          DoCmd.Close acForm, obj.Name, acSaveNo
        End If
    Next obj
    DoCmd.OpenForm "Login", acNormal, , , , acWindowNormal
    Call LogError(0 & " " & "User Id not found or not passed to home screen", formModule)
End Sub

Public Function Clean(text As Variant)
    Dim scrubbed As String: scrubbed = ""
    
    If Not IsNull(text) Then
        scrubbed = Trim(text)
        scrubbed = Replace(scrubbed, vbTab, "")
        scrubbed = Replace(scrubbed, vbLf, " ")
        scrubbed = Replace(scrubbed, vbCr, " ")
        scrubbed = Replace(scrubbed, vbCrLf, " ")
        scrubbed = Replace(scrubbed, vbNewLine, " ")
        scrubbed = Replace(scrubbed, Chr(160), " ")
    End If

    Clean = scrubbed

End Function

