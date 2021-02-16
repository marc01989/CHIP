VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command0_Click()


    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim varApps As Variant
    Dim strPath As String
    Dim fs As Object, a As Object
    Dim elName As String
    Dim elNum As Integer
    
    elName = InputBox("enter element name")
    elNum = InputBox("enter element number")
    elName = "[" & elName & "]"
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT [Record#], " & elNum & ", " & elName & " FROM [CHIP_RENEWAL OF COVERAGE] WHERE " & elName & " IS NOT NULL;")

'    strPath = "C:\Users\carusomj\OneDrive - UPMC\Projects\CHIP\data\"
'
'    Set fs = CreateObject("Scripting.FileSystemObject")
'    If fs.FileExists(strPath & "CHIPEle" & elNum & ".txt") = True Then
'        Set a = fs.Opentextfile(strPath & "CHIPEle" & elNum & ".txt", 8)
'    Else
'        Set a = fs.createtextfile(strPath & "CHIPEle" & elNum & ".txt")
'    End If

    With rs
        .MoveFirst
        .MoveLast
        .MoveFirst
        If .RecordCount > 0 Then
            varApps = .GetRows(.RecordCount)
        End If
    End With
    rs.Close
    
    Set rs = db.OpenRecordset("ren_elements_junc", dbOpenDynaset)
    With rs
    
        For j = 0 To UBound(varApps, 2)
        
            .AddNew
                ![Record#] = varApps(0, j)
                ![element_id] = varApps(1, j)
                ![result] = varApps(2, j)
            .Update
            
'            For i = 0 To UBound(varApps) - LBound(varApps)
'                row = row & "," & varApps(i, j)
'            Next i
            'Debug.Print Right(row, Len(row) - 1)
'            a.writeline Right(row, Len(row) - 1)
            row = ""
        Next j
        
   End With


    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing
    
    MsgBox ("complete")
        
'    a.Close
'    Set fs = Nothing
    
    
    
    
End Sub
