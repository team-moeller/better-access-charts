Attribute VB_Name = "mdl_Helpers"
Option Compare Database
Option Explicit

Public Function File2OLE(ByVal Table As String, ByVal PrimaryKeyFieldName As String, _
                         ByVal TargetFieldName As String, ByVal PrimaryKeyValue As Long, _
                         ByVal FileName As String, Optional ByVal InCurrentProjectPath As Boolean) As Long

'Prerequisit: Record with ID must already exist
'Call: File2OLE("USys_FileData","ID","FileData","1","Chart.min.js",True)

    On Error GoTo Handle_Error

    Dim cnn As ADODB.Connection
    Dim strSQL As String
    Dim rst As ADODB.Recordset
    Dim FileID As Long
    Dim PathFilename As String
    Dim Buffer() As Byte
    Dim FileSize As Long

    If InCurrentProjectPath = True Then
        PathFilename = CurrentProject.Path & "\" & FileName
    Else
        PathFilename = FileName
    End If

    If Dir$(PathFilename) = vbNullString Then
        MsgBox "Thge file '" & PathFilename & "' does not exist."
        Exit Function
    End If

    strSQL = "SELECT " & TargetFieldName & ", FileName " & _
             "FROM " & Table & " " & _
             "WHERE " & PrimaryKeyFieldName & " = " & PrimaryKeyValue
    
    Set cnn = CurrentProject.Connection
    Set rst = New ADODB.Recordset
    rst.Open strSQL, cnn, adOpenDynamic, adLockOptimistic
    FileID = FreeFile

    Open PathFilename For Binary Access Read Lock Read Write As FileID

    FileSize = FileLen(PathFilename)
    ReDim Buffer(FileSize)
    rst(TargetFieldName) = Null
    Get FileID, , Buffer
    rst(TargetFieldName).AppendChunk Buffer
    rst("FileName") = FileName
    rst.Update
    Close FileID
    File2OLE = True

Exit_Here:
    rst.Close
    Set rst = Nothing
    Set cnn = Nothing
    Close FileID
    Exit Function

Handle_Error:
    File2OLE = Err.Number
    Resume Exit_Here

End Function

Public Function IsFormOpen(ByVal strFormName As String) As Boolean

    If SysCmd(acSysCmdGetObjectState, acForm, strFormName) <> 0 Then
        If Forms(strFormName).CurrentView <> 0 Then
            IsFormOpen = True
        End If
    End If
    
End Function

Public Sub PrepareAndExportClassModule()

    'Declarations
    Dim Version As String
    Dim CodeLine As String
    
    Version = DLast("V_Number", "tbl_VersionHistory")
    CodeLine = "'# Version " & Version & "  published: " & Format$(Date, "dd.mm.yyyy") & "                                                  #"
    Application.VBE.ActiveVBProject.VBComponents("cls_Better_Access_Chart").CodeModule.InsertLines 4, CodeLine
    Application.VBE.ActiveVBProject.VBComponents("cls_Better_Access_Chart").CodeModule.DeleteLines 5, 1
    
    Application.VBE.ActiveVBProject.VBComponents("cls_better_access_Chart").Export CurrentProject.Path & "\cls_Better_Access_Chart.cls"
    
    MsgBox "Export done", vbInformation, "Better Access Charts"

End Sub
