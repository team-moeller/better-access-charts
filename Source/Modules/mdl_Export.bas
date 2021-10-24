Attribute VB_Name = "mdl_Export"
Option Compare Database
Option Explicit

  Private Const VB_MODULE               As Integer = 1
  Private Const VB_CLASS                As Integer = 2
  Private Const VB_FORM                 As Integer = 100
  
  Private Const EXT_TABLE               As String = ".tbl"
  Private Const EXT_QUERY               As String = ".qry"
  Private Const EXT_MODULE              As String = ".bas"
  Private Const EXT_CLASS               As String = ".cls"
  Private Const EXT_FORM                As String = ".frm"
  Private Const EXT_MACRO               As String = ".mcr"
  
  Private Const FLD_SOURCE              As String = "Source"


Public Sub saveAllAsText()
  Dim oDatabase               As DAO.Database
  Dim oTable                  As TableDef
  Dim oQuery                  As QueryDef
  Dim oCont                   As Container
  Dim oForm                   As Document
  Dim oMacro                  As Object
  Dim oModule                 As Object
  Dim oFSO                    As Object

  Dim strBasePath             As String
  Dim strPath                 As String
  Dim strName                 As String
  Dim strFileName             As String
  
  On Error GoTo errHandler
  
  Set oDatabase = CurrentDb
  
  Set oFSO = CreateObject("Scripting.FileSystemObject")
  strBasePath = addFolder(oFSO, CurrentProject.Path, FLD_SOURCE)

  strPath = addFolder(oFSO, strBasePath, "Tables")
  For Each oTable In oDatabase.TableDefs
    strName = oTable.Name
    If Left(strName, 4) <> "MSys" Then
      strFileName = strPath & "\" & strName & EXT_TABLE
      Application.ExportXML acExportTable, strName, strFileName, , , , acUTF8, acEmbedSchema + acExportAllTableAndFieldProperties
    End If
  Next
  
  strPath = addFolder(oFSO, strBasePath, "Queries")
  For Each oQuery In oDatabase.QueryDefs
    strName = oQuery.Name
    If Left(strName, 1) <> "~" Then
      strFileName = strPath & "\" & strName & EXT_QUERY
      Application.SaveAsText acQuery, strName, strFileName
    End If
  Next
  
  strPath = addFolder(oFSO, strBasePath, "Forms")
  Set oCont = oDatabase.Containers("Forms")
  For Each oForm In oCont.Documents
    strName = oForm.Name
    strFileName = strPath & "\" & strName & EXT_FORM
    Application.SaveAsText acForm, strName, strFileName
    CleanupForm strFileName
  Next
  
  strPath = addFolder(oFSO, strBasePath, "Macros")
  Set oCont = oDatabase.Containers("Scripts")
  For Each oMacro In oCont.Documents
    strName = oMacro.Name
    strFileName = strPath & "\" & strName & EXT_MACRO
    Application.SaveAsText acForm, strName, strFileName
  Next
  
  
  strPath = addFolder(oFSO, strBasePath, "Modules")
  For Each oModule In Application.VBE.ActiveVBProject.VBComponents
    strName = oModule.Name
    strFileName = strPath & "\" & strName
    Select Case oModule.Type
      Case VB_MODULE
        oModule.Export strFileName & EXT_MODULE
      Case VB_CLASS
        oModule.Export strFileName & EXT_CLASS
      Case VB_FORM
        ' Do not export form modules (already exported the complete forms)
      Case Else
        Debug.Print "Unknown module type: " & oModule.Type, oModule.Name
    End Select
  Next
  
Exit Sub

errHandler:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf
  Stop: Resume

End Sub

'
' Create a folder when necessary. Append the folder name to the given path.
'
Private Function addFolder(ByRef fso As Object, ByVal strPath As String, ByVal strAdd As String) As String
  addFolder = strPath & "\" & strAdd
  If Not fso.FolderExists(addFolder) Then MkDir addFolder
End Function

Private Sub CleanupForm(ByVal strPath As String)
  Const ForReading = 1
  Const ForWriting = 2
  Const asUnicode = -1
  
  Dim fso         As Object ' FileSystemObject
  Dim sourceFile  As Object ' TextStream
  Dim destFile    As Object ' TextStream
  
  Dim line As String
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set sourceFile = fso.OpenTextFile(strPath, ForReading, , Format:=asUnicode)
  Set destFile = fso.OpenTextFile(strPath & "-tmp", ForWriting, Create:=True, Format:=asUnicode)
  
  While Not sourceFile.AtEndOfStream
    line = sourceFile.ReadLine
    
    If line Like "Version =21" Then
      line = "Version =20"   ' Bug in MSAccess 2016 ??
    ElseIf (line Like "*PrtMip = *") _
        Or (line Like "*PrtDevMode = *") Or (line Like "*PrtDevModeW = *") _
        Or (line Like "*PrtDevNames = *") Or (line Like "*PrtDevNamesW = *") Then
      While line <> "    End"
        line = sourceFile.ReadLine
      Wend
      line = vbNullString
    ElseIf (line Like "*LayoutCached*") Then
      line = vbNullString
    ElseIf (line Like "*WebImagePadding*") Then
      line = vbNullString
    End If
    
    If line <> vbNullString Then
      destFile.Writeline line
    End If
  Wend
  sourceFile.Close
  destFile.Close
  
  fso.DeleteFile strPath
  fso.MoveFile strPath & "-tmp", strPath
End Sub
