# SourceControl.ExportProjectComponents Method

Exports and saves project's components, from `Source` workbook to the location which is specified in `Path` argument.

```vb
Public Sub ExportProjectComponents(ByVal Source As VBProject, ByVal Path As String)
```

### Parameters

**Source** `VBProject` <br>
Project from which, components will be exported.

**Path** `String` <br>
Location of directory, where components will be stored when export is completed.

### Errors

`OnInvalidOperation` <br>
`Source` project is protected.

`OnDirectoryNotFound` <br>
`Path` does not exists or is unreachable.

## Examples

The following example demonstrates, who to setup a using which modules will be exported every time when workbooks is saved. In the example, source code will not be exported when component is empty.

```vb
Option Explicit

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    
    ' Let's you control whether components should be exported or not. Useful, while you are 
    ' deploying code onto production.
    Const SourceControlEnabled As Boolean = False
    Dim Path As String
    Path = ThisWorkbook.Path & "/src"
    
    If SourceControlEnabled Then
        If Tools.Fso.FolderExists(Path) = False Then
            DirectoryExt.CreateDirectory Path
            
        ElseIf DirectoryExt.GetFiles(Path).Count > 0 Then
            DirectoryExt.RemoveAllFilesFromFolder Path
            
        End If
        
        SourceControl.ExportEmptyComponents = False
        SourceControl.ExportProjectComponents ThisWorkbook.VBProject, Path
    End If
    
End Sub
```

