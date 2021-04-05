# SourceControl.ImportProjectComponents Method

Imports and saves project's components, from `Path` directory to the specified `Target`.

```vb
Public Sub ImportProjectComponents(ByVal Target As Workbook, ByVal Path As String)
```

### Parameters

**Target** `Workbook` <br>
Workbook to which, components will be imported.

**Path** `String` <br>
Location of directory, from which components will be stored imported.

### Errors

`OnInvalidOperation` <br>
`Target` project is protected.

`OnDirectoryNotFound` <br>
`Path` does not exists or is unreachable.

