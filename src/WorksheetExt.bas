Attribute VB_Name = "WorksheetExt"
Option Explicit
'@Folder("Helper")

Private Const ModuleName As String = "WorksheetHelper"

Public Enum SearchLastOrder

    Row = 1
    Column = 2
    Cell = 3
    
End Enum


' Indicates if ListObject exists at the specified Location and with specified name.
' If object with the specified name does not exists in the specified location
' VBA with throw an run-time error ("The item with the specified name wasn't found),
' which will be catched by the ErrHandler and False will be returned.
Public Function ListObjectExistAt(ByRef Location As Worksheet, ByVal ObjectName As String) As Boolean
    
    Const MethodName = "ListObjectExistAt"
    
    If Location Is Nothing Then
        Errors.OnArgumentNull "Location", ModuleName & "." & MethodName
    End If
    
    If ObjectName = vbNullString Then
        Errors.OnArgumentError "ObjectName", "Value can't be an empty string. " & _
                                            ModuleName & "." & MethodName
    End If
    
    On Error GoTo ErrHandler
    ListObjectExistAt = (Location.ListObjects.Item(ObjectName).Name = ObjectName)
    Exit Function
    
ErrHandler:
    On Error GoTo 0
End Function


Public Function TableContainsData(ByRef Table As ListObject) As Boolean

    Const MethodName = "TableContainsData"

    If Table Is Nothing Then
        Errors.OnArgumentNull "Table", ModuleName & "." & MethodName
    End If
    
    If Table.DataBodyRange Is Nothing Then
        TableContainsData = False
        Exit Function
    End If
    
    TableContainsData = (Table.DataBodyRange.Rows.Count > 0)

End Function


' Method clears selected table from items and shrinks table to only first row.
Public Sub ClearTable(ByRef Table As ListObject)

    Const MethodName = "ClearTable"

    If Table Is Nothing Then
        Errors.OnArgumentNull "Table", ModuleName & "." & MethodName
    End If
    
    With Table
        If .Parent.AutoFilterMode = True Then
            .Parent.ShowAllData
        End If
        
        ' Remove rows only if there are more than 1 row.
        If TableContainsData(Table) Then
            .DataBodyRange.Rows.Delete
        End If
        
        ' If table does contains any excess items in first row, remove them.
        If Not .DataBodyRange Is Nothing Then
            If WorksheetFunction.CountA(.DataBodyRange) > 0 Then
                .DataBodyRange.ClearContents
            End If
        End If
    End With
    
End Sub


' Author: Ron de Bruin
Public Function Last(ByVal Order As SearchLastOrder, ByRef Rng As Range) As Long

    Dim Lrw As Long
    Dim Lcol As Long
  
    Select Case Order
  
        Case 1:
            On Error Resume Next
            '@Ignore ExcelMemberMayReturnNothing
            Last = Rng.Find(What:="*", _
                            After:=Rng.Cells.Item(1), _
                            Lookat:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Row
            On Error GoTo 0
  
        Case 2:
            On Error Resume Next
            '@Ignore ExcelMemberMayReturnNothing
            Last = Rng.Find(What:="*", _
                            After:=Rng.Cells.Item(1), _
                            Lookat:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Column
            On Error GoTo 0
  
        Case 3:
            On Error Resume Next
            '@Ignore ExcelMemberMayReturnNothing
            Lrw = Rng.Find(What:="*", _
                           After:=Rng.Cells.Item(1), _
                           Lookat:=xlPart, _
                           LookIn:=xlFormulas, _
                           SearchOrder:=xlByRows, _
                           SearchDirection:=xlPrevious, _
                           MatchCase:=False).Row
            On Error GoTo 0
  
            On Error Resume Next
            '@Ignore ExcelMemberMayReturnNothing
            Lcol = Rng.Find(What:="*", _
                            After:=Rng.Cells.Item(1), _
                            Lookat:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Column
            On Error GoTo 0
  
            On Error Resume Next
            Last = Rng.Parent.Cells(Lrw, Lcol).Address(False, False)
            If Err.Number > 0 Then
                Last = Rng.Cells.Item(1).Address(False, False)
                Err.Clear
            End If
            On Error GoTo 0
  
    End Select
    
End Function



