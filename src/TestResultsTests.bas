Attribute VB_Name = "TestResultsTests"
'@Folder("Tests")
Option Explicit

Private Const ModuleName As String = "TestResultsTests"


Public Sub Start()
    
    Add
    AddTestResultsWithSameSource
    AddThrowsArgumentNullExceptionWhenItemIsNothing
    AddThrowsArgumentExceptionWhenItemSourceIsVbNullString
    
End Sub


Private Sub Add()

    On Error GoTo ErrHandler
    Const MethodName = "Add"
    
    ' Arrange
    Dim Item As New TestResult
    With Item
        .Source = "Example.Source"
        .Result = TestOutcome.Passed
        .Description = "Example"
    End With
    Dim TestResults As New TestResults
    
    ' Act
    TestResults.Add Item
    
    ' Assert
    ExUnit.AreEqual "Example.Source", TestResults.Enumerator.Item("Example.Source").Source, GetSig(MethodName)
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Function GetSig(ByVal MethodName As String) As String
    GetSig = ModuleName & "." & MethodName
End Function


Private Sub AddTestResultsWithSameSource()

    On Error GoTo ErrHandler
    Const MethodName = "AddTestResultsWithSameSource"
    
    ' Arrange
    Dim Item1 As New TestResult
    With Item1
        .Source = "Example.Source"
        .Result = TestOutcome.Passed
        .Description = "Example"
    End With
    
    Dim Item2 As New TestResult
    With Item2
        .Source = "Example.Source"
        .Result = TestOutcome.Failed
        .Description = "Example"
    End With
    
    Dim Item3 As New TestResult
    With Item3
        .Source = "Example.Source"
        .Result = TestOutcome.Passed
        .Description = "Example"
    End With
    
    Dim TestResults As New TestResults
    
    ' Act
    TestResults.Add Item1
    TestResults.Add Item2
    TestResults.Add Item3
    
    ' Assert
    ExUnit.AreEqual 1, TestResults.Count, GetSig(MethodName)
    ExUnit.AreEqual "Example.Source", TestResults.Enumerator.Item("Example.Source").Source, GetSig(MethodName)
    ExUnit.AreEqual TestOutcome.Failed, TestResults.Enumerator.Item("Example.Source").Result, GetSig(MethodName)
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub AddThrowsArgumentNullExceptionWhenItemIsNothing()
    
    On Error GoTo ErrHandler
    Const MethodName = "AddThrowsArgumentNullExceptionWhenItemIsNothing"
    
    ' Arrange
    Dim TestResults As New TestResults
    
    ' Act
    TestResults.Add Nothing
    
    ' Assert
    Exit Sub
ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentNull, Err.Number, GetSig(MethodName)
    
End Sub


Private Sub AddThrowsArgumentExceptionWhenItemSourceIsVbNullString()

    On Error GoTo ErrHandler
    Const MethodName = "AddThrowsArgumentExceptionWhenItemSourceIsVbNullString"
    
    ' Arrange
    Dim Item As New TestResult
    With Item
        .Source = vbNullString
        .Result = TestOutcome.Passed
        .Description = "Example"
    End With
    Dim TestResults As New TestResults
    
    ' Act
    TestResults.Add Item

ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentException, Err.Number, GetSig(MethodName)

End Sub
