Attribute VB_Name = "TestResultsTests"
'@Folder("Tests")
Option Explicit

Private Const ModuleName As String = "TestResultsTests"


Public Sub Start()
    
    Add
    AddTestResultsWithSameSource
    AddThrowsArgumentNullExceptionWhenItemIsNothing
    AddThrowsArgumentExceptionWhenItemSourceIsVbNullString
    LoopingOverTestResultsTest
    
End Sub


Private Sub Add()

    On Error GoTo ErrHandler
    Const MethodName = "Add"
    
    ' Arrange
    Dim Item As New TestResult
    With Item
        .Source = "Example.Source"
        .Outcome = TestOutcome.Passed
        .Description = "Example"
    End With
    Dim TestResults As New TestResults
    
    ' Act
    TestResults.Add Item
    
    ' Assert
    ExUnit.AreEqual "Example.Source", TestResults.Item("Example.Source").Source, GetSig(MethodName)
    
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
        .Outcome = TestOutcome.Passed
        .Description = "Example"
    End With
    
    Dim Item2 As New TestResult
    With Item2
        .Source = "Example.Source"
        .Outcome = TestOutcome.Failed
        .Description = "Example"
    End With
    
    Dim Item3 As New TestResult
    With Item3
        .Source = "Example.Source"
        .Outcome = TestOutcome.Passed
        .Description = "Example"
    End With
    
    Dim TestResults As New TestResults
    
    ' Act
    TestResults.Add Item1
    TestResults.Add Item2
    TestResults.Add Item3
    
    ' Assert
    ExUnit.AreEqual 1, TestResults.Count, GetSig(MethodName)
    ExUnit.AreEqual "Example.Source", TestResults.Item("Example.Source").Source, GetSig(MethodName)
    ExUnit.AreEqual TestOutcome.Failed, TestResults.Item("Example.Source").Outcome, GetSig(MethodName)
    
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
        .Outcome = TestOutcome.Passed
        .Description = "Example"
    End With
    Dim TestResults As New TestResults
    
    ' Act
    TestResults.Add Item

ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentException, Err.Number, GetSig(MethodName)

End Sub


Private Sub LoopingOverTestResultsTest()

    On Error GoTo ErrHandler
    Const MethodName = "LoopingOverTestResultsTest"

    ' Arrange
    Dim Tr1 As New TestResult
    Tr1.Source = "Source1"
    Dim Tr2 As New TestResult
    Tr2.Source = "Source2"
    
    ' Act
    Dim Results As New TestResults
    Results.Add Tr1
    Results.Add Tr2

    ' Assert
    Dim Key As Variant
    For Each Key In Results.NewEnum
    Next Key

    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


