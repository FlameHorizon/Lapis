Attribute VB_Name = "ExUnitTests"
'@Folder("Tests")
Option Explicit

' This test suite has to be ran from testrunner with
' already existing setup configuration.
Private Const ModuleName As String = "ExUnitTests"
Private Const TestResultSig As String = "Module.Method"


Public Sub Start()
    
    AreEqualFailTest
    AreNotEqualFailTest
    
End Sub


Private Sub AreEqualFailTest()

    On Error GoTo ErrHandler
    Const MethodName = "AreEqualFailTest"

    ' Arrange
    Const ExpectedDescription As String = "Expected value [1], Actual value [2]"
    
    ' Act
    ExUnit.AreEqual 1, 2, TestResultSig
    
    Dim Actual As TestResult
    ' We have to remove this TestResult as it is only relevant inside this method.
    Set Actual = ExUnit.TestResults.Item(TestResultSig)
    ExUnit.TestResults.Remove TestResultSig
    
    ' Assert
    ExUnit.AreEqual ExpectedDescription, Actual.Description, GetSig(MethodName)
    ExUnit.AreEqual TestOutcome.Failed, Actual.Outcome, GetSig(MethodName)
    ExUnit.AreEqual TestResultSig, Actual.Source, GetSig(MethodName)

    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Function GetSig(ByVal MethodName As String) As String
    GetSig = ModuleName & "." & MethodName
End Function


Private Sub AreNotEqualFailTest()

    On Error GoTo ErrHandler
    Const MethodName = "AreNotEqualFailTest"

    ' Arrange
    Const ExpectedDescription As String = "Expected value [1] to not be equal but with the actual, but it isn't."
    
    ' Act
    ExUnit.AreNotEqual 1, 1, TestResultSig
    
    Dim Actual As TestResult
    ' We have to remove this TestResult as it is only relevant inside this method.
    Set Actual = ExUnit.TestResults.Item(TestResultSig)
    ExUnit.TestResults.Remove TestResultSig
    
    ' Assert
    ExUnit.AreEqual ExpectedDescription, Actual.Description, GetSig(MethodName)
    ExUnit.AreEqual TestOutcome.Failed, Actual.Outcome, GetSig(MethodName)
    ExUnit.AreEqual TestResultSig, Actual.Source, GetSig(MethodName)

    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


