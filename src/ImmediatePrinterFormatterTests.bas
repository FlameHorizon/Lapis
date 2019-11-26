Attribute VB_Name = "ImmediatePrinterFormatterTests"
'@Folder("Tests")
Option Explicit

Private Const ModuleName As String = "ImmediatePrinterFormatterTests"


Public Sub Start()
    
    FormatPassed
    FormatFailed

End Sub


Private Sub FormatPassed()

    On Error GoTo ErrHandler
    Const MethodName = "FormatPassed"
    
    ' Arrange
    Dim Item As New TestResult
    With Item
        .Source = "Example.Source"
        .Outcome = TestOutcome.Passed
        .Description = "Example"
    End With
    Dim Formatter As New ImmediatePrinterFormatter
    
    ' Act
    Dim Actual As String
    Actual = Formatter.Format(Item)
    
    ' Assert
    ExUnit.AreEqual "Passed; Example.Source", Actual, GetSig(MethodName)
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Function GetSig(ByVal MethodName As String) As String
    GetSig = ModuleName & "." & MethodName
End Function


Private Sub FormatFailed()
    
    On Error GoTo ErrHandler
    Const MethodName = "FormatPassed"
    
    ' Arrange
    Dim Item As New TestResult
    With Item
        .Source = "Example.Source"
        .Outcome = TestOutcome.Failed
        .Description = "Example"
    End With
    Dim Formatter As New ImmediatePrinterFormatter
    
    ' Act
    Dim Actual As String
    Actual = Formatter.Format(Item)
    
    ' Assert
    ExUnit.AreEqual "Failed; Example; Example.Source", Actual, GetSig(MethodName)
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)
    
End Sub
