Attribute VB_Name = "ExUnit"
Option Explicit
'@Folder("ExUnit")

Private Const ModuleName As String = "ExUnit"
Private Const FailedTag As String = "FAILED"
Private Const PassedTag As String = "PASSED"
Private pShowPassingTest As Boolean
Private pTestResults As TestResults


Public Sub Setup()
    Set pTestResults = New TestResults
End Sub


Public Sub Teardown()
    Set pTestResults = Nothing
End Sub


' Gets or sets if passing tests should be displayed.
Public Property Get ShowPassingTest() As Boolean
    ShowPassingTest = pShowPassingTest
End Property


Public Property Let ShowPassingTest(ByVal Value As Boolean)
    pShowPassingTest = Value
End Property


' Gets results of last test run.
Public Property Get TestResults() As TestResults
    Set TestResults = pTestResults
End Property


' Verifies that two value types values are equivalent.
Public Sub AreEqual(ByVal Expected As Variant, _
                    ByVal Actual As Variant, _
                    ByVal Source As String, _
                    Optional ByVal Msg As String)
    
    If IsObject(Expected) Or IsObject(Actual) Then
        Exception.ArgumentException vbNullString, _
                                    "AreEqual supports only value type comparisons." _
                                    & ModuleName & ".NotEqual"
    End If

    IsTrue Expected = Actual, Source, FormatExpectedAndActualValues(Expected, Actual) & Msg
    
End Sub


' Verifies that two value types values are not equivalent.
Public Sub AreNotEqual(ByRef Expected As Variant, _
                       ByRef Actual As Variant, _
                       ByVal Source As String, _
                       Optional ByVal Msg As String)
    
    If IsObject(Expected) Or IsObject(Actual) Then
        Exception.ArgumentException vbNullString, _
                                    "AreNotEqual supports only value type comparisons." _
                                    & ModuleName & ".NotEqual"
    End If
    
    IsTrue Expected <> Actual, Source, FormatExpectedAndActualValues(Expected, Actual) & Msg
    
End Sub


Private Function FormatExpectedAndActualValues(ByVal Expected As String, _
                                               ByVal Actual As String) As String

    FormatExpectedAndActualValues = "Expected value [" & Expected & "], Actual value [" & Actual & "]"
End Function


Public Sub IsFalse(ByVal Eval As Boolean, ByVal Source As String, Optional ByVal Msg As String)
    IsTrue Not (Eval), Source, Msg
End Sub


Public Sub IsTrue(ByVal Eval As Boolean, ByVal Source As String, Optional ByVal Msg As String)
    
    If Eval = True Then
        TestPass Source, Msg
    Else
        TestFail Source, Msg
    End If

End Sub


Public Sub TestPass(ByVal Source As String, Optional ByVal Msg As String)
    
    If IsAdhocRun Then
        PrintTestResult PassedTag, Source, Msg
        
    ElseIf ShowPassingTest Then
        pTestResults.Add CreateTestResult(Source, TestOutcome.Passed, Msg)
        
    End If

End Sub


Private Function CreateTestResult(ByVal Source As String, _
                                  ByVal Outcome As TestOutcome, _
                                  ByVal Description As String) As TestResult

    Dim Output As New TestResult
    With Output
        .Source = Source
        .Result = Outcome
        .Description = Description
    End With
    Set CreateTestResult = Output

End Function


Private Sub PrintTestResult(ByVal Result As String, _
                            ByVal Source As String, _
                            Optional ByVal Msg As String)
    
    Dim Printer As New TestResultImmediatePrinter
    
    If Result = PassedTag Then
        Printer.PrintSingle CreateTestResult(Source, TestOutcome.Passed, Msg)
    Else
        Printer.PrintSingle CreateTestResult(Source, TestOutcome.Failed, Msg)
    End If
    
End Sub


Public Sub TestFail(ByVal Source As String, Optional ByVal Msg As String)

    If IsAdhocRun Then
        PrintTestResult FailedTag, Source, Msg
    Else
        pTestResults.Add CreateTestResult(Source, TestOutcome.Failed, Msg)
    End If
    
End Sub


' AdhocRun means that test is ran from the test method.
Private Function IsAdhocRun() As Boolean
    IsAdhocRun = (pTestResults Is Nothing)
End Function


Public Sub TestFailRunTime(ByVal Source As String, Optional ByVal Msg As String)
    
    If IsAdhocRun Then
        Debug.Print Err.Description, Source
        PrintTestResult FailedTag, Source, Msg
    Else
        pTestResults.Add CreateTestResult(Source, TestOutcome.Failed, Msg)
    End If
    
End Sub


' Verifies that a string contains a given sub-string.
Public Sub ContainsSubstring(ByVal Expected As String, _
                             ByVal Actual As String, _
                             ByVal Source As String, _
                             Optional ByVal Msg As String)

    IsTrue StringH.Contains(Expected, Actual), Source, Msg
End Sub


'Verifies that a collection contains a given object, using an equality comparer.
Public Sub Contains(ByVal Expected As Variant, _
                    ByRef Items As Collection, _
                    ByRef Comparer As IVBAEqualityComparer, _
                    ByVal Source As String, _
                    Optional ByVal Msg As String)
                    
    IsTrue CollectionH.Contains(Expected, Items, Comparer), Source, Msg
End Sub


'Verifies that a collection contains a given object, using an equality comparer.
Public Sub DoesNotContains(ByVal Expected As Variant, _
                           ByRef Items As Collection, _
                           ByRef Comparer As IVBAEqualityComparer, _
                           ByVal Source As String, _
                           Optional ByVal Msg As String)
                           
    IsFalse CollectionH.Contains(Expected, Items, Comparer), Source, Msg
End Sub


' Verifies that an object reference is not nothing.
Public Sub IsNotNothing(ByVal Obj As Object, _
                        ByVal Source As String, _
                        Optional ByVal Msg As String)

    IsTrue Not (Obj Is Nothing), Source, Msg
End Sub


' Verifies that an object reference is nothing.
Public Sub IsNothing(ByVal Obj As Object, ByVal Source As String, Optional ByVal Msg As String)
    IsTrue Obj Is Nothing, Source, Msg
End Sub


' Verifies that two objects are the same instance.
Public Sub AreSame(ByVal Expected As Object, _
                   ByVal Actual As Object, _
                   ByVal Source As String, _
                   Optional ByVal Msg As String)

    IsTrue Expected Is Actual, Source, Msg
End Sub


' Verifies that two objects are not the same instance.
Public Sub AreNotSame(ByVal Expected As Object, _
                      ByVal Actual As Object, _
                      ByVal Source As String, _
                      Optional ByVal Msg As String)

    IsTrue Not (Expected Is Actual), Source, Msg
End Sub


' Verifies that expected error code matches actual error code.
Public Sub IsException(ByVal Expected As ExceptionCode, _
                       ByVal Actual As Long, _
                       ByVal Source As String, _
                       Optional ByVal Msg As String)
    
    IsTrue Expected = Actual, Source, FormatExpectedAndActualExceptions(Expected, Actual) & Msg
End Sub


Private Function FormatExpectedAndActualExceptions(ByVal Expected As Long, _
                                                   ByVal Actual As Long) As String

    Dim ExpectedExpName As String
    ExpectedExpName = GetExceptionName(Expected)
    
    Dim ActualExpName As String
    ActualExpName = GetExceptionName(Actual)
    
    FormatExpectedAndActualExceptions = "Expected exception [" & ExpectedExpName & "], Actual exception [" & ActualExpName & "]"

End Function


' Returns name of the exception based on the ExceptionCode
' or VBAErrorCode enumerators.
' Some errors may not be defined, then
' number of the error (ExpCode) will be returned instead.
Private Function GetExceptionName(ByVal ExpCode As Long) As String

    Dim Output As String
    
    If ExceptionCodeEnum.TryToString(ExpCode, Output) Then
        GetExceptionName = Output
        
    ElseIf ErrorNumberEnum.TryToString(ExpCode, Output) Then
        GetExceptionName = Output
        
    Else
        GetExceptionName = ExpCode
    End If

End Function


Public Sub PrintTestResults(ByVal Printer As ITestResultPrinter)
    
    If IsAdhocRun Then
        Exception.InvalidOperationException vbNullString, ModuleName & ".PrintTestResults"
    End If
    Printer.PrintMany pTestResults
    
End Sub
