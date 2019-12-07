Attribute VB_Name = "ExUnit"
Option Explicit
'@Folder("ExUnit")

Private Const ModuleName As String = "ExUnit"
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
    
    If Expected = Actual Then
        TestPass Source, Msg
    Else
        TestFail Source, FormatExpectedAndActualValues(Expected, Actual) & Msg
    End If
    
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
    
    If Expected <> Actual Then
        TestPass Source, Msg
    Else
        TestFail Source, "Expected value [" & Expected & "] to not be equal but with the actual, but it is." & Msg
    End If
    
End Sub


Private Function FormatExpectedAndActualValues(ByVal Expected As String, _
                                               ByVal Actual As String) As String

    FormatExpectedAndActualValues = "Expected value [" & Expected & "], Actual value [" & Actual & "]"
End Function


Public Sub IsFalse(ByVal Eval As Boolean, ByVal Source As String, Optional ByVal Msg As String)
    
    If Eval = False Then
        TestPass Source, Msg
    Else
        TestFail Source, FormatExpectedAndActualValues(False, True) & Msg
    End If
        
End Sub


Public Sub IsTrue(ByVal Eval As Boolean, ByVal Source As String, Optional ByVal Msg As String)
    
    If Eval = True Then
        TestPass Source, Msg
    Else
        TestFail Source, FormatExpectedAndActualValues(True, False) & Msg
    End If

End Sub


Public Sub TestPass(ByVal Source As String, Optional ByVal Msg As String)
    
    If IsAdhocRun Then
        PrintTestResult TestOutcome.Passed, Source, Msg
        
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
        .Outcome = Outcome
        .Description = Description
    End With
    Set CreateTestResult = Output

End Function


Private Sub PrintTestResult(ByVal Outcome As TestOutcome, _
                            ByVal Source As String, _
                            Optional ByVal Msg As String)
    
    Dim Printer As New TestResultImmediatePrinter
    Printer.PrintSingle CreateTestResult(Source, Outcome, Msg)
    
End Sub


Public Sub TestFail(ByVal Source As String, Optional ByVal Msg As String)

    If IsAdhocRun Then
        PrintTestResult TestOutcome.Failed, Source, Msg
    Else
        pTestResults.Add CreateTestResult(Source, TestOutcome.Failed, Msg)
    End If
    
End Sub


' AdhocRun means that test is run from the test method.
Private Function IsAdhocRun() As Boolean
    IsAdhocRun = (pTestResults Is Nothing)
End Function


Public Sub TestFailRunTime(ByVal Source As String, Optional ByVal Msg As String)
    
    If IsAdhocRun Then
        Debug.Print Err.Description, Source
        PrintTestResult TestOutcome.Failed, Source, Msg
    Else
        pTestResults.Add CreateTestResult(Source, TestOutcome.Failed, Err.Description)
    End If
    
End Sub


' Verifies that a string contains a given sub-string.
Public Sub ContainsSubstring(ByVal Value As String, _
                             ByVal Substring As String, _
                             ByVal Source As String, _
                             Optional ByVal Msg As String)

    If StringH.Contains(Value, Substring) Then
        TestPass Source, Msg
    Else
        TestFail Source, "Expected [ " & Value & " ] to be in [" & Substring & "], but it doesn't. " & Msg
    End If
    
End Sub


'Verifies that a collection contains a given object, using an equality comparer.
Public Sub Contains(ByVal Expected As Variant, _
                    ByRef Items As Collection, _
                    ByRef Comparer As IVBAEqualityComparer, _
                    ByVal Source As String, _
                    Optional ByVal Msg As String)
                 
    If CollectionH.Contains(Expected, Items, Comparer) Then
        TestPass Source, Msg
    Else
        TestFail Source, "Expected Value to be in collection, but it wasn't found. " & Msg
    End If
                 
End Sub


'Verifies that a collection contains a given object, using an equality comparer.
Public Sub DoesNotContains(ByVal Expected As Variant, _
                           ByRef Items As Collection, _
                           ByRef Comparer As IVBAEqualityComparer, _
                           ByVal Source As String, _
                           Optional ByVal Msg As String)
                           
    If CollectionH.Contains(Expected, Items, Comparer) Then
        TestFail Source, "Expected Value to not be in collection, but it was found. " & Msg
    Else
        TestPass Source, Msg
    End If
    
End Sub


' Verifies that an object reference is not nothing.
Public Sub IsNotNothing(ByVal Obj As Object, _
                        ByVal Source As String, _
                        Optional ByVal Msg As String)
    
    If Obj Is Nothing Then
        TestFail Source, "Expected Obj to be not nothing, but it is. " & Msg
    Else
        TestPass Source, Msg
    End If

End Sub


' Verifies that an object reference is nothing.
Public Sub IsNothing(ByVal Obj As Object, ByVal Source As String, Optional ByVal Msg As String)
    
    If Obj Is Nothing Then
        TestPass Source, Msg
    Else
        TestFail Source, "Expected Obj to be nothing, but it isn't. " & Msg
    End If
    
End Sub


' Verifies that two objects are the same instance.
Public Sub AreSame(ByVal Expected As Object, _
                   ByVal Actual As Object, _
                   ByVal Source As String, _
                   Optional ByVal Msg As String)
    
    If Expected Is Actual Then
        TestPass Source, Msg
    Else
        TestFail Source, "Expected Value to be the same, but it isn't. " & Msg
    End If

End Sub


' Verifies that two objects are not the same instance.
Public Sub AreNotSame(ByVal Expected As Object, _
                      ByVal Actual As Object, _
                      ByVal Source As String, _
                      Optional ByVal Msg As String)

    If Expected Is Actual Then
        TestFail Source, "Expected Value to be not the same, but it is. " & Msg
    Else
        TestPass Source, Msg
    End If
    
End Sub


' Verifies that expected error code matches actual error code.
Public Sub IsException(ByVal Expected As ExceptionCode, _
                       ByVal Actual As Long, _
                       ByVal Source As String, _
                       Optional ByVal Msg As String)
    
    If Expected = Actual Then
        TestPass Source, Msg
    Else
        TestFail Source, FormatExpectedAndActualExceptions(Expected, Actual) & Msg
    End If

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


' By defualt, method will print test results in the Immediate window.
Public Sub PrintTestResults(Optional ByVal Printer As ITestResultPrinter)
    
    If IsAdhocRun Then
        Exception.InvalidOperationException vbNullString, ModuleName & ".PrintTestResults"
    End If
    
    If Printer Is Nothing Then
        Set Printer = New TestResultImmediatePrinter
    End If
    Printer.PrintMany pTestResults
    
End Sub



