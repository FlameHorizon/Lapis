Attribute VB_Name = "ExUnit"
Option Explicit
'@Folder("Lapis")

Private Const ModuleName As String = "ExUnit"
Private pShowPassingTest As Boolean


Public Property Get ShowPassingTest() As Boolean
    ShowPassingTest = pShowPassingTest
End Property


Public Property Let ShowPassingTest(ByVal Value As Boolean)
    pShowPassingTest = Value
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
    
    If ShowPassingTest = True Then
        PrintTestResult "PASSED", Source, Msg
    End If

End Sub


Private Sub PrintTestResult(ByVal Result As String, _
                            ByVal Source As String, _
                            Optional ByVal Msg As String)
    
    If Result = "PASSED" Then
        Debug.Print Result & "; " & Source
        Exit Sub
    End If
    
    If Msg = vbNullString Then
        Debug.Print Result & "; ; " & Source
    Else
        Debug.Print Result & "; " & Msg & "; " & Source
    End If
    
End Sub


Public Sub TestFail(ByVal Source As String, Optional ByVal Msg As String)
    PrintTestResult "FAILED", Source, Msg
End Sub


Public Sub TestFailRunTime(ByVal Source As String, Optional ByVal Msg As String)
    
    Debug.Print Err.Description, Source
    PrintTestResult "FAILED", Source, Msg
    
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
                           ByVal Source As String, Optional _
                                                  ByVal Msg As String)
                           
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
                       
    AreEqual Expected, Actual, Source, Msg
End Sub


