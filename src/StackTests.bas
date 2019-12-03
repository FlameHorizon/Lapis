Attribute VB_Name = "StackTests"
Option Explicit
'@Folder("Tests")

Private Const ModuleName As String = "StackTests"


Public Sub Start()
    
    InitializationTest
    ThrowArgumentNullExceptionWhenInitialCapacityIsNegativeTest
    PushTest
    PopTest
    ContainsTest
    ClearTest
    PushItemAfterClearTest
    ToArrayTest
    PeekTest
    CopyToTest
    PushPeekPopObjectTest
    
End Sub


Private Sub InitializationTest()
    
    Const MethodName As String = "InitializationTest"
    On Error GoTo ErrHandler
    
    Dim Stack As New Stack: Stack.Init1 10
    ExUnit.AreEqual 0, Stack.Count, GetFullSig(MethodName)

    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetFullSig(MethodName)
    
End Sub


Private Sub ThrowArgumentNullExceptionWhenInitialCapacityIsNegativeTest()
    
    Const MethodName As String = "ThrowArgumentNullExceptionWhenInitialCapacityIsNegativeTest"
    On Error GoTo ErrHandler
    
    Dim Stack As New Stack: Stack.Init1 -1
    
ErrHandler:
    ExUnit.IsException ExceptionCode.ArgumentOutOfRange, Err.Number, GetFullSig(MethodName)
    
End Sub


Private Function GetFullSig(ByVal MethodName As String) As String
    GetFullSig = ModuleName & "." & MethodName
End Function


Private Sub PushTest()

    Const MethodName As String = "PushTest"
    On Error GoTo ErrHandler
    
    Dim Stack As New Stack: Stack.Init1 10
    Stack.Push "a"
    ExUnit.AreEqual "a", Stack.Peek, GetFullSig(MethodName)
    
    Stack.Push "b"
    ExUnit.AreEqual "b", Stack.Peek, GetFullSig(MethodName)
    ExUnit.AreEqual 2, Stack.Count, GetFullSig(MethodName)
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetFullSig(MethodName)

End Sub


Private Sub PopTest()

    Const MethodName As String = "PopTest"
    On Error GoTo ErrHandler

    Dim Stack As New Stack: Stack.Init1 10
    Stack.Push "a"
    Stack.Push "b"

    ExUnit.AreEqual 2, Stack.Count, GetFullSig(MethodName)
    ExUnit.AreEqual "b", Stack.Pop, GetFullSig(MethodName)
    ExUnit.AreEqual 1, Stack.Count, GetFullSig(MethodName)
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetFullSig(MethodName)

End Sub


Private Sub ContainsTest()

    Const MethodName As String = "ContainsTest"
    On Error GoTo ErrHandler
    
    Dim Stack As New Stack: Stack.Init1 10
    Stack.Push 123

    ExUnit.IsTrue Stack.Contains(123, New LongEqualityComparer), GetFullSig(MethodName)
    
    Exit Sub
ErrHandler:
    TestFailRunTime GetFullSig(MethodName)

End Sub


Private Sub PushItemAfterClearTest()

    Const MethodName As String = "PushItemAfterClearTest"
    On Error GoTo ErrHandler

    Dim Stack As New Stack: Stack.Init1 10
    Stack.Push 123
    Stack.Clear
    Stack.Push 123
    
    ExUnit.AreEqual 123, Stack.Peek, GetFullSig(MethodName)

    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetFullSig(MethodName)

End Sub


Private Sub ClearTest()

    Const MethodName As String = "ClearTest"
    On Error GoTo ErrHandler

    Dim Stack As New Stack: Stack.Init1 10
    Stack.Push 123
    
    ExUnit.AreEqual 1, Stack.Count, GetFullSig(MethodName)
    Stack.Clear
    ExUnit.AreEqual 0, Stack.Count, GetFullSig(MethodName)

    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetFullSig(MethodName)

End Sub


Private Sub ToArrayTest()

    Const MethodName As String = "ToArrayTest"
    On Error GoTo ErrHandler

    Dim Stack As New Stack: Stack.Init1 10
    Stack.Push "a"
    Stack.Push "b"
    Stack.Push "c"

    Dim Arr() As Variant
    Arr = Stack.ToArray
     
    ExUnit.AreEqual "c", Arr(0), GetFullSig(MethodName)
    ExUnit.AreEqual "b", Arr(1), GetFullSig(MethodName)
    ExUnit.AreEqual "a", Arr(2), GetFullSig(MethodName)
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetFullSig(MethodName)

End Sub


Private Sub PeekTest()

    Const MethodName As String = "PeekTest"
    On Error GoTo ErrHandler

    Dim Stack As New Stack: Stack.Init1 10
    Stack.Push "a"
    ExUnit.AreEqual "a", Stack.Peek, GetFullSig(MethodName)

    Stack.Push "b"
    ExUnit.AreEqual "b", Stack.Peek, GetFullSig(MethodName)

    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetFullSig(MethodName)

End Sub


Private Sub CopyToTest()

    Const MethodName As String = "CopyToTest"
    On Error GoTo ErrHandler

    Dim Stack As New Stack: Stack.Init1 10
    Stack.Push "a"
    Stack.Push "b"
    Stack.Push "c"

    Dim Arr(2) As Variant
    Stack.CopyTo Arr, 0

    ExUnit.AreEqual "c", Arr(0), GetFullSig(MethodName)
    ExUnit.AreEqual "b", Arr(1), GetFullSig(MethodName)
    ExUnit.AreEqual "a", Arr(2), GetFullSig(MethodName)

    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetFullSig(MethodName)

End Sub


Private Sub PushPeekPopObjectTest()

    Const MethodName As String = "PushPeekPopObjectTest"
    On Error GoTo ErrHandler
    
    Dim Stk As New Stack: Stk.Init1 10

    Stk.Push ThisWorkbook
    ExUnit.AreSame ThisWorkbook, Stk.Peek, GetFullSig(MethodName)
    ExUnit.AreSame ThisWorkbook, Stk.Pop, GetFullSig(MethodName)
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetFullSig(MethodName)

End Sub


